import os
import pandas as pd
import numpy as np

# ==============================================================================
# 1. LEITURA RESILIENTE E EXTRAÇÃO DE DATA REAL (YYYY-MM-01)
# ==============================================================================

def carregar_csv_resiliente(caminho_arquivo, colunas_desejadas=None):
    """
    Leitor de CSV resiliente a falhas comuns de exportação do Excel no Windows:
    - Detecta automaticamente o separador (';', ',', '\t')
    - Tolera codificações variadas ('utf-8-sig', 'utf-8', 'latin1')
    - Remove espaços invisíveis ou caracteres especiais nos nomes das colunas
    - Busca em pastas alternativas se executado no sandbox do Gemini Notebook
    """
    caminho_busca = caminho_arquivo
    if not os.path.exists(caminho_busca):
        alternativo = os.path.join('/workspace/knowledge', caminho_arquivo)
        if os.path.exists(alternativo):
            caminho_busca = alternativo
        else:
            return None

    for encoding in ['utf-8-sig', 'utf-8', 'latin1']:
        for separador in [';', ',', '\t']:
            try:
                df = pd.read_csv(caminho_busca, sep=separador, encoding=encoding, engine='python')
                df.columns = df.columns.astype(str).str.strip().str.lower()
                
                if colunas_desejadas:
                    cols_alvo = [c.lower() for c in colunas_desejadas]
                    cols_encontradas = [c for c in cols_alvo if c in df.columns]
                    if len(cols_encontradas) > 0:
                        return df[cols_encontradas]
                elif len(df.columns) > 1:
                    return df
            except Exception:
                continue

    try:
        df = pd.read_csv(caminho_busca, sep=None, engine='python', encoding='latin1')
        df.columns = df.columns.astype(str).str.strip().str.lower()
        return df
    except Exception:
        return None

def extrair_data_primeiro_dia(valor_data):
    """
    Normalizador de datas (Defeito D3):
    - Converte formatos brasileiros (DD/MM/YYYY) e ISO (YYYY-MM-DD).
    - Descarte automático de datas nulas ou inválidas (D3).
    - Retorna a data no primeiro dia do mês (YYYY-MM-01) para ordenação perfeita no Power BI.
    """
    if pd.isna(valor_data):
        return None
    s = str(valor_data).strip()
    if not s or s in ['0000-00-00', 'nan', 'none', 'null']:
        return None
    
    dt = pd.to_datetime(s, format='%d/%m/%Y', errors='coerce')
    if pd.isna(dt):
        dt = pd.to_datetime(s, format='%Y-%m-%d', errors='coerce')
    if pd.isna(dt):
        dt = pd.to_datetime(s, errors='coerce')
        
    if pd.isna(dt):
        return None
    
    # Retorna o primeiro dia do mês formatado (ex: '2026-03-01')
    return dt.replace(day=1).strftime('%Y-%m-%d')

def normalizar_volume_litros(linha):
    """Converte volume em KG para Litros (1 L = 1,032 kg) (Defeito D5)."""
    try:
        vol = float(linha['volume'])
    except (ValueError, TypeError):
        return 0.0
    
    unidade = str(linha.get('unidade', 'L')).strip().upper()
    if unidade == 'KG':
        return vol / 1.032
    return vol

def normalizar_ccs_milhares(linha):
    """Converte cel_ml para mil_cel_ml (divisão por 1.000) (Defeito D5)."""
    try:
        ccs = float(linha['ccs'])
    except (ValueError, TypeError):
        return np.nan
    
    unidade = str(linha.get('ccs_unidade', 'mil_cel_ml')).strip().lower()
    if unidade in ['cel_ml', 'cel/ml', 'cél_ml', 'células/ml']:
        return ccs / 1000.0
    return ccs

# ==============================================================================
# 2. CARREGAMENTO E BUSCA DINÂMICA DE FAIXAS E TREINAMENTOS
# ==============================================================================

def carregar_faixas_bonificacao(caminho='faixas_bonificacao.csv'):
    df = carregar_csv_resiliente(caminho)
    if df is None:
        raise FileNotFoundError(f"Erro crítico: O arquivo de parâmetros '{caminho}' não foi encontrado.")
    df['ccs_min_mil'] = pd.to_numeric(df['ccs_min_mil'], errors='coerce')
    df['ccs_max_mil'] = pd.to_numeric(df['ccs_max_mil'], errors='coerce')
    df['bonus_reais_por_litro'] = pd.to_numeric(df['bonus_reais_por_litro'], errors='coerce')
    df = df.sort_values('ccs_min_mil').reset_index(drop=True)
    return df

def carregar_efeito_treinamento(caminho='efeito_treinamento.csv'):
    df = carregar_csv_resiliente(caminho)
    if df is None:
        raise FileNotFoundError(f"Erro crítico: O arquivo de parâmetros '{caminho}' não foi encontrado.")
    df['ccs_min_mil'] = pd.to_numeric(df['ccs_min_mil'], errors='coerce')
    df['ccs_max_mil'] = pd.to_numeric(df['ccs_max_mil'], errors='coerce')
    df['reducao_esperada_pct'] = pd.to_numeric(df['reducao_esperada_pct'], errors='coerce')
    df['custo_reais'] = pd.to_numeric(df['custo_reais'], errors='coerce')
    df = df.sort_values('ccs_min_mil').reset_index(drop=True)
    return df

def classificar_faixa_e_bonus_dinamico(ccs, df_faixas):
    if pd.isna(ccs) or ccs < 0:
        return 'Indefinida', 0.0
    
    for idx, row in df_faixas.iterrows():
        c_min = float(row['ccs_min_mil'])
        c_max = float(row['ccs_max_mil'])
        if c_min <= ccs < c_max:
            return str(row['faixa']).strip().upper(), float(row['bonus_reais_por_litro'])
            
    ultima = df_faixas.iloc[-1]
    return str(ultima['faixa']).strip().upper(), float(ultima['bonus_reais_por_litro'])

def obter_meta_ccs_dinamica(faixa_atual, df_faixas):
    """Determina o teto de CCS da faixa superior diretamente da planilha."""
    # Garante que funciona mesmo se o valor de faixa vier formatado como "Faixa B" ou apenas "B"
    f_clean = str(faixa_atual).replace("Faixa ", "").strip().upper()
    match = df_faixas[df_faixas['faixa'].astype(str).str.strip().str.upper() == f_clean]
    if match.empty:
        return None
    
    idx_val = int(match.index[0])
    
    # Se já estiver na faixa de melhor qualidade (índice 0, ex: Faixa A), não há meta superior
    if idx_val == 0:
        return None
    
    # Pega a linha imediatamente superior usando um número inteiro comum
    row_superior = df_faixas.iloc[idx_val - 1]
    return float(row_superior['ccs_max_mil'])

def ccs_pos_treinamento_dinamico(ccs, df_treino):
    if pd.isna(ccs) or ccs < 0:
        return ccs, 0.0
    
    for idx, row in df_treino.iterrows():
        c_min = float(row['ccs_min_mil'])
        c_max = float(row['ccs_max_mil'])
        if c_min <= ccs < c_max:
            red_pct = float(row['reducao_esperada_pct']) / 100.0
            custo = float(row['custo_reais'])
            return ccs * (1.0 - red_pct), custo
            
    ultima = df_treino.iloc[-1]
    red_pct = float(ultima['reducao_esperada_pct']) / 100.0
    custo = float(ultima['custo_reais'])
    return ccs * (1.0 - red_pct), custo

# ==============================================================================
# 3. PIPELINE INTEGRADO E INTERATIVO
# ==============================================================================

def executar_pipeline_squad2(preco_base_leite=2.45):
    print("=" * 80)
    print("🚀  SISTEMA ANALÍTICO DE INTELIGÊNCIA DA BACIA LEITEIRA - SQUAD 2  🚀")
    print("=" * 80)
    print(f"👉 Parâmetro Configurado: Preço Base do Leite = R$ {preco_base_leite:.2f} por litro\n")

    print("[PASSO 1/4] 📂 Ingestão Resiliente e Profiling das Planilhas de Dados...")
    
    # Ingestão de Faixas e Treinamentos
    df_faixas = carregar_faixas_bonificacao('faixas_bonificacao.csv')
    print(f"  └─ ✔️ 'faixas_bonificacao.csv' carregado: {len(df_faixas)} faixas dinâmicas de qualidade mapeadas.")
    df_treino = carregar_efeito_treinamento('efeito_treinamento.csv')
    print(f"  └─ ✔️ 'efeito_treinamento.csv' carregado: {len(df_treino)} faixas de impacto de treinamento mapeadas.")

    # Ingestão de Tabelas Operacionais
    df_fazendas = carregar_csv_resiliente('fazendas.csv', ['fazenda_id', 'tanque_id'])
    if df_fazendas is not None:
        print(f"  └─ ✔️ 'fazendas.csv' carregado: {len(df_fazendas)} fazendas cadastradas no sistema.")
    else:
        raise FileNotFoundError("Erro crítico: Arquivo 'fazendas.csv' não foi encontrado.")

    df_tanques = carregar_csv_resiliente('tanques.csv', ['tanque_id', 'tipo'])
    if df_tanques is not None:
        print(f"  └─ ✔️ 'tanques.csv' carregado: {len(df_tanques)} tanques registrados.")
    else:
        raise FileNotFoundError("Erro crítico: Arquivo 'tanques.csv' não foi encontrado.")

    df_analises = carregar_csv_resiliente('analises.csv', ['origem_id', 'data', 'ccs', 'ccs_unidade'])
    if df_analises is not None:
        print(f"  └─ ✔️ 'analises.csv' carregado: {len(df_analises)} análises laboratoriais de CCS encontradas.")
    else:
        raise FileNotFoundError("Erro crítico: Arquivo 'analises.csv' não foi encontrado.")

    # Tentativa de carregamento de coletas.csv com fallback de simulação automática
    df_coletas = carregar_csv_resiliente('coletas.csv', ['fazenda_id', 'data', 'volume', 'unidade'])
    if df_coletas is not None:
        print(f"  └─ ✔️ 'coletas.csv' carregado: {len(df_coletas)} coletas de leite operacionais registradas.")
    else:
        print("  └─ ⚠️  AVISO: Arquivo 'coletas.csv' não foi localizado no diretório atual.")
        print("     🧪 Iniciando simulação dinâmica de volumes a partir de 'analises.csv' para viabilizar cálculos...")
        
        # Simula volumes estáveis deterministicamente com base na fazenda_id e data de analises
        simulados = []
        analises_validas = df_analises.dropna(subset=['origem_id', 'data'])
        for idx, r in analises_validas.iterrows():
            f_id = str(r['origem_id']).strip().upper()
            hash_val = sum(ord(c) for c in f_id)
            vol_simulado = 1000.0 + (hash_val % 15) * 500.0  # de 1000L a 8000L por coleta
            simulados.append({
                'fazenda_id': f_id,
                'data': r['data'],
                'volume': vol_simulado,
                'unidade': 'L'
            })
        df_coletas = pd.DataFrame(simulados)
        print(f"     ✔️ {len(df_coletas)} registros de coleta gerados automaticamente para teste.")

    print("\n[PASSO 2/4] 🧹 Saneamento, Higienização e Equivalências de Dados (D1 a D5)...")

    # Tratamento Fazendas e Tanques
    df_fazendas['fazenda_id'] = df_fazendas['fazenda_id'].astype(str).str.strip().str.upper()
    df_fazendas['tanque_id'] = df_fazendas['tanque_id'].astype(str).str.strip().str.upper()
    
    # Remoção de Chaves de Teste (D1 & D4)
    linhas_antes_fazendas = len(df_fazendas)
    df_fazendas = df_fazendas[~df_fazendas['fazenda_id'].str.startswith('F9', na=False)]
    linhas_removidas_f = linhas_antes_fazendas - len(df_fazendas)
    if linhas_removidas_f > 0:
        print(f"  └─ 🚫 Filtro D1/D4: {linhas_removidas_f} registros de fazendas de teste (F9XXX) descartados.")

    df_tanques['tanque_id'] = df_tanques['tanque_id'].astype(str).str.strip().str.upper()
    df_tanques['tipo'] = df_tanques['tipo'].astype(str).str.strip().str.lower()
    
    # Isolamento de Tanques Comunitários
    tanques_comunitarios = set(df_tanques[df_tanques['tipo'] == 'comunitario']['tanque_id'])
    df_fazendas = df_fazendas[df_fazendas['tanque_id'].isin(tanques_comunitarios)]
    print(f"  └─ 🥛 Isolamento: Restringindo análise para as {len(df_fazendas)} fazendas associadas aos {len(tanques_comunitarios)} Tanques Comunitários cadastrados.")

    # Tratamento de Análises (D3 & D5)
    df_analises['origem_id'] = df_analises['origem_id'].astype(str).str.strip().str.upper()
    df_analises = df_analises[df_analises['origem_id'].isin(df_fazendas['fazenda_id'])].copy()
    
    df_analises['data_ref'] = df_analises['data'].apply(extrair_data_primeiro_dia)
    linhas_sem_data_a = df_analises['data_ref'].isna().sum()
    df_analises = df_analises.dropna(subset=['data_ref'])
    if linhas_sem_data_a > 0:
        print(f"  └─ 📅 Filtro D3 (Análises): {linhas_sem_data_a} registros com datas inválidas ou nulas descartados.")
        
    df_analises['ccs_mil'] = df_analises.apply(normalizar_ccs_milhares, axis=1)
    df_analises = df_analises.dropna(subset=['ccs_mil'])
    print(f"  └─ 🔬 Conversão D5 (CCS): Normalizado e convertido todos os valores para milhares de cél/ml (mil_cel_ml).")

    # Tratamento de Coletas (D3 & D5)
    df_coletas['fazenda_id'] = df_coletas['fazenda_id'].astype(str).str.strip().str.upper()
    df_coletas = df_coletas[df_coletas['fazenda_id'].isin(df_fazendas['fazenda_id'])].copy()
    
    df_coletas['data_ref'] = df_coletas['data'].apply(extrair_data_primeiro_dia)
    linhas_sem_data_c = df_coletas['data_ref'].isna().sum()
    df_coletas = df_coletas.dropna(subset=['data_ref'])
    if linhas_sem_data_c > 0:
        print(f"  └─ 📅 Filtro D3 (Coletas): {linhas_sem_data_c} registros com datas inválidas descartados.")
        
    df_coletas['volume_litros'] = df_coletas.apply(normalizar_volume_litros, axis=1)
    print(f"  └─ ⚖️  Conversão D5 (Volumes): Todos os volumes em KG convertidos para Litros usando a densidade oficial de 1,032 kg/L.")

    # Agrupa volumes mensais com base na data_ref
    vol_mensal = df_coletas.groupby(['fazenda_id', 'data_ref'])['volume_litros'].sum().reset_index()

    resumo_tanques = []
    simulacoes_leave_one_out = []
    plano_treinamento = []

    lista_tanques = sorted(list(set(df_fazendas['tanque_id'])))
    lista_periodos = sorted(list(set(df_analises['data_ref'])))

    print("\n[PASSO 3/4] 🧠 Execução dos Algoritmos de Negócio (Níveis 1, 2 e 3)...")
    print("  ⚙️ Processando tanques comunitários ao longo dos períodos mensais...")

    # Variáveis de Consolidação para a saída limpa
    total_combinacoes = 0
    tanques_processados = set()
    periodos_processados = set()
    treinamentos_viaveis = 0
    tanques_inviaveis = 0
    total_investimento_capacitacao = 0.0
    total_retorno_grupo = 0.0

    for t_id in lista_tanques:
        fazendas_do_tanque = df_fazendas[df_fazendas['tanque_id'] == t_id]['fazenda_id'].unique()
        
        for p_ref in lista_periodos:
            analises_m = df_analises[(df_analises['origem_id'].isin(fazendas_do_tanque)) & (df_analises['data_ref'] == p_ref)]
            coletas_m = vol_mensal[(vol_mensal['fazenda_id'].isin(fazendas_do_tanque)) & (vol_mensal['data_ref'] == p_ref)]
            
            if analises_m.empty or coletas_m.empty:
                continue
            
            ccs_dict = analises_m.groupby('origem_id')['ccs_mil'].mean().to_dict()
            vol_dict = coletas_m.set_index('fazenda_id')['volume_litros'].to_dict()
            
            fazendas_ativas = [f for f in fazendas_do_tanque if f in ccs_dict and f in vol_dict and vol_dict[f] > 0]
            if not fazendas_ativas:
                continue

            # NÍVEL 1: Média Ponderada
            vol_total = sum(vol_dict[f] for f in fazendas_ativas)
            ccs_ponderado = sum(ccs_dict[f] * vol_dict[f] for f in fazendas_ativas) / vol_total
            
            faixa_letra, bonus_atual = classificar_faixa_e_bonus_dinamico(ccs_ponderado, df_faixas)
            faixa_atual_formatada = f"Faixa {faixa_letra}"
            faturamento_bruto = vol_total * (preco_base_leite + bonus_atual)

            # Define a CHAVE COMPOSTA EXCLUSIVA (Evita duplicados no Power BI)
            chave_composta = f"{t_id}_{p_ref}"

            # NÍVEL 3: Otimização de Treinamento
            meta_ccs = obter_meta_ccs_dinamica(faixa_atual_formatada, df_faixas)
            
            selecionadas = []
            custo_acumulado = 0.0
            ccs_sim = ccs_ponderado
            atingiu_meta = False
            nova_faixa_estimada_formatada = faixa_atual_formatada
            novo_bonus = bonus_atual
            ganho_mensal_grupo = 0.0

            if meta_ccs is not None:
                impactos = []
                for f in fazendas_ativas:
                    c_orig = ccs_dict[f]
                    c_trein, custo_f = ccs_pos_treinamento_dinamico(c_orig, df_treino)
                    reducao_tanque = (c_orig - c_trein) * (vol_dict[f] / vol_total)
                    if reducao_tanque > 0:
                        impactos.append((f, reducao_tanque, custo_f))
                
                # Ordenação gulosa pelo maior impacto marginal
                impactos.sort(key=lambda x: x[1], reverse=True)
                
                for f, red, custo_f in impactos:
                    if ccs_sim < meta_ccs: break
                    selecionadas.append(f)
                    custo_acumulado += custo_f
                    ccs_sim -= red
                
                nova_faixa_t_letra, novo_bonus = classificar_faixa_e_bonus_dinamico(ccs_sim, df_faixas)
                nova_faixa_estimada_formatada = f"Faixa {nova_faixa_t_letra}"
                atingiu_meta = ccs_sim < meta_ccs
                
                # CORREÇÃO CRÍTICA DO SINAL DE GESTÃO FINANCEIRA
                if not atingiu_meta:
                    selecionadas = []
                    custo_acumulado = 0.0
                    ganho_mensal_grupo = 0.0
                else:
                    ganho_mensal_grupo = vol_total * (novo_bonus - bonus_atual)

            # Salva o Plano de Treinamento Nível 3
            plano_treinamento.append({
                'chave_tanque_data': chave_composta,
                'tanque_id': t_id,
                'data_ref': p_ref,
                'faixa_inicial': faixa_atual_formatada,
                'ccs_inicial': round(ccs_ponderado, 2),
                'meta_ccs': meta_ccs if meta_ccs is not None else 0.0,
                'possivel_subir_faixa': atingiu_meta,
                'qtd_fazendas_treinar': len(selecionadas),
                'fazendas_sugeridas': ", ".join(selecionadas) if selecionadas else "Nenhuma",
                'novo_ccs_estimado': round(ccs_sim, 2),
                'nova_faixa_estimada': nova_faixa_estimada_formatada,
                'custo_treinamento_reais': custo_acumulado,
                'ganho_mensal_adicional_grupo_reais': round(ganho_mensal_grupo, 2)
            })

            # NÍVEL 2: Simulação Leave-One-Out
            qtd_alteram_faixa = 0
            for f_out in fazendas_ativas:
                vol_rem = vol_total - vol_dict[f_out]
                if vol_rem <= 0: continue
                
                ccs_rem = sum(ccs_dict[f] * vol_dict[f] for f in fazendas_ativas if f != f_out) / vol_rem
                faixa_nova_letra, bonus_novo = classificar_faixa_e_bonus_dinamico(ccs_rem, df_faixas)
                
                mudou = (faixa_nova_letra != faixa_letra)
                faixa_sem_fazenda_formatada = f"Faixa {faixa_nova_letra}"
                
                if mudou:
                    qtd_alteram_faixa += 1
                    
                ganho_remanescentes = vol_rem * (bonus_novo - bonus_atual)

                # Define o status dinâmico do produtor para Storytelling (Nível 2)
                if mudou and (bonus_novo > bonus_atual):
                    status = "🔴 Derruba a Nota"
                elif ccs_dict[f_out] < 400.0:
                    status = "🟢 Protege o Grupo"
                else:
                    status = "⚪ Sem Alteração"

                # Lógica de Treinamento Individual (Nível 3)
                if f_out in selecionadas:
                    status_treino = "☑️ Treinar (R$ 850)"
                    custo_treino_ind = 850.0
                    retorno_treino_ind = ganho_mensal_grupo if atingiu_meta else 0.0
                else:
                    status_treino = "⚪ Manter Monitoramento"
                    custo_treino_ind = 0.0
                    retorno_treino_ind = 0.0

                c_orig = ccs_dict[f_out]
                c_trein, _ = ccs_pos_treinamento_dinamico(c_orig, df_treino)

                simulacoes_leave_one_out.append({
                    'chave_tanque_data': chave_composta,
                    'tanque_id': t_id,
                    'data_ref': p_ref,
                    'fazenda_id': f_out,
                    'volume_entregue_litros': round(vol_dict[f_out], 2),
                    'participacao_volume_pct': round(vol_dict[f_out] / vol_total, 4),
                    
                    # Nível 2
                    'ccs_individual_fazenda': round(ccs_dict[f_out], 2),
                    'ccs_tanque_sem_fazenda': round(ccs_rem, 2), 
                    'faixa_atual_tanque': faixa_atual_formatada,
                    'faixa_sem_fazenda': faixa_sem_fazenda_formatada,
                    'altera_faixa_grupo': mudou,
                    'status_produtor': status,
                    'delta_bonus_por_litro': round(bonus_novo - bonus_atual, 2),
                    'ganho_financeiro_remanescentes_reais': round(ganho_remanescentes, 2),
                    
                    # Nível 3
                    'status_treinamento_fazenda': status_treino,
                    'ccs_pos_treinamento_fazenda': round(c_trein, 2),
                    'custo_treinamento_individual_reais': custo_treino_ind,
                    'retorno_treinamento_grupo_reais': round(retorno_treino_ind, 2),
                    
                    # Backward compatibility
                    'ganho_outros_reais': round(ganho_remanescentes, 2),
                    'delta_receita_remanescentes_reais': round(ganho_remanescentes, 2)
                })

            resumo_tanques.append({
                'chave_tanque_data': chave_composta,
                'tanque_id': t_id,
                'data_ref': p_ref,
                'qtd_fazendas_ativas': len(fazendas_ativas),
                'volume_total_litros': round(vol_total, 2),
                'ccs_ponderado_tanque': round(ccs_ponderado, 2),
                'faixa_atual': faixa_atual_formatada,
                'bonus_por_litro_reais': bonus_atual,
                'faturamento_total_bruto_reais': round(faturamento_bruto, 2),
                'qtd_fazendas_que_alteram_faixa': qtd_alteram_faixa
            })

            # Estatísticas Consolidadas
            total_combinacoes += 1
            tanques_processados.add(t_id)
            periodos_processados.add(p_ref)
            if meta_ccs is not None:
                if atingiu_meta:
                    treinamentos_viaveis += 1
                    total_investimento_capacitacao += custo_acumulado
                    total_retorno_grupo += ganho_mensal_grupo
                else:
                    tanques_inviaveis += 1

    # Criação dos DataFrames para Exportação (PASSO CRÍTICO FORA DO LOOP!)
    df_resumo = pd.DataFrame(resumo_tanques)
    df_simulacoes = pd.DataFrame(simulacoes_leave_one_out)
    df_plano = pd.DataFrame(plano_treinamento)

    # Saída Consolidada Elegante e Limpa (Sem spammar o console!)
    print(f"  └─ ✔️ Sucesso: {len(tanques_processados)} tanques comunitários processados em {len(periodos_processados)} períodos temporais.")
    print(f"  └─ 📊 Métricas Consolidadas do Pipeline:")
    print(f"     ├─ Total de Filtros Tanque-Mês Processados: {total_combinacoes}")
    print(f"     ├─ Planos de Otimização Viáveis (Nível 3): {treinamentos_viaveis} períodos")
    print(f"     ├─ Casos Limite Inviáveis (Nível 3): {tanques_inviaveis} períodos")
    if treinamentos_viaveis > 0:
        roi_medio = (total_retorno_grupo / total_investimento_capacitacao) * 100 if total_investimento_capacitacao > 0 else 0
        print(f"     ├─ Investimento Total em Capacitação Mapeado: R$ {total_investimento_capacitacao:,.2f}")
        print(f"     ├─ Retorno Financeiro Adicional do Grupo: R$ {total_retorno_grupo:,.2f}")
        print(f"     └─ ROI Médio Global Calculado: {roi_medio:.1f}%")

    print("\n[PASSO 4/4] 💾 Exportação e Consolidação das Tabelas de Saída...")

    # Exportação segura para o Power BI
    try:
        df_resumo.to_csv('squad2_tanques_resumo_mensal.csv', sep=';', index=False, encoding='utf-8-sig')
        print(f"  └─ ✔️ 'squad2_tanques_resumo_mensal.csv' exportado com sucesso! ({len(df_resumo)} registros)")
        
        df_simulacoes.to_csv('squad2_simulacao_leave_one_out.csv', sep=';', index=False, encoding='utf-8-sig')
        print(f"  └─ ✔️ 'squad2_simulacao_leave_one_out.csv' exportado com sucesso! ({len(df_simulacoes)} registros)")
        
        df_plano.to_csv('squad2_plano_treinamento.csv', sep=';', index=False, encoding='utf-8-sig')
        print(f"  └─ ✔️ 'squad2_plano_treinamento.csv' exportado com sucesso! ({len(df_plano)} registros)")
        
        print("\n" + "=" * 80)
        print("🎉 PROCESSAMENTO CONCLUÍDO COM SUCESSO! AS TABELAS DO POWER BI FORAM ATUALIZADAS!")
        print("=" * 80)
    except PermissionError:
        print("\n❌ ERRO DE PERMISSÃO: Feche os arquivos CSV no Excel ou Power BI e tente novamente!")

if __name__ == "__main__":
    executar_pipeline_squad2()
