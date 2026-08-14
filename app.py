import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import requests
from datetime import datetime
from functools import lru_cache

st.set_page_config(page_title="Gerador de Contratos", layout="wide")
st.title("🚛 Gerador de Contratos de Transporte")

# ------------------------------------------------------------
# FUNÇÕES AUXILIARES (com cache para agilizar)
# ------------------------------------------------------------
@lru_cache(maxsize=128)
def limpar_cpf_cnpj(val):
    if pd.isna(val): return ''
    return ''.join(filter(str.isdigit, str(val)))

@lru_cache(maxsize=128)
def formatar_cpf(cpf):
    cpf = limpar_cpf_cnpj(cpf)
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    return cpf

@lru_cache(maxsize=128)
def formatar_cnpj(cnpj):
    cnpj = limpar_cpf_cnpj(cnpj)
    if len(cnpj) == 14:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return cnpj

def formatar_cpf_cnpj(val):
    val_limpo = limpar_cpf_cnpj(val)
    if len(val_limpo) == 11:
        return formatar_cpf(val_limpo)
    elif len(val_limpo) == 14:
        return formatar_cnpj(val_limpo)
    return val

def limpar_placa(p):
    if pd.isna(p): return ''
    return str(p).strip().upper().replace('-', '').replace(' ', '')

def safe_str(val):
    if pd.isna(val): return ''
    return str(val).strip()

def normalizar_nome(nome):
    return ' '.join(str(nome).upper().split())

meses = {1:'janeiro',2:'fevereiro',3:'março',4:'abril',5:'maio',6:'junho',
         7:'julho',8:'agosto',9:'setembro',10:'outubro',11:'novembro',12:'dezembro'}

def formatar_data_cadastro(data_str):
    """Converte data com fallback robusto."""
    if pd.isna(data_str) or not str(data_str).strip():
        hoje = datetime.now()
        return str(hoje.day), meses[hoje.month], str(hoje.year)
    try:
        data_str = str(data_str).strip()
        # Tentar dd/mm/aaaa
        if '/' in data_str:
            partes = data_str.split('/')
            if len(partes) == 3:
                dt = datetime(int(partes[2]), int(partes[1]), int(partes[0]))
                return str(dt.day), meses[dt.month], str(dt.year)
        # Tentar aaaa-mm-dd
        if '-' in data_str:
            partes = data_str.split('-')
            if len(partes) == 3 and len(partes[0]) == 4:
                dt = datetime(int(partes[0]), int(partes[1]), int(partes[2]))
                return str(dt.day), meses[dt.month], str(dt.year)
        # Fallback para pandas
        dt = pd.to_datetime(data_str)
        return str(dt.day), meses[dt.month], str(dt.year)
    except:
        hoje = datetime.now()
        return str(hoje.day), meses[hoje.month], str(hoje.year)

def extrair_endereco_pj(texto_endereco):
    """Separa rua, número e complemento, removendo prefixos comuns."""
    if pd.isna(texto_endereco) or not str(texto_endereco).strip():
        return '', '', ''
    texto = str(texto_endereco).strip()
    numero = ''
    complemento = ''
    match_num = re.search(r',\s*(\d+)\s*[-/]?\s*([^,]*)$', texto)
    if match_num:
        numero = match_num.group(1)
        complemento = match_num.group(2).strip()
        texto = texto[:match_num.start()].strip()
    else:
        match_num = re.search(r'\b(\d+)\b', texto)
        if match_num:
            numero = match_num.group(1)
    rua = texto.strip()
    # Remover prefixos comuns
    prefixos = ['RUA ', 'R. ', 'AVENIDA ', 'AV. ', 'ALAMEDA ', 'TRAVESSA ', 'TRAV. ', 'AV ', 'R ']
    for prefixo in prefixos:
        if rua.upper().startswith(prefixo):
            rua = rua[len(prefixo):].strip()
            break
    return rua, numero, complemento

def buscar_cep_online(rua, cidade, uf):
    """Busca CEP via ViaCEP (com timeout) e retorna ou string vazia."""
    if not rua or not cidade or not uf:
        return ''
    termos = re.findall(r'\b[A-ZÀ-Ú]{3,}\b', rua.upper())
    if not termos:
        return ''
    query = ' '.join(termos[:3])
    url = f"https://viacep.com.br/ws/{uf}/{cidade}/{query}/json/"
    try:
        resp = requests.get(url, timeout=3)
        if resp.status_code == 200:
            dados = resp.json()
            if isinstance(dados, list) and dados:
                return dados[0].get('cep', '')
            elif isinstance(dados, dict) and dados.get('cep'):
                return dados['cep']
    except Exception:
        pass
    return ''

def buscar_rg_proprietario(cpf_prop, df_cond, df_forn):
    cpf_limpo = limpar_cpf_cnpj(cpf_prop)
    if not cpf_limpo:
        return ''
    m = df_cond[df_cond['CPF N°'].apply(limpar_cpf_cnpj) == cpf_limpo]
    if not m.empty:
        rg = safe_str(m.iloc[0].get('RG N°', ''))
        if rg:
            return rg
    if 'RG' in df_forn.columns:
        f = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cpf_limpo]
        if not f.empty:
            rg = safe_str(f.iloc[0].get('RG', ''))
            if rg:
                return rg
    return ''

# ------------------------------------------------------------
# FUNÇÕES DE NEGÓCIO (otimizadas)
# ------------------------------------------------------------
COLUNAS_MOTORISTA = ['Motorista Cnpj', 'Motorista CPF', 'Condutor CPF', 'Motorista', 'CPF Motorista', 'Condutor']

def buscar_veiculo_por_placa(df_veic_antt, termo, df_forn):
    if not termo:
        return []
    mask = df_veic_antt['Placa'].str.upper().str.contains(termo.upper(), na=False)
    resultados = df_veic_antt[mask]
    veiculos = []
    # Pré-cria um dicionário de CPF->nome para acelerar a busca
    # Usando um cache local para evitar percorrer df_forn várias vezes
    if not df_forn.empty:
        cpf_to_name = {limpar_cpf_cnpj(row['Cnpj/Cpf']): safe_str(row['Nome Fornecedor']) 
                       for _, row in df_forn.iterrows()}
    else:
        cpf_to_name = {}
    for _, row in resultados.iterrows():
        placa = safe_str(row['Placa'])
        marca = safe_str(row.get('Marca', ''))
        modelo = safe_str(row.get('Modelo', ''))
        ano = safe_str(row.get('Ano', ''))
        prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
        prop_nome = cpf_to_name.get(prop_cnpj, '')
        veiculos.append({
            'placa': placa,
            'marca': marca,
            'modelo': modelo,
            'ano': ano,
            'prop_cnpj': prop_cnpj,
            'prop_nome': prop_nome,
            'row': row
        })
    return veiculos

def obter_veiculos_do_proprietario(cnpj_prop, df_veic_antt, df_veic_compl, df_cond, df_forn):
    cnpj_limpo = limpar_cpf_cnpj(cnpj_prop)
    if not cnpj_limpo:
        return []
    mask = df_veic_antt['Proprietário Cnpj'].apply(limpar_cpf_cnpj) == cnpj_limpo
    veiculos = []
    # Pré-cria dicionários para acelerar buscas de motorista e fornecedor
    if not df_cond.empty:
        cpf_cond_to_nome = {limpar_cpf_cnpj(row['CPF N°']): safe_str(row.get('Nome', '')) for _, row in df_cond.iterrows()}
        cpf_cond_to_cnh = {limpar_cpf_cnpj(row['CPF N°']): safe_str(row.get('CNH N°', '')) for _, row in df_cond.iterrows()}
        cpf_cond_to_rg = {limpar_cpf_cnpj(row['CPF N°']): safe_str(row.get('RG N°', '')) for _, row in df_cond.iterrows()}
    else:
        cpf_cond_to_nome = {}
        cpf_cond_to_cnh = {}
        cpf_cond_to_rg = {}
    if not df_forn.empty:
        cpf_forn_to_nome = {limpar_cpf_cnpj(row['Cnpj/Cpf']): safe_str(row.get('Nome Fornecedor', '')) for _, row in df_forn.iterrows()}
    else:
        cpf_forn_to_nome = {}

    for _, row in df_veic_antt[mask].iterrows():
        placa = safe_str(row['Placa'])
        placa_limpa = limpar_placa(placa)
        marca = safe_str(row.get('Marca', ''))
        modelo = safe_str(row.get('Modelo', ''))
        ano = safe_str(row.get('Ano', ''))
        renavan = safe_str(row.get('Renavan', ''))
        n_equip = ''
        if not df_veic_compl.empty:
            vc = df_veic_compl[df_veic_compl['Placa'].apply(limpar_placa) == placa_limpa]
            if not vc.empty:
                rv = vc.iloc[0]
                marca = safe_str(rv.get('Marca', '')) or marca
                modelo = safe_str(rv.get('Modelo', '')) or modelo
                n_equip = safe_str(rv.get('Nº Equipamento', ''))
        mot_cpf = ''
        for col in COLUNAS_MOTORISTA:
            if col in row:
                val = row.get(col, '')
                if val and str(val).strip():
                    mot_cpf = limpar_cpf_cnpj(val)
                    if mot_cpf:
                        break
        if not mot_cpf and not df_veic_compl.empty:
            vc = df_veic_compl[df_veic_compl['Placa'].apply(limpar_placa) == placa_limpa]
            if not vc.empty:
                for col in COLUNAS_MOTORISTA:
                    if col in vc.columns:
                        val = vc.iloc[0].get(col, '')
                        if val and str(val).strip():
                            mot_cpf = limpar_cpf_cnpj(val)
                            if mot_cpf:
                                break
        mot_nome = ''
        if mot_cpf:
            mot_nome = cpf_cond_to_nome.get(mot_cpf) or cpf_forn_to_nome.get(mot_cpf, '')
        veiculos.append({
            'placa': placa,
            'marca': marca,
            'modelo': modelo,
            'ano': ano,
            'renavan': renavan,
            'n_equipamento': n_equip,
            'mot_cpf': mot_cpf,
            'mot_nome': mot_nome
        })
    return veiculos

def obter_motoristas_dos_veiculos(veiculos, df_condutores, df_fornecedores):
    cpfs_vistos = set()
    motoristas = []
    # Pré-cria dicionários
    if not df_condutores.empty:
        cpf_cond_to_data = {
            limpar_cpf_cnpj(row['CPF N°']): {
                'nome': safe_str(row.get('Nome', '')),
                'cnh': safe_str(row.get('CNH N°', '')),
                'rg': safe_str(row.get('RG N°', ''))
            }
            for _, row in df_condutores.iterrows()
        }
    else:
        cpf_cond_to_data = {}
    if not df_fornecedores.empty:
        cpf_forn_to_nome = {limpar_cpf_cnpj(row['Cnpj/Cpf']): safe_str(row.get('Nome Fornecedor', '')) for _, row in df_fornecedores.iterrows()}
    else:
        cpf_forn_to_nome = {}

    for v in veiculos:
        cpf = v.get('mot_cpf')
        if cpf and cpf not in cpfs_vistos:
            cpfs_vistos.add(cpf)
            nome = v.get('mot_nome', '')
            cnh = ''
            rg = ''
            if cpf in cpf_cond_to_data:
                data = cpf_cond_to_data[cpf]
                nome = data['nome'] or nome
                cnh = data['cnh']
                rg = data['rg']
            else:
                # Tenta no fornecedor PF
                if cpf in cpf_forn_to_nome:
                    nome = cpf_forn_to_nome[cpf] or nome
            motoristas.append({
                'nome': nome,
                'cpf': cpf,
                'cnh': cnh,
                'rg': rg
            })
    return motoristas

def obter_todos_condutores_da_base(df_condutores, df_fornecedores):
    condutores = []
    cpfs_vistos = set()
    for _, row in df_condutores.iterrows():
        nome = safe_str(row.get('Nome', ''))
        cpf = safe_str(row.get('CPF N°', ''))
        if nome and cpf:
            cpf_limpo = limpar_cpf_cnpj(cpf)
            if cpf_limpo not in cpfs_vistos:
                condutores.append({
                    'nome': nome,
                    'cpf': cpf_limpo,
                    'cnh': safe_str(row.get('CNH N°', '')),
                    'rg': safe_str(row.get('RG N°', ''))
                })
                cpfs_vistos.add(cpf_limpo)
    mask_pf = df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj).str.len() == 11
    for _, row in df_fornecedores[mask_pf].iterrows():
        nome = safe_str(row.get('Nome Fornecedor', ''))
        cpf = safe_str(row.get('Cnpj/Cpf', ''))
        if nome and cpf:
            cpf_limpo = limpar_cpf_cnpj(cpf)
            if cpf_limpo not in cpfs_vistos:
                condutores.append({
                    'nome': nome,
                    'cpf': cpf_limpo,
                    'cnh': '',
                    'rg': ''
                })
                cpfs_vistos.add(cpf_limpo)
    return condutores

def buscar_antt(proprietario, veiculos_selecionados, df_veiculos_antt, df_veiculos_compl):
    antt = proprietario.get('rntrc', '')
    if antt and antt.lower() != 'nan':
        return antt
    for v in veiculos_selecionados:
        placa_limpa = limpar_placa(v['placa'])
        if not df_veiculos_compl.empty:
            vc = df_veiculos_compl[df_veiculos_compl['Placa'].apply(limpar_placa) == placa_limpa]
            if not vc.empty:
                antt_tmp = safe_str(vc.iloc[0].get('ANTT/RNTRC nº', ''))
                if antt_tmp and antt_tmp.lower() != 'nan':
                    return antt_tmp
    if 'Rntrc' in df_veiculos_antt.columns:
        for v in veiculos_selecionados:
            placa_limpa = limpar_placa(v['placa'])
            va = df_veiculos_antt[df_veiculos_antt['Placa'].apply(limpar_placa) == placa_limpa]
            if not va.empty:
                antt_tmp = safe_str(va.iloc[0].get('Rntrc', ''))
                if antt_tmp and antt_tmp.lower() != 'nan':
                    return antt_tmp
    return ''

def buscar_serial(veiculos_selecionados):
    for v in veiculos_selecionados:
        if v.get('n_equipamento') and v['n_equipamento'].lower() != 'nan':
            return v['n_equipamento']
    return ''

# ------------------------------------------------------------
# INTERFACE STREAMLIT
# ------------------------------------------------------------
st.sidebar.header("📤 Upload das planilhas")
fornecedores_file = st.sidebar.file_uploader("Fornecedores", type="xlsx")
veiculos_antt_file = st.sidebar.file_uploader("Veículos ANTT", type="xlsx")
condutores_file = st.sidebar.file_uploader("Condutores", type="xlsx")
veiculos_compl_file = st.sidebar.file_uploader("Veículos Complementares", type="xlsx")
template_file = st.sidebar.file_uploader("Template do Contrato (DOCX)", type="docx")

if not (fornecedores_file and veiculos_antt_file and condutores_file and veiculos_compl_file and template_file):
    st.info("📂 Faça o upload de todos os arquivos necessários na barra lateral.")
    st.stop()

# Ler e carregar DataFrames com otimização (engine='openpyxl' já é padrão)
try:
    with st.spinner("Carregando planilhas..."):
        df_fornecedores = pd.read_excel(fornecedores_file, dtype=str, engine='openpyxl', header=4)
        df_veiculos_antt = pd.read_excel(veiculos_antt_file, dtype=str, engine='openpyxl', header=4)
        df_condutores = pd.read_excel(condutores_file, dtype=str, engine='openpyxl', header=4)
        df_veiculos_compl = pd.read_excel(veiculos_compl_file, dtype=str, engine='openpyxl', header=4)
except Exception as e:
    st.error(f"❌ Erro ao ler planilhas: {e}")
    st.stop()

for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]:
    df.dropna(how='all', inplace=True)
    df.columns = df.columns.str.strip()

# ============================================================
# FLUXO PRINCIPAL (com estado da sessão)
# ============================================================
if "etapa" not in st.session_state:
    st.session_state.etapa = "busca_veiculo"
if "veiculo_escolhido" not in st.session_state:
    st.session_state.veiculo_escolhido = None
if "proprietario" not in st.session_state:
    st.session_state.proprietario = None
if "todos_veiculos" not in st.session_state:
    st.session_state.todos_veiculos = []
if "veiculos_selecionados" not in st.session_state:
    st.session_state.veiculos_selecionados = []
if "motoristas_selecionados" not in st.session_state:
    st.session_state.motoristas_selecionados = []

# ------------------------------------------------------------
# ETAPA 1: BUSCAR VEÍCULO POR PLACA
# ------------------------------------------------------------
st.header("🔍 1. Buscar Veículo")

placa_busca = st.text_input("Digite a placa (ou parte):", key="placa_busca")
if placa_busca:
    with st.spinner("Buscando veículos..."):
        resultados = buscar_veiculo_por_placa(df_veiculos_antt, placa_busca, df_fornecedores)
    if resultados:
        opcoes = []
        for v in resultados:
            opcoes.append(f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']}) - Proprietário: {v['prop_nome'] or 'N/I'}")
        selecao = st.selectbox("Selecione o veículo principal:", opcoes, key="selecao_veiculo")
        if st.button("✅ Confirmar veículo"):
            idx = opcoes.index(selecao)
            st.session_state.veiculo_escolhido = resultados[idx]
            st.session_state.etapa = "confirmar_proprietario"
            st.rerun()
    else:
        st.warning("Nenhum veículo encontrado.")

# ------------------------------------------------------------
# ETAPA 2: CONFIRMAR PROPRIETÁRIO
# ------------------------------------------------------------
if st.session_state.etapa == "confirmar_proprietario":
    veic = st.session_state.veiculo_escolhido
    st.header("👤 2. Confirmar Proprietário")

    cnpj_prop = veic['prop_cnpj']
    prop_nome_sugerido = veic['prop_nome']
    proprietario = None
    if cnpj_prop:
        p = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cnpj_prop]
        if not p.empty:
            row = p.iloc[0]
            proprietario = {
                'nome': safe_str(row['Nome Fornecedor']),
                'cpf_cnpj': safe_str(row['Cnpj/Cpf']),
                'endereco': safe_str(row.get('Endereço', '')),
                'bairro': safe_str(row.get('Bairro', '')),
                'cidade': safe_str(row.get('Cidade', '')),
                'uf': safe_str(row.get('UF', '')),
                'data_inclusao': safe_str(row.get('Data Inclusão', '')),
                'rntrc': safe_str(row.get('Rntrc', ''))
            }

    if proprietario:
        st.write(f"**Nome:** {proprietario['nome']}")
        st.write(f"**CPF/CNPJ:** {formatar_cpf_cnpj(proprietario['cpf_cnpj'])}")
        st.write(f"**Endereço:** {proprietario['endereco']}, {proprietario['bairro']}, {proprietario['cidade']}/{proprietario['uf']}")
        if st.button("✅ Confirmar Proprietário"):
            st.session_state.proprietario = proprietario
            st.session_state.etapa = "selecionar_veiculos"
            with st.spinner("Carregando veículos do proprietário..."):
                st.session_state.todos_veiculos = obter_veiculos_do_proprietario(
                    cnpj_prop, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
                )
            if not st.session_state.todos_veiculos:
                st.session_state.todos_veiculos = [{
                    'placa': veic['placa'],
                    'marca': veic['marca'],
                    'modelo': veic['modelo'],
                    'ano': veic['ano'],
                    'renavan': '',
                    'n_equipamento': '',
                    'mot_cpf': '',
                    'mot_nome': ''
                }]
            st.rerun()
    else:
        st.warning("Proprietário não encontrado na base. Preencha manualmente:")
        with st.form("manual_proprietario"):
            nome = st.text_input("Nome:", value=prop_nome_sugerido or '')
            cpf = st.text_input("CPF/CNPJ:")
            endereco = st.text_input("Endereço:")
            bairro = st.text_input("Bairro:")
            cidade = st.text_input("Cidade:")
            uf = st.text_input("UF:")
            data_inclusao = st.text_input("Data de inclusão:")
            rntrc = st.text_input("RNTRC:")
            if st.form_submit_button("✅ Confirmar"):
                st.session_state.proprietario = {
                    'nome': nome,
                    'cpf_cnpj': cpf,
                    'endereco': endereco,
                    'bairro': bairro,
                    'cidade': cidade,
                    'uf': uf,
                    'data_inclusao': data_inclusao,
                    'rntrc': rntrc
                }
                st.session_state.etapa = "selecionar_veiculos"
                cnpj_limpo = limpar_cpf_cnpj(cpf)
                if cnpj_limpo:
                    with st.spinner("Carregando veículos..."):
                        st.session_state.todos_veiculos = obter_veiculos_do_proprietario(
                            cnpj_limpo, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
                        )
                if not st.session_state.todos_veiculos:
                    st.session_state.todos_veiculos = [{
                        'placa': veic['placa'],
                        'marca': veic['marca'],
                        'modelo': veic['modelo'],
                        'ano': veic['ano'],
                        'renavan': '',
                        'n_equipamento': '',
                        'mot_cpf': '',
                        'mot_nome': ''
                    }]
                st.rerun()

# ------------------------------------------------------------
# ETAPA 3: SELECIONAR VEÍCULOS
# ------------------------------------------------------------
if st.session_state.etapa == "selecionar_veiculos":
    st.header("🚛 3. Selecionar Veículos para o Contrato")

    if st.session_state.todos_veiculos:
        opcoes_veiculos = [f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})" for v in st.session_state.todos_veiculos]
        selecionados = st.multiselect(
            "Veículos selecionados (o primeiro será o principal):",
            opcoes_veiculos,
            default=opcoes_veiculos[:1]
        )
        if selecionados:
            st.session_state.veiculos_selecionados = [st.session_state.todos_veiculos[opcoes_veiculos.index(s)] for s in selecionados]
    else:
        st.warning("Nenhum veículo disponível.")

    with st.expander("➕ Adicionar veículo extra da base completa"):
        modo_busca = st.radio("Como buscar?", ["Por placa", "Listar todos (filtrar)"], key="modo_busca_veic")
        if modo_busca == "Por placa":
            placa_extra = st.text_input("Digite a placa (ou parte):", key="placa_extra")
            if placa_extra:
                with st.spinner("Buscando..."):
                    extras = buscar_veiculo_por_placa(df_veiculos_antt, placa_extra, df_fornecedores)
                if extras:
                    opcoes_extra = [f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})" for v in extras]
                    escolha_extra = st.selectbox("Selecione o veículo:", opcoes_extra, key="escolha_extra")
                    if st.button("➕ Adicionar", key="add_veic_extra"):
                        idx = opcoes_extra.index(escolha_extra)
                        novo = {
                            'placa': extras[idx]['placa'],
                            'marca': extras[idx]['marca'],
                            'modelo': extras[idx]['modelo'],
                            'ano': extras[idx]['ano'],
                            'renavan': '',
                            'n_equipamento': '',
                            'mot_cpf': '',
                            'mot_nome': ''
                        }
                        if any(limpar_placa(v['placa']) == limpar_placa(novo['placa']) for v in st.session_state.todos_veiculos):
                            st.warning("Veículo já está na lista.")
                        else:
                            st.session_state.todos_veiculos.append(novo)
                            st.success(f"✅ {novo['placa']} adicionado!")
                            st.rerun()
                else:
                    st.warning("Nenhum veículo encontrado.")
        else:
            filtro = st.text_input("Filtrar por placa/modelo (opcional):", key="filtro_veic")
            todos_da_base = []
            for _, row in df_veiculos_antt.iterrows():
                pl = safe_str(row['Placa'])
                ma = safe_str(row.get('Marca', ''))
                mo = safe_str(row.get('Modelo', ''))
                an = safe_str(row.get('Ano', ''))
                if filtro:
                    if filtro.upper() in pl.upper() or filtro.upper() in ma.upper() or filtro.upper() in mo.upper():
                        todos_da_base.append(f"{pl} - {ma} {mo} ({an})")
                else:
                    todos_da_base.append(f"{pl} - {ma} {mo} ({an})")
            if todos_da_base:
                st.write(f"Total: {len(todos_da_base)} veículos")
                escolha_lista = st.selectbox("Selecione um veículo:", todos_da_base[:50], key="escolha_lista")
                if st.button("➕ Adicionar da lista", key="add_lista"):
                    placa_sel = escolha_lista.split(" - ")[0]
                    mask = df_veiculos_antt['Placa'].apply(limpar_placa) == limpar_placa(placa_sel)
                    if mask.any():
                        row = df_veiculos_antt[mask].iloc[0]
                        novo = {
                            'placa': safe_str(row['Placa']),
                            'marca': safe_str(row.get('Marca', '')),
                            'modelo': safe_str(row.get('Modelo', '')),
                            'ano': safe_str(row.get('Ano', '')),
                            'renavan': safe_str(row.get('Renavan', '')),
                            'n_equipamento': '',
                            'mot_cpf': '',
                            'mot_nome': ''
                        }
                        if any(limpar_placa(v['placa']) == limpar_placa(novo['placa']) for v in st.session_state.todos_veiculos):
                            st.warning("Veículo já está na lista.")
                        else:
                            st.session_state.todos_veiculos.append(novo)
                            st.success(f"✅ {novo['placa']} adicionado!")
                            st.rerun()
            else:
                st.info("Nenhum veículo encontrado com esse filtro.")

    if st.session_state.veiculos_selecionados:
        if st.button("➡️ Próximo (Motoristas)"):
            st.session_state.etapa = "selecionar_motoristas"
            st.rerun()

# ------------------------------------------------------------
# ETAPA 4: SELECIONAR MOTORISTAS
# ------------------------------------------------------------
if st.session_state.etapa == "selecionar_motoristas":
    st.header("👤 4. Selecionar Motoristas")

    with st.spinner("Carregando motoristas..."):
        motoristas_auto = obter_motoristas_dos_veiculos(st.session_state.todos_veiculos, df_condutores, df_fornecedores)
        if not motoristas_auto:
            st.info("Nenhum motorista vinculado. Buscando todos os condutores da base...")
            motoristas_auto = obter_todos_condutores_da_base(df_condutores, df_fornecedores)

    if motoristas_auto:
        opcoes_mot = [f"{m['nome']} (CPF: {formatar_cpf(m['cpf'])})" for m in motoristas_auto]
        selecionados = st.multiselect(
            "Selecione os motoristas para o contrato (o primeiro será o principal):",
            opcoes_mot,
            default=opcoes_mot[:1] if opcoes_mot else []
        )
        if selecionados:
            st.session_state.motoristas_selecionados = [motoristas_auto[opcoes_mot.index(s)] for s in selecionados]
    else:
        st.warning("Nenhum motorista encontrado na base. Preencha manualmente:")
        with st.form("motorista_manual"):
            nome = st.text_input("Nome do motorista:")
            cpf = st.text_input("CPF:")
            cnh = st.text_input("CNH:")
            rg = st.text_input("RG:")
            if st.form_submit_button("✅ Adicionar motorista"):
                st.session_state.motoristas_selecionados = [{'nome': nome, 'cpf': cpf, 'cnh': cnh, 'rg': rg}]
                st.rerun()

    if st.session_state.motoristas_selecionados:
        if st.button("➡️ Gerar Contrato"):
            st.session_state.etapa = "gerar_contrato"
            st.rerun()

# ------------------------------------------------------------
# ETAPA 5: GERAR CONTRATO
# ------------------------------------------------------------
if st.session_state.etapa == "gerar_contrato":
    st.header("📄 5. Gerar Contrato")

    proprietario = st.session_state.proprietario
    veiculos_selecionados = st.session_state.veiculos_selecionados
    motoristas_selecionados = st.session_state.motoristas_selecionados

    if not proprietario or not veiculos_selecionados or not motoristas_selecionados:
        st.error("Dados incompletos. Volte e preencha todas as etapas.")
        st.stop()

    # ANTT
    antt = buscar_antt(proprietario, veiculos_selecionados, df_veiculos_antt, df_veiculos_compl)
    if not antt:
        antt = st.text_input("ANTT/RNTRC (não encontrado automaticamente):", value="", key="antt_manual")
        if not antt:
            st.warning("Digite o ANTT manualmente para continuar.")
            st.stop()

    # Serial
    serial = buscar_serial(veiculos_selecionados)
    if not serial:
        serial = st.text_input("Serial do rastreador:", value="A DEFINIR", key="serial_manual")

    # Valor
    valor = st.text_input("Valor por entrega (R$):", value="X (a definir)", key="valor")

    is_pf = len(limpar_cpf_cnpj(proprietario['cpf_cnpj'])) == 11
    if is_pf:
        numero_casa = st.text_input("Número da residência:", value="")
        rg_prop = buscar_rg_proprietario(proprietario['cpf_cnpj'], df_condutores, df_fornecedores)
        if not rg_prop:
            rg_prop = st.text_input("RG do proprietário:", value="")
        estado_civil = st.text_input("Estado civil:", value="")
    else:
        # Para PJ, vamos obter CEP
        cidade = proprietario.get('cidade', '')
        uf = proprietario.get('uf', '')
        endereco_raw = proprietario.get('endereco', '')
        rua, num, comp = extrair_endereco_pj(endereco_raw)
        if not rua:
            rua = endereco_raw
        if not num:
            num = 'S/N'
        if not comp:
            comp = proprietario.get('bairro', '')
        cep = buscar_cep_online(rua, cidade, uf)
        if not cep:
            cep = st.text_input(f"CEP para {cidade}/{uf}:", value="", key="cep_manual")
            if not cep:
                cep = "00000-000"
        numero_casa = num
        rg_prop = ''
        estado_civil = ''

    if st.button("🚀 Gerar Contrato"):
        cpf_cnpj_limpo = limpar_cpf_cnpj(proprietario['cpf_cnpj'])
        is_pf = len(cpf_cnpj_limpo) == 11

        dia_cad, mes_cad, ano_cad = formatar_data_cadastro(proprietario.get('data_inclusao', ''))
        cidade = proprietario.get('cidade', '') or "Cotia"
        uf = proprietario.get('uf', '') or "SP"
        endereco_raw = proprietario.get('endereco', '')
        bairro = proprietario.get('bairro', '')
        endereco_completo = ', '.join(filter(None, [endereco_raw, bairro, cidade, uf])) or 'ENDEREÇO NÃO CADASTRADO'

        veiculo_principal = veiculos_selecionados[0]
        motorista_principal = motoristas_selecionados[0]

        motoristas_lista = [
            {
                'nome': m.get('nome', ''),
                'cpf': formatar_cpf(m.get('cpf', '')),
                'cnh': m.get('cnh', ''),
            }
            for m in motoristas_selecionados
        ]

        veiculos_lista = [
            {
                'modelo': f"{v.get('marca','')} {v.get('modelo','')}".strip(),
                'ano': v.get('ano', ''),
                'renavam': v.get('renavan', ''),
            }
            for v in veiculos_selecionados
        ]

        contexto = {
            'dia_cadastro': dia_cad,
            'mes_cadastro': mes_cad,
            'ano_cadastro': ano_cad,
            'valor_por_entrega': valor,
            'endereco_contratado': endereco_completo,
            'antt_tac': antt,
            'antt_etc': antt,
            'placa_veiculo': veiculo_principal.get('placa', ''),
            'serial_equipamento': serial,
            'is_pf': is_pf,
            'nome_motorista': motorista_principal.get('nome', ''),
            'cpf_motorista': formatar_cpf(motorista_principal.get('cpf', '')),
            'cnh_motorista': motorista_principal.get('cnh', ''),
            'modelo_veiculo': f"{veiculo_principal.get('marca','')} {veiculo_principal.get('modelo','')}".strip(),
            'ano_veiculo': veiculo_principal.get('ano', ''),
            'renavam_veiculo': veiculo_principal.get('renavan', ''),
            'motoristas_lista': motoristas_lista,
            'veiculos_lista': veiculos_lista,
        }

        if is_pf:
            contexto.update({
                'nome_contratado_pf': proprietario['nome'],
                'estado_civil': estado_civil,
                'rg_contratado': rg_prop,
                'cpf_contratado': formatar_cpf(proprietario['cpf_cnpj']),
                'numero_da_casa': numero_casa,
                'razao_social_pj': '',
                'cnpj_pj': '',
                'rua_pj': '',
                'numero_pj': '',
                'complemento_pj': '',
                'municipio_pj': '',
                'estado_pj': '',
                'cep_pj_1': '',
            })
        else:
            # Usa rua, num, comp já extraídos acima
            rua, num, comp = extrair_endereco_pj(endereco_raw)
            if not rua:
                rua = endereco_raw
            if not num:
                num = 'S/N'
            if not comp:
                comp = bairro
            # Já temos cep
            contexto.update({
                'nome_contratado_pf': '',
                'estado_civil': '',
                'rg_contratado': '',
                'cpf_contratado': '',
                'numero_da_casa': '',
                'razao_social_pj': proprietario['nome'],
                'cnpj_pj': formatar_cnpj(proprietario['cpf_cnpj']),
                'rua_pj': rua,
                'numero_pj': num,
                'complemento_pj': comp,
                'municipio_pj': cidade,
                'estado_pj': uf,
                'cep_pj_1': cep,
            })

        try:
            with st.spinner("Gerando contrato..."):
                doc = DocxTemplate(template_file)
                doc.render(contexto)
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                nome_arquivo = f"CONTRATO_{proprietario['nome'].replace(' ', '_').upper()}_{veiculo_principal['placa']}.docx"
                st.download_button(
                    label="📥 Baixar Contrato",
                    data=buffer,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("✅ Contrato gerado com sucesso!")
        except Exception as e:
            st.error(f"❌ Erro ao gerar contrato: {e}")

    if st.button("🔄 Novo contrato"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
