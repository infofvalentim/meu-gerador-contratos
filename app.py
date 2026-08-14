import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import requests
from datetime import datetime
import os

st.set_page_config(page_title="Gerador de Contratos", layout="wide")
st.title("🚛 Gerador de Contratos de Transporte")

# ============================================================
# CONFIGURAÇÃO DOS ARQUIVOS (pasta dados/)
# ============================================================
PASTA_DADOS = "dados"
ARQUIVOS = {
    "fornecedores": os.path.join(PASTA_DADOS, "fornecedores.xlsx"),
    "veiculos_antt": os.path.join(PASTA_DADOS, "veiculos_antt.xlsx"),
    "condutores": os.path.join(PASTA_DADOS, "condutores.xlsx"),
    "veiculos_compl": os.path.join(PASTA_DADOS, "veiculos_complementares.xlsx"),
    "template": os.path.join(PASTA_DADOS, "template_contrato.docx"),
}

# ============================================================
# FUNÇÕES DE CARREGAMENTO (com cache)
# ============================================================
@st.cache_data
def carregar_excel(caminho, header=4):
    try:
        return pd.read_excel(caminho, dtype=str, engine='openpyxl', header=header)
    except Exception:
        return None

@st.cache_data
def carregar_template(caminho):
    try:
        with open(caminho, "rb") as f:
            return f.read()
    except Exception:
        return None

def verificar_arquivos():
    faltantes = []
    for nome, caminho in ARQUIVOS.items():
        if not os.path.exists(caminho):
            faltantes.append(nome)
    return faltantes

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def limpar_cpf_cnpj(val):
    if pd.isna(val): return ''
    return ''.join(filter(str.isdigit, str(val)))

def formatar_cpf(cpf):
    cpf = limpar_cpf_cnpj(cpf)
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    return cpf

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
    if pd.isna(data_str) or not str(data_str).strip():
        hoje = datetime.now()
        return str(hoje.day), meses[hoje.month], str(hoje.year)
    try:
        data_str = str(data_str).strip()
        if '/' in data_str:
            partes = data_str.split('/')
            if len(partes) == 3:
                dt = datetime(int(partes[2]), int(partes[1]), int(partes[0]))
                return str(dt.day), meses[dt.month], str(dt.year)
        if '-' in data_str:
            partes = data_str.split('-')
            if len(partes) == 3 and len(partes[0]) == 4:
                dt = datetime(int(partes[0]), int(partes[1]), int(partes[2]))
                return str(dt.day), meses[dt.month], str(dt.year)
        dt = pd.to_datetime(data_str)
        return str(dt.day), meses[dt.month], str(dt.year)
    except:
        hoje = datetime.now()
        return str(hoje.day), meses[hoje.month], str(hoje.year)

def extrair_endereco_pj(texto_endereco):
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
    # Remover prefixos para evitar duplicação "Rua RUA"
    prefixos = ['RUA ', 'R. ', 'AVENIDA ', 'AV. ', 'ALAMEDA ', 'TRAVESSA ', 'TRAV. ']
    for prefixo in prefixos:
        if rua.upper().startswith(prefixo):
            rua = rua[len(prefixo):].strip()
            break
    return rua, numero, complemento

def buscar_cep_online(rua, cidade, uf):
    if not rua or not cidade or not uf:
        return ''
    termos = re.findall(r'\b[A-ZÀ-Ú]{3,}\b', rua.upper())
    if not termos:
        return ''
    query = ' '.join(termos[:3])
    url = f"https://viacep.com.br/ws/{uf}/{cidade}/{query}/json/"
    try:
        resp = requests.get(url, timeout=5)
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

# ============================================================
# FUNÇÃO PARA IDENTIFICAR COLUNA DE RNTRC
# ============================================================
def get_rntrc_column(df):
    """Retorna o nome da coluna de RNTRC em um DataFrame, ou None se não encontrar."""
    possiveis = ['Rntrc', 'RNTRC', 'ANTT', 'ANTT/RNTRC', 'ANTT/RNTRC nº', 'Rntrc/ANTT']
    for col in df.columns:
        if col.strip().upper() in [p.upper() for p in possiveis]:
            return col
    return None

# ============================================================
# FUNÇÕES DE NEGÓCIO ATUALIZADAS (BUSCA PELO RNTRC, com fallback)
# ============================================================
COLUNAS_MOTORISTA = ['Motorista Cnpj', 'Motorista CPF', 'Condutor CPF', 'Motorista', 'CPF Motorista', 'Condutor']

def buscar_veiculos_por_nome_fallback(df_veic_antt, df_forn, termo):
    """Fallback: busca veículos pelo nome do proprietário (coluna Proprietário Cnpj)."""
    termo_norm = normalizar_nome(termo)
    veiculos = []
    for _, row in df_veic_antt.iterrows():
        prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
        if prop_cnpj:
            p = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
            if not p.empty:
                nome_forn = safe_str(p.iloc[0]['Nome Fornecedor'])
                if termo_norm in normalizar_nome(nome_forn):
                    veiculos.append({
                        'placa': safe_str(row['Placa']),
                        'marca': safe_str(row.get('Marca', '')),
                        'modelo': safe_str(row.get('Modelo', '')),
                        'ano': safe_str(row.get('Ano', '')),
                        'rntrc': '',
                        'prop_nome': nome_forn,
                        'prop_cnpj': prop_cnpj,
                        'row': row
                    })
    return veiculos

def buscar_veiculos_por_nome(df_veic_antt, df_forn, termo):
    """Busca veículos por parte do nome do transportador (fornecedor) e retorna com os dados do transportador (via RNTRC)."""
    if not termo:
        return []
    termo_norm = normalizar_nome(termo)
    # Filtrar fornecedores que contenham o termo no nome
    mask_forn = df_forn['Nome Fornecedor'].apply(normalizar_nome).str.contains(termo_norm, na=False)
    fornecedores_filtrados = df_forn[mask_forn]
    if fornecedores_filtrados.empty:
        return []
    
    col_rntrc_forn = get_rntrc_column(df_forn)
    if col_rntrc_forn is None:
        st.warning("⚠️ Coluna de RNTRC não encontrada em fornecedores. Usando busca pelo CNPJ/CPF do proprietário.")
        return buscar_veiculos_por_nome_fallback(df_veic_antt, df_forn, termo)
    
    # Obter os RNTRCs desses fornecedores
    rntrcs = set()
    for _, row in fornecedores_filtrados.iterrows():
        rntrc = safe_str(row.get(col_rntrc_forn, ''))
        if rntrc:
            rntrcs.add(limpar_cpf_cnpj(rntrc))
    if not rntrcs:
        return buscar_veiculos_por_nome_fallback(df_veic_antt, df_forn, termo)
    
    col_rntrc_veic = get_rntrc_column(df_veic_antt)
    if col_rntrc_veic is None:
        return buscar_veiculos_por_nome_fallback(df_veic_antt, df_forn, termo)
    
    veiculos = []
    for _, row in df_veic_antt.iterrows():
        rntrc_veic = safe_str(row.get(col_rntrc_veic, ''))
        if limpar_cpf_cnpj(rntrc_veic) in rntrcs:
            placa = safe_str(row['Placa'])
            marca = safe_str(row.get('Marca', ''))
            modelo = safe_str(row.get('Modelo', ''))
            ano = safe_str(row.get('Ano', ''))
            p = df_forn[df_forn[col_rntrc_forn].apply(limpar_cpf_cnpj) == limpar_cpf_cnpj(rntrc_veic)]
            if not p.empty:
                prop_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
                prop_cnpj = safe_str(p.iloc[0]['Cnpj/Cpf'])
            else:
                prop_nome = ''
                prop_cnpj = ''
            veiculos.append({
                'placa': placa,
                'marca': marca,
                'modelo': modelo,
                'ano': ano,
                'rntrc': rntrc_veic,
                'prop_nome': prop_nome,
                'prop_cnpj': prop_cnpj,
                'row': row
            })
    return veiculos

def buscar_veiculo_por_placa(df_veic_antt, termo, df_forn):
    """Busca veículo por placa (ou parte) e retorna com os dados do transportador via RNTRC."""
    if not termo:
        return []
    mask = df_veic_antt['Placa'].str.upper().str.contains(termo.upper(), na=False)
    resultados = df_veic_antt[mask]
    veiculos = []
    col_rntrc_veic = get_rntrc_column(df_veic_antt)
    col_rntrc_forn = get_rntrc_column(df_forn)
    for _, row in resultados.iterrows():
        placa = safe_str(row['Placa'])
        marca = safe_str(row.get('Marca', ''))
        modelo = safe_str(row.get('Modelo', ''))
        ano = safe_str(row.get('Ano', ''))
        rntrc_veic = safe_str(row.get(col_rntrc_veic, '')) if col_rntrc_veic else ''
        prop_nome = ''
        prop_cnpj = ''
        if col_rntrc_forn and col_rntrc_veic and rntrc_veic:
            p = df_forn[df_forn[col_rntrc_forn].apply(limpar_cpf_cnpj) == limpar_cpf_cnpj(rntrc_veic)]
            if not p.empty:
                prop_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
                prop_cnpj = safe_str(p.iloc[0]['Cnpj/Cpf'])
        # Fallback: se não achou pelo RNTRC, tenta pelo proprietário
        if not prop_nome:
            prop_cnpj_fallback = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
            if prop_cnpj_fallback:
                p2 = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj_fallback]
                if not p2.empty:
                    prop_nome = safe_str(p2.iloc[0]['Nome Fornecedor'])
                    prop_cnpj = safe_str(p2.iloc[0]['Cnpj/Cpf'])
        veiculos.append({
            'placa': placa,
            'marca': marca,
            'modelo': modelo,
            'ano': ano,
            'rntrc': rntrc_veic,
            'prop_nome': prop_nome,
            'prop_cnpj': prop_cnpj,
            'row': row
        })
    return veiculos

def obter_veiculos_do_transportador(rntrc, df_veic_antt, df_veic_compl, df_cond, df_forn):
    """Retorna todos os veículos que possuem o mesmo RNTRC (transportador)."""
    if not rntrc:
        return []
    rntrc_limpo = limpar_cpf_cnpj(rntrc)
    if not rntrc_limpo:
        return []
    col_rntrc_veic = get_rntrc_column(df_veic_antt)
    if col_rntrc_veic is None:
        st.warning("⚠️ Coluna de RNTRC não encontrada em veículos. Não é possível listar veículos do transportador.")
        return []
    
    mask = df_veic_antt[col_rntrc_veic].apply(limpar_cpf_cnpj) == rntrc_limpo
    veiculos = []
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
            m = df_cond[df_cond['CPF N°'].apply(limpar_cpf_cnpj) == mot_cpf]
            if not m.empty:
                mot_nome = safe_str(m.iloc[0]['Nome'])
            else:
                m2 = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == mot_cpf]
                if not m2.empty:
                    mot_nome = safe_str(m2.iloc[0]['Nome Fornecedor'])
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
    for v in veiculos:
        cpf = v.get('mot_cpf')
        if cpf and cpf not in cpfs_vistos:
            cpfs_vistos.add(cpf)
            nome = v.get('mot_nome', '')
            cnh = ''
            rg = ''
            if nome:
                m = df_condutores[df_condutores['CPF N°'].apply(limpar_cpf_cnpj) == cpf]
                if not m.empty:
                    row = m.iloc[0]
                    nome = safe_str(row.get('Nome', '')) or nome
                    cnh = safe_str(row.get('CNH N°', ''))
                    rg = safe_str(row.get('RG N°', ''))
                else:
                    m2 = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cpf]
                    if not m2.empty:
                        nome = safe_str(m2.iloc[0].get('Nome Fornecedor', '')) or nome
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

def buscar_antt(transportador, veiculos_selecionados, df_veiculos_antt, df_veiculos_compl):
    antt = transportador.get('rntrc', '')
    if antt and antt.lower() != 'nan':
        return antt
    # Fallback: busca nos veículos
    for v in veiculos_selecionados:
        placa_limpa = limpar_placa(v['placa'])
        if not df_veiculos_compl.empty:
            vc = df_veiculos_compl[df_veiculos_compl['Placa'].apply(limpar_placa) == placa_limpa]
            if not vc.empty:
                antt_tmp = safe_str(vc.iloc[0].get('ANTT/RNTRC nº', ''))
                if antt_tmp and antt_tmp.lower() != 'nan':
                    return antt_tmp
    col_rntrc = get_rntrc_column(df_veiculos_antt)
    if col_rntrc:
        for v in veiculos_selecionados:
            placa_limpa = limpar_placa(v['placa'])
            va = df_veiculos_antt[df_veiculos_antt['Placa'].apply(limpar_placa) == placa_limpa]
            if not va.empty:
                antt_tmp = safe_str(va.iloc[0].get(col_rntrc, ''))
                if antt_tmp and antt_tmp.lower() != 'nan':
                    return antt_tmp
    return ''

def buscar_serial(veiculos_selecionados):
    for v in veiculos_selecionados:
        if v.get('n_equipamento') and v['n_equipamento'].lower() != 'nan':
            return v['n_equipamento']
    return ''

# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

# --- Verificar arquivos ---
faltantes = verificar_arquivos()
if faltantes:
    st.warning(f"⚠️ Arquivos não encontrados na pasta `dados/`: {', '.join(faltantes)}")
    st.info("📂 Certifique-se de que todos os arquivos estejam na pasta `dados/`.")
    st.stop()
else:
    st.success("✅ Todos os arquivos encontrados na pasta `dados/`!")

# --- Carregar dados ---
df_fornecedores = carregar_excel(ARQUIVOS["fornecedores"], header=4)
df_veiculos_antt = carregar_excel(ARQUIVOS["veiculos_antt"], header=4)
df_condutores = carregar_excel(ARQUIVOS["condutores"], header=4)
df_veiculos_compl = carregar_excel(ARQUIVOS["veiculos_compl"], header=4)
template_bytes = carregar_template(ARQUIVOS["template"])

if any(df is None for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]) or template_bytes is None:
    st.error("❌ Erro ao carregar um ou mais arquivos.")
    st.stop()

for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]:
    df.dropna(how='all', inplace=True)
    df.columns = df.columns.str.strip()

# ============================================================
# ESTADO DA SESSÃO
# ============================================================
if "etapa" not in st.session_state:
    st.session_state.etapa = "busca"
if "veiculo_escolhido" not in st.session_state:
    st.session_state.veiculo_escolhido = None
if "transportador" not in st.session_state:
    st.session_state.transportador = None
if "todos_veiculos" not in st.session_state:
    st.session_state.todos_veiculos = []
if "veiculos_selecionados" not in st.session_state:
    st.session_state.veiculos_selecionados = []
if "motoristas_selecionados" not in st.session_state:
    st.session_state.motoristas_selecionados = []

# ============================================================
# ETAPA 1: BUSCA (por placa OU por nome do transportador)
# ============================================================
st.header("🔍 1. Buscar Veículo")

col_busca1, col_busca2 = st.columns(2)

with col_busca1:
    modo_busca = st.radio("Buscar por:", ["Placa", "Nome do Transportador (Fornecedor)"], horizontal=True)

with col_busca2:
    if modo_busca == "Placa":
        termo_busca = st.text_input("Digite a placa (ou parte):", key="placa_busca")
    else:
        termo_busca = st.text_input("Digite o nome do transportador (ou parte):", key="nome_busca")

resultados = []
if termo_busca:
    if modo_busca == "Placa":
        resultados = buscar_veiculo_por_placa(df_veiculos_antt, termo_busca, df_fornecedores)
    else:
        resultados = buscar_veiculos_por_nome(df_veiculos_antt, df_fornecedores, termo_busca)

if resultados:
    st.success(f"✅ {len(resultados)} veículo(s) encontrado(s).")
    
    opcoes = []
    for v in resultados:
        opcoes.append(f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']}) - Transportador: {v['prop_nome'] or 'N/I'}")
    
    selecao = st.selectbox("Selecione o veículo principal:", opcoes, key="selecao_veiculo_result")
    
    if st.button("✅ Confirmar veículo"):
        idx = opcoes.index(selecao)
        st.session_state.veiculo_escolhido = resultados[idx]
        st.session_state.etapa = "confirmar_transportador"
        st.rerun()
else:
    if termo_busca:
        st.warning("Nenhum veículo encontrado.")

# ============================================================
# ETAPA 2: CONFIRMAR TRANSPORTADOR (pelo RNTRC)
# ============================================================
if st.session_state.etapa == "confirmar_transportador":
    veic = st.session_state.veiculo_escolhido
    st.header("👤 2. Confirmar Transportador")

    rntrc_veic = veic.get('rntrc', '')
    transportador = None

    if rntrc_veic:
        col_rntrc_forn = get_rntrc_column(df_fornecedores)
        if col_rntrc_forn:
            p = df_fornecedores[df_fornecedores[col_rntrc_forn].apply(limpar_cpf_cnpj) == limpar_cpf_cnpj(rntrc_veic)]
            if not p.empty:
                row = p.iloc[0]
                transportador = {
                    'nome': safe_str(row['Nome Fornecedor']),
                    'cpf_cnpj': safe_str(row['Cnpj/Cpf']),
                    'endereco': safe_str(row.get('Endereço', '')),
                    'bairro': safe_str(row.get('Bairro', '')),
                    'cidade': safe_str(row.get('Cidade', '')),
                    'uf': safe_str(row.get('UF', '')),
                    'data_inclusao': safe_str(row.get('Data Inclusão', '')),
                    'rntrc': rntrc_veic
                }

    if transportador:
        st.write(f"**Nome:** {transportador['nome']}")
        st.write(f"**CPF/CNPJ:** {formatar_cpf_cnpj(transportador['cpf_cnpj'])}")
        st.write(f"**RNTRC:** {transportador['rntrc']}")
        st.write(f"**Endereço:** {transportador['endereco']}, {transportador['bairro']}, {transportador['cidade']}/{transportador['uf']}")
        if st.button("✅ Confirmar Transportador"):
            st.session_state.transportador = transportador
            st.session_state.etapa = "selecionar_veiculos"
            st.session_state.todos_veiculos = obter_veiculos_do_transportador(
                transportador['rntrc'], df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
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
        st.warning("Transportador não encontrado na base para o RNTRC deste veículo. Preencha manualmente:")
        with st.form("manual_transportador"):
            nome = st.text_input("Nome do Transportador:", value=veic.get('prop_nome', ''))
            cpf_cnpj = st.text_input("CPF/CNPJ:", value=veic.get('prop_cnpj', ''))
            endereco = st.text_input("Endereço:")
            bairro = st.text_input("Bairro:")
            cidade = st.text_input("Cidade:")
            uf = st.text_input("UF:")
            data_inclusao = st.text_input("Data de inclusão:")
            rntrc_manual = st.text_input("RNTRC:", value=rntrc_veic)
            if st.form_submit_button("✅ Confirmar"):
                st.session_state.transportador = {
                    'nome': nome,
                    'cpf_cnpj': cpf_cnpj,
                    'endereco': endereco,
                    'bairro': bairro,
                    'cidade': cidade,
                    'uf': uf,
                    'data_inclusao': data_inclusao,
                    'rntrc': rntrc_manual
                }
                st.session_state.etapa = "selecionar_veiculos"
                if rntrc_manual:
                    st.session_state.todos_veiculos = obter_veiculos_do_transportador(
                        rntrc_manual, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
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

# ============================================================
# ETAPA 3: SELECIONAR VEÍCULOS
# ============================================================
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
            st.warning("Selecione pelo menos um veículo.")
    else:
        st.warning("Nenhum veículo disponível.")

    # Expandir para adicionar veículo extra da base completa
    with st.expander("➕ Adicionar veículo extra da base completa"):
        modo_busca_extra = st.radio("Como buscar?", ["Por placa", "Por nome do transportador", "Listar todos (filtrar)"], key="modo_busca_extra")
        if modo_busca_extra == "Por placa":
            placa_extra = st.text_input("Digite a placa (ou parte):", key="placa_extra")
            if placa_extra:
                extras = buscar_veiculo_por_placa(df_veiculos_antt, placa_extra, df_fornecedores)
                if extras:
                    rntrc_atual = st.session_state.transportador.get('rntrc', '') if st.session_state.transportador else ''
                    if rntrc_atual:
                        col_rntrc_veic = get_rntrc_column(df_veiculos_antt)
                        if col_rntrc_veic:
                            extras = [e for e in extras if limpar_cpf_cnpj(e.get('rntrc', '')) == limpar_cpf_cnpj(rntrc_atual)]
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
                        st.warning("Nenhum veículo com o mesmo RNTRC encontrado.")
                else:
                    st.warning("Nenhum veículo encontrado.")
        elif modo_busca_extra == "Por nome do transportador":
            nome_extra = st.text_input("Digite o nome do transportador (ou parte):", key="nome_extra")
            if nome_extra:
                extras = buscar_veiculos_por_nome(df_veiculos_antt, df_fornecedores, nome_extra)
                rntrc_atual = st.session_state.transportador.get('rntrc', '') if st.session_state.transportador else ''
                if rntrc_atual:
                    extras = [e for e in extras if limpar_cpf_cnpj(e.get('rntrc', '')) == limpar_cpf_cnpj(rntrc_atual)]
                if extras:
                    opcoes_extra = [f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']}) - Transportador: {v['prop_nome']}" for v in extras]
                    escolha_extra = st.selectbox("Selecione o veículo:", opcoes_extra, key="escolha_extra_nome")
                    if st.button("➕ Adicionar", key="add_veic_extra_nome"):
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
                    st.warning("Nenhum veículo com o mesmo RNTRC encontrado.")
        else:
            # Listar todos com filtro por placa/modelo, mas restringir ao RNTRC atual
            filtro = st.text_input("Filtrar por placa/modelo (opcional):", key="filtro_veic")
            rntrc_atual = st.session_state.transportador.get('rntrc', '') if st.session_state.transportador else ''
            col_rntrc_veic = get_rntrc_column(df_veiculos_antt)
            todos_da_base = []
            for _, row in df_veiculos_antt.iterrows():
                if col_rntrc_veic and rntrc_atual:
                    rntrc_veic = safe_str(row.get(col_rntrc_veic, ''))
                    if limpar_cpf_cnpj(rntrc_veic) != limpar_cpf_cnpj(rntrc_atual):
                        continue
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
                st.info("Nenhum veículo encontrado com esse filtro ou com o RNTRC do transportador.")

    if st.session_state.veiculos_selecionados:
        if st.button("➡️ Próximo (Motoristas)"):
            st.session_state.etapa = "selecionar_motoristas"
            st.rerun()

# ============================================================
# ETAPA 4: SELECIONAR MOTORISTAS
# ============================================================
if st.session_state.etapa == "selecionar_motoristas":
    st.header("👤 4. Selecionar Motoristas")

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
            st.warning("Selecione pelo menos um motorista.")
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

# ============================================================
# ETAPA 5: GERAR CONTRATO
# ============================================================
if st.session_state.etapa == "gerar_contrato":
    st.header("📄 5. Gerar Contrato")

    transportador = st.session_state.transportador
    veiculos_selecionados = st.session_state.veiculos_selecionados
    motoristas_selecionados = st.session_state.motoristas_selecionados

    if not transportador or not veiculos_selecionados or not motoristas_selecionados:
        st.error("Dados incompletos. Volte e preencha todas as etapas.")
        st.stop()

    st.subheader("📝 Dados Complementares")

    col1, col2 = st.columns(2)
    with col1:
        antt = transportador.get('rntrc', '')
        if not antt:
            antt = buscar_antt(transportador, veiculos_selecionados, df_veiculos_antt, df_veiculos_compl)
        if not antt:
            antt = st.text_input("ANTT/RNTRC (não encontrado automaticamente):", value="", key="antt_manual")
            if not antt:
                st.warning("Digite o ANTT manualmente para continuar.")
                st.stop()
        else:
            st.write(f"✅ ANTT/RNTRC encontrado: {antt}")

        serial = buscar_serial(veiculos_selecionados)
        if not serial:
            serial = st.text_input("Serial do rastreador:", value="A DEFINIR", key="serial_manual")
        else:
            st.write(f"✅ Serial encontrado: {serial}")

        valor = st.text_input("Valor por entrega (R$):", value="X (a definir)", key="valor")

    with col2:
        is_pf = len(limpar_cpf_cnpj(transportador['cpf_cnpj'])) == 11
        if not is_pf:
            cidade = transportador.get('cidade', '')
            uf = transportador.get('uf', '')
            endereco_raw = transportador.get('endereco', '')
            rua, _, _ = extrair_endereco_pj(endereco_raw)
            cep_sugerido = buscar_cep_online(rua, cidade, uf) if rua and cidade and uf else ''
            if cep_sugerido:
                st.write(f"✅ CEP sugerido: {cep_sugerido}")
                if st.button("Usar este CEP"):
                    st.session_state.cep_manual = cep_sugerido
                    st.rerun()
            else:
                cep_manual = st.text_input("CEP (ex: 00000-000):", value=st.session_state.get("cep_manual", ""), key="cep_input")
        else:
            st.write("🔹 Pessoa Física – CEP será buscado do endereço.")

    if is_pf:
        st.subheader("👤 Dados da Pessoa Física (Transportador)")
        col3, col4 = st.columns(2)
        with col3:
            numero_casa = st.text_input("Número da residência:", value=st.session_state.get("numero_casa", ""), key="num_casa")
            rg_prop = buscar_rg_proprietario(transportador['cpf_cnpj'], df_condutores, df_fornecedores)
            if not rg_prop:
                rg_prop = st.text_input("RG do transportador:", value=st.session_state.get("rg_prop", ""), key="rg_prop")
            else:
                st.write(f"✅ RG encontrado: {rg_prop}")
        with col4:
            estado_civil = st.text_input("Estado civil:", value=st.session_state.get("estado_civil", ""), key="estado_civil")
    else:
        numero_casa = ""
        rg_prop = ""
        estado_civil = ""

    if st.button("🚀 Gerar Contrato"):
        cpf_cnpj_limpo = limpar_cpf_cnpj(transportador['cpf_cnpj'])
        is_pf = len(cpf_cnpj_limpo) == 11

        dia_cad, mes_cad, ano_cad = formatar_data_cadastro(transportador.get('data_inclusao', ''))
        cidade = transportador.get('cidade', '') or "Cotia"
        uf = transportador.get('uf', '') or "SP"
        endereco_raw = transportador.get('endereco', '')
        bairro = transportador.get('bairro', '')
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
            rg_final = rg_prop or buscar_rg_proprietario(transportador['cpf_cnpj'], df_condutores, df_fornecedores) or ''
            contexto.update({
                'nome_contratado_pf': transportador['nome'],
                'estado_civil': estado_civil,
                'rg_contratado': rg_final,
                'cpf_contratado': formatar_cpf(transportador['cpf_cnpj']),
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
            rua, num, comp = extrair_endereco_pj(endereco_raw)
            if not rua: rua = endereco_raw
            if not num: num = 'S/N'
            if not comp: comp = bairro
            cep_final = st.session_state.get("cep_manual", "")
            if not cep_final:
                cep_final = buscar_cep_online(rua, cidade, uf)
                if not cep_final:
                    cep_final = "00000-000"
            contexto.update({
                'nome_contratado_pf': '',
                'estado_civil': '',
                'rg_contratado': '',
                'cpf_contratado': '',
                'numero_da_casa': '',
                'razao_social_pj': transportador['nome'],
                'cnpj_pj': formatar_cnpj(transportador['cpf_cnpj']),
                'rua_pj': rua,
                'numero_pj': num,
                'complemento_pj': comp,
                'municipio_pj': cidade,
                'estado_pj': uf,
                'cep_pj_1': cep_final,
            })

        try:
            doc = DocxTemplate(io.BytesIO(template_bytes))
            doc.render(contexto)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            nome_arquivo = f"CONTRATO_{transportador['nome'].replace(' ', '_').upper()}_{veiculo_principal['placa']}.docx"
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
            if key not in ["df_fornecedores", "df_veiculos_antt", "df_condutores", "df_veiculos_compl", "template_bytes"]:
                del st.session_state[key]
        st.rerun()
