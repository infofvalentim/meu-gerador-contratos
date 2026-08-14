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
def carregar_excel(caminho, skiprows=4):
    """
    Carrega um arquivo Excel pulando as linhas iniciais.
    Assume que a linha com os cabeçalhos está em skiprows (0-indexado).
    Se não encontrar, tenta com header=0.
    """
    try:
        # Primeiro tenta com skiprows e header=0 (primeira linha após os pulos)
        df = pd.read_excel(caminho, dtype=str, engine='openpyxl', skiprows=skiprows, header=0)
        # Verifica se alguma coluna parece ser de dados (se todas as colunas forem nulas, pode ser erro)
        if df.dropna(how='all').empty:
            raise ValueError("Dados vazios após skiprows")
        return df
    except Exception:
        # Fallback: tenta ler com header=0 sem skiprows
        try:
            df = pd.read_excel(caminho, dtype=str, engine='openpyxl', header=0)
            if df.dropna(how='all').empty:
                raise ValueError("Dados vazios com header=0")
            return df
        except Exception:
            # Último recurso: tentar com header=None e depois definir colunas manualmente
            df = pd.read_excel(caminho, dtype=str, engine='openpyxl', header=None)
            # Procura a linha que contém os cabeçalhos conhecidos
            for i, row in df.iterrows():
                row_str = ' '.join(row.astype(str))
                if 'Placa' in row_str or 'Nome Fornecedor' in row_str or 'CPF N°' in row_str:
                    # Usa essa linha como cabeçalho
                    new_header = row
                    df = df.iloc[i+1:]
                    df.columns = new_header
                    df.reset_index(drop=True, inplace=True)
                    return df
            # Se não encontrar, retorna None
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
    if pd.isna(val) or val is None:
        return ''
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
    if pd.isna(nome): return ''
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
    # Remover prefixos
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
# FUNÇÕES DE NEGÓCIO - BUSCA POR TRANSPORTADOR (Fornecedor)
# ============================================================
COLUNAS_MOTORISTA = ['Motorista Cnpj', 'Motorista CPF', 'Condutor CPF', 'Motorista', 'CPF Motorista', 'Condutor']

def buscar_veiculos_por_transportador(df_veic_antt, df_forn, termo):
    """
    Busca veículos cujo proprietário (campo 'Proprietário Cnpj') corresponda a um fornecedor
    cujo nome contenha o termo buscado.
    """
    if not termo:
        return []
    termo_norm = normalizar_nome(termo)
    # Filtrar fornecedores pelo nome
    mask_forn = df_forn['Nome Fornecedor'].apply(normalizar_nome).str.contains(termo_norm, na=False, case=False)
    fornecedores_filtrados = df_forn[mask_forn]
    if fornecedores_filtrados.empty:
        return []
    # Obter CPF/CNPJ dos fornecedores filtrados
    cpfs_cnpjs = set()
    for _, row in fornecedores_filtrados.iterrows():
        cpf_cnpj = limpar_cpf_cnpj(row['Cnpj/Cpf'])
        if cpf_cnpj:
            cpfs_cnpjs.add(cpf_cnpj)
    # Filtrar veículos que tenham 'Proprietário Cnpj' igual a algum dos CPF/CNPJ
    veiculos = []
    for _, row in df_veic_antt.iterrows():
        prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
        if prop_cnpj in cpfs_cnpjs:
            placa = safe_str(row['Placa'])
            marca = safe_str(row.get('Marca', ''))
            modelo = safe_str(row.get('Modelo', ''))
            ano = safe_str(row.get('Ano', ''))
            # Buscar o nome do fornecedor (transportador)
            p = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
            if not p.empty:
                transportador_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
            else:
                transportador_nome = 'N/I'
            veiculos.append({
                'placa': placa,
                'marca': marca,
                'modelo': modelo,
                'ano': ano,
                'prop_cnpj': prop_cnpj,
                'transportador_nome': transportador_nome,
                'row': row
            })
    return veiculos

def buscar_veiculo_por_placa(df_veic_antt, termo, df_forn):
    if not termo:
        return []
    mask = df_veic_antt['Placa'].str.upper().str.contains(termo.upper(), na=False)
    resultados = df_veic_antt[mask]
    veiculos = []
    for _, row in resultados.iterrows():
        placa = safe_str(row['Placa'])
        marca = safe_str(row.get('Marca', ''))
        modelo = safe_str(row.get('Modelo', ''))
        ano = safe_str(row.get('Ano', ''))
        prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
        transportador_nome = ''
        if prop_cnpj:
            p = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
            if not p.empty:
                transportador_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
        veiculos.append({
            'placa': placa,
            'marca': marca,
            'modelo': modelo,
            'ano': ano,
            'prop_cnpj': prop_cnpj,
            'transportador_nome': transportador_nome,
            'row': row
        })
    return veiculos

def obter_veiculos_do_transportador(cnpj_prop, df_veic_antt, df_veic_compl, df_cond, df_forn):
    """
    Retorna todos os veículos de um determinado transportador (CNPJ/CPF).
    """
    cnpj_limpo = limpar_cpf_cnpj(cnpj_prop)
    if not cnpj_limpo:
        return []
    mask = df_veic_antt['Proprietário Cnpj'].apply(limpar_cpf_cnpj) == cnpj_limpo
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

def obter_motoristas_dos_veiculos(veiculos, df_cond, df_forn):
    """
    Para uma lista de veículos, retorna uma lista de motoristas (CPF, Nome) únicos.
    """
    motoristas = {}
    for v in veiculos:
        cpf = v.get('mot_cpf', '')
        if cpf:
            if cpf not in motoristas:
                nome = v.get('mot_nome', '')
                if not nome:
                    # Buscar nos dataframes
                    m = df_cond[df_cond['CPF N°'].apply(limpar_cpf_cnpj) == cpf]
                    if not m.empty:
                        nome = safe_str(m.iloc[0]['Nome'])
                    else:
                        m2 = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cpf]
                        if not m2.empty:
                            nome = safe_str(m2.iloc[0]['Nome Fornecedor'])
                motoristas[cpf] = nome
    return [{'cpf': cpf, 'nome': nome} for cpf, nome in motoristas.items()]

def obter_todos_condutores_da_base(df_cond):
    """
    Retorna todos os condutores da base para seleção manual.
    """
    condutores = []
    for _, row in df_cond.iterrows():
        cpf = limpar_cpf_cnpj(row.get('CPF N°', ''))
        nome = safe_str(row.get('Nome', ''))
        if cpf and nome:
            condutores.append({'cpf': cpf, 'nome': nome})
    return condutores

def buscar_antt(placa, df_veic_antt):
    """
    Busca dados ANTT de um veículo pela placa.
    """
    placa_limpa = limpar_placa(placa)
    mask = df_veic_antt['Placa'].apply(limpar_placa) == placa_limpa
    if mask.any():
        row = df_veic_antt[mask].iloc[0]
        return {
            'rntrc': safe_str(row.get('Rntrc', '')),
            'renavan': safe_str(row.get('Renavan', '')),
            'validade_rntrc': safe_str(row.get('Validade Rntrc', ''))
        }
    return {}

def buscar_serial(chassi, df_veic_antt):
    """
    Busca o número de série (chassi) de um veículo.
    """
    if not chassi:
        return ''
    mask = df_veic_antt['Nr Chassis'].apply(lambda x: limpar_cpf_cnpj(x) == limpar_cpf_cnpj(chassi))
    if mask.any():
        return safe_str(df_veic_antt[mask].iloc[0].get('Nr Chassis', ''))
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
df_fornecedores = carregar_excel(ARQUIVOS["fornecedores"], skiprows=4)
df_veiculos_antt = carregar_excel(ARQUIVOS["veiculos_antt"], skiprows=4)
df_condutores = carregar_excel(ARQUIVOS["condutores"], skiprows=4)
df_veiculos_compl = carregar_excel(ARQUIVOS["veiculos_compl"], skiprows=4)
template_bytes = carregar_template(ARQUIVOS["template"])

if any(df is None for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]) or template_bytes is None:
    st.error("❌ Erro ao carregar um ou mais arquivos. Verifique a estrutura dos arquivos Excel.")
    st.stop()

# Limpeza básica
for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]:
    df.dropna(how='all', inplace=True)
    df.columns = df.columns.str.strip()

# Debug: mostrar colunas para verificação (pode ser removido depois)
# st.write("Colunas de fornecedores:", df_fornecedores.columns.tolist())
# st.write("Colunas de veículos ANTT:", df_veiculos_antt.columns.tolist())

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
    modo_busca = st.radio("Buscar por:", ["Placa", "Nome do Transportador"], horizontal=True)

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
        resultados = buscar_veiculos_por_transportador(df_veiculos_antt, df_fornecedores, termo_busca)

if resultados:
    st.success(f"✅ {len(resultados)} veículo(s) encontrado(s).")
    
    # Exibir lista de veículos para seleção
    opcoes = []
    for v in resultados:
        opcoes.append(f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']}) - Transportador: {v['transportador_nome'] or 'N/I'}")
    
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
# ETAPA 2: CONFIRMAR TRANSPORTADOR (Fornecedor)
# ============================================================
if st.session_state.etapa == "confirmar_transportador":
    veic = st.session_state.veiculo_escolhido
    st.header("👤 2. Confirmar Transportador")

    cnpj_prop = veic['prop_cnpj']
    transportador_nome_sugerido = veic['transportador_nome']
    transportador = None
    if cnpj_prop:
        p = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cnpj_prop]
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
                'rntrc': safe_str(row.get('Rntrc', ''))
            }

    if transportador:
        st.write(f"**Transportador:** {transportador['nome']}")
        st.write(f"**CPF/CNPJ:** {formatar_cpf_cnpj(transportador['cpf_cnpj'])}")
        st.write(f"**Endereço:** {transportador['endereco']}, {transportador['bairro']}, {transportador['cidade']}/{transportador['uf']}")
        if st.button("✅ Confirmar Transportador"):
            st.session_state.transportador = transportador
            st.session_state.etapa = "selecionar_veiculos"
            st.session_state.todos_veiculos = obter_veiculos_do_transportador(
                cnpj_prop, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
            )
            if not st.session_state.todos_veiculos:
                # Fallback: usar o veículo atual
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
        st.warning("Transportador não encontrado na base. Preencha manualmente:")
        with st.form("manual_transportador"):
            nome = st.text_input("Nome:", value=transportador_nome_sugerido or '')
            cpf = st.text_input("CPF/CNPJ:")
            endereco = st.text_input("Endereço:")
            bairro = st.text_input("Bairro:")
            cidade = st.text_input("Cidade:")
            uf = st.text_input("UF:")
            data_inclusao = st.text_input("Data de inclusão:")
            rntrc = st.text_input("RNTRC:")
            if st.form_submit_button("✅ Confirmar"):
                st.session_state.transportador = {
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
                    st.session_state.todos_veiculos = obter_veiculos_do_transportador(
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

# ============================================================
# ETAPA 3: SELECIONAR VEÍCULOS
# ============================================================
if st.session_state.etapa == "selecionar_veiculos":
    st.header("🚗 3. Selecionar Veículos do Transportador")
    st.write(f"Transportador: **{st.session_state.transportador['nome']}**")
    
    if st.session_state.todos_veiculos:
        # Exibir lista com checkboxes
        veiculos_opcoes = []
        for v in st.session_state.todos_veiculos:
            label = f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})"
            if v['mot_nome']:
                label += f" - Motorista: {v['mot_nome']}"
            veiculos_opcoes.append(label)
        
        selecionados = st.multiselect(
            "Selecione os veículos que farão parte do contrato:",
            options=veiculos_opcoes,
            default=veiculos_opcoes  # pré-seleciona todos
        )
        
        # Mapear seleção para objetos
        veiculos_selecionados = []
        for i, label in enumerate(veiculos_opcoes):
            if label in selecionados:
                veiculos_selecionados.append(st.session_state.todos_veiculos[i])
        
        if st.button("✅ Confirmar veículos selecionados"):
            st.session_state.veiculos_selecionados = veiculos_selecionados
            st.session_state.etapa = "selecionar_motoristas"
            st.rerun()
    else:
        st.warning("Nenhum veículo encontrado para este transportador.")
        if st.button("🔙 Voltar e buscar novamente"):
            st.session_state.etapa = "busca"
            st.rerun()

# ============================================================
# ETAPA 4: SELECIONAR MOTORISTAS
# ============================================================
if st.session_state.etapa == "selecionar_motoristas":
    st.header("👨‍✈️ 4. Selecionar Motoristas")
    
    # Coletar motoristas dos veículos selecionados
    motoristas_dos_veiculos = obter_motoristas_dos_veiculos(
        st.session_state.veiculos_selecionados, df_condutores, df_fornecedores
    )
    
    # Opções adicionais da base
    todos_condutores = obter_todos_condutores_da_base(df_condutores)
    
    # Criar lista de opções combinando motoristas dos veículos e todos os condutores
    opcoes_motoristas = {}
    for m in motoristas_dos_veiculos:
        opcoes_motoristas[m['cpf']] = m['nome']
    for c in todos_condutores:
        if c['cpf'] not in opcoes_motoristas:
            opcoes_motoristas[c['cpf']] = c['nome']
    
    # Converter para lista para multiselect
    lista_opcoes = [f"{cpf} - {nome}" for cpf, nome in opcoes_motoristas.items()]
    
    # Pré-selecionar os motoristas que já estão associados aos veículos selecionados
    default_selecionados = [f"{m['cpf']} - {m['nome']}" for m in motoristas_dos_veiculos if m['cpf']]
    
    selecionados = st.multiselect(
        "Selecione os motoristas que farão parte do contrato:",
        options=lista_opcoes,
        default=default_selecionados
    )
    
    # Converter seleção para lista de dicionários
    motoristas_selecionados = []
    for item in selecionados:
        cpf = item.split(' - ')[0]
        nome = item.split(' - ')[1]
        motoristas_selecionados.append({'cpf': cpf, 'nome': nome})
    
    if st.button("✅ Confirmar motoristas"):
        st.session_state.motoristas_selecionados = motoristas_selecionados
        st.session_state.etapa = "gerar_contrato"
        st.rerun()

# ============================================================
# ETAPA 5: GERAR CONTRATO
# ============================================================
if st.session_state.etapa == "gerar_contrato":
    st.header("📄 5. Gerar Contrato")
    
    transportador = st.session_state.transportador
    veiculos = st.session_state.veiculos_selecionados
    motoristas = st.session_state.motoristas_selecionados
    
    # Exibir resumo
    st.subheader("Resumo do contrato")
    st.write(f"**Transportador:** {transportador['nome']}")
    st.write(f"**CNPJ/CPF:** {formatar_cpf_cnpj(transportador['cpf_cnpj'])}")
    st.write(f"**Endereço:** {transportador['endereco']}, {transportador['bairro']}, {transportador['cidade']}/{transportador['uf']}")
    st.write(f"**Veículos selecionados:** {len(veiculos)}")
    for v in veiculos:
        st.write(f"- {v['placa']} - {v['marca']} {v['modelo']}")
    st.write(f"**Motoristas selecionados:** {len(motoristas)}")
    for m in motoristas:
        st.write(f"- {m['nome']} (CPF: {formatar_cpf(m['cpf'])})")
    
    # Botão para gerar
    if st.button("🚀 Gerar Contrato"):
        # Preparar dados para o template
        # (Aqui você deve mapear os campos do template conforme necessário)
        # Exemplo simples:
        context = {
            'transportador_nome': transportador['nome'],
            'transportador_cpf_cnpj': formatar_cpf_cnpj(transportador['cpf_cnpj']),
            'transportador_endereco': transportador['endereco'],
            'transportador_bairro': transportador['bairro'],
            'transportador_cidade': transportador['cidade'],
            'transportador_uf': transportador['uf'],
            'transportador_rntrc': transportador.get('rntrc', ''),
            'data_dia': datetime.now().day,
            'data_mes': meses[datetime.now().month],
            'data_ano': datetime.now().year,
            # Adicionar listas de veículos e motoristas conforme necessário
        }
        
        # Carregar template e gerar documento
        try:
            template = DocxTemplate(io.BytesIO(template_bytes))
            template.render(context)
            # Salvar em memória
            output = io.BytesIO()
            template.save(output)
            output.seek(0)
            
            st.success("✅ Contrato gerado com sucesso!")
            st.download_button(
                label="📥 Baixar Contrato",
                data=output,
                file_name=f"contrato_{transportador['nome']}_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Erro ao gerar contrato: {e}")
    
    if st.button("🔙 Voltar ao início"):
        # Resetar estado
        for key in ['etapa', 'veiculo_escolhido', 'transportador', 'todos_veiculos', 'veiculos_selecionados', 'motoristas_selecionados']:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()
