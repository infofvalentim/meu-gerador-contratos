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
# CONFIGURAÇÃO DOS ARQUIVOS (com nomes atualizados)
# ============================================================
FORNECEDORES = os.path.join('dados', 'fornecedores_.xlsx')
VEICULOS_ANTT = os.path.join('dados', 'veiculos_.xlsx')
CONDUTORES = os.path.join('dados', 'condutores.xlsx')
VEICULOS_COMPL = os.path.join('dados', 'veiculos.xlsx')
TEMPLATE = os.path.join('dados', 'template_contrato.docx')
PASTA_SAIDA = 'contratos_gerados'

if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

# ============================================================
# FUNÇÕES DE CARREGAMENTO
# ============================================================
@st.cache_data
def carregar_excel(caminho, header=4):
    try:
        return pd.read_excel(caminho, dtype=str, engine='openpyxl', header=header)
    except Exception as e:
        st.error(f"Erro ao carregar {caminho}: {e}")
        return None

@st.cache_data
def carregar_template(caminho):
    try:
        with open(caminho, "rb") as f:
            return f.read()
    except Exception as e:
        st.error(f"Erro ao carregar template: {e}")
        return None

def verificar_arquivos():
    faltantes = []
    for arquivo in [FORNECEDORES, VEICULOS_ANTT, CONDUTORES, VEICULOS_COMPL, TEMPLATE]:
        if not os.path.exists(arquivo):
            faltantes.append(os.path.basename(arquivo))
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
# FUNÇÕES DE NEGÓCIO
# ============================================================
COLUNAS_MOTORISTA = ['Motorista Cnpj', 'Motorista CPF', 'Condutor CPF', 'Motorista', 'CPF Motorista', 'Condutor']

def buscar_veiculo_por_placa(df_veic_antt, df_veic_compl, df_forn, df_cond, termo):
    """Busca veículos por placa (ou parte) e retorna lista com dados."""
    if not termo:
        return []
    mask = df_veic_antt['Placa'].str.upper().str.contains(termo.upper(), na=False)
    resultados = df_veic_antt[mask]
    veiculos = []
    for _, row in resultados.iterrows():
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
        prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
        prop_nome = ''
        prop_rntrc = ''
        if prop_cnpj:
            p = df_forn[df_forn['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
            if not p.empty:
                prop_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
                prop_rntrc = safe_str(p.iloc[0].get('Rntrc', ''))
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
            'prop_cnpj': prop_cnpj,
            'prop_nome': prop_nome,
            'prop_rntrc': prop_rntrc,
            'mot_cpf': mot_cpf,
            'mot_nome': mot_nome
        })
    return veiculos

def obter_veiculos_do_proprietario(cnpj_prop, df_veic_antt, df_veic_compl, df_cond, df_forn):
    """Retorna todos os veículos de um determinado CNPJ/CPF (proprietário)."""
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

def obter_motoristas_dos_veiculos(veiculos, df_condutores, df_fornecedores):
    """Extrai motoristas únicos a partir da lista de veículos."""
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
    """Retorna todos os condutores (da planilha condutores + fornecedores PF)."""
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

# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

# --- Verificar arquivos ---
faltantes = verificar_arquivos()
if faltantes:
    st.warning(f"⚠️ Arquivos não encontrados: {', '.join(faltantes)}")
    st.info("Certifique-se de que todos os arquivos estejam na pasta `dados/`.")
    st.stop()
else:
    st.success("✅ Todos os arquivos encontrados!")

# --- Carregar dados ---
df_fornecedores = carregar_excel(FORNECEDORES, header=4)
df_veiculos_antt = carregar_excel(VEICULOS_ANTT, header=4)
df_condutores = carregar_excel(CONDUTORES, header=4)
df_veiculos_compl = carregar_excel(VEICULOS_COMPL, header=4)
template_bytes = carregar_template(TEMPLATE)

if any(df is None for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]) or template_bytes is None:
    st.error("❌ Erro ao carregar um ou mais arquivos. Verifique os logs.")
    st.stop()

# Limpeza básica
for df in [df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl]:
    df.dropna(how='all', inplace=True)
    df.columns = df.columns.str.strip()

# ============================================================
# ESTADO DA SESSÃO
# ============================================================
if "etapa" not in st.session_state:
    st.session_state.etapa = "busca"
if "resultados_busca" not in st.session_state:
    st.session_state.resultados_busca = []
if "veiculo_escolhido" not in st.session_state:
    st.session_state.veiculo_escolhido = None
if "proprietario" not in st.session_state:
    st.session_state.proprietario = None
if "todos_veiculos" not in st.session_state:
    st.session_state.todos_veiculos = []
if "veiculos_selecionados" not in st.session_state:
    st.session_state.veiculos_selecionados = []
if "motoristas_disponiveis" not in st.session_state:
    st.session_state.motoristas_disponiveis = []
if "motoristas_selecionados" not in st.session_state:
    st.session_state.motoristas_selecionados = []
if "antt" not in st.session_state:
    st.session_state.antt = ""
if "serial" not in st.session_state:
    st.session_state.serial = ""
if "cep_manual" not in st.session_state:
    st.session_state.cep_manual = ""
if "numero_casa" not in st.session_state:
    st.session_state.numero_casa = ""
if "rg_prop" not in st.session_state:
    st.session_state.rg_prop = ""
if "estado_civil" not in st.session_state:
    st.session_state.estado_civil = ""

# ============================================================
# ETAPA 1: BUSCA POR PLACA
# ============================================================
if st.session_state.etapa == "busca":
    st.header("🔍 1. Buscar Veículo")

    termo_busca = st.text_input("Digite a placa (ou parte):", key="busca_placa")
    if st.button("Buscar", key="botao_buscar"):
        if termo_busca:
            st.session_state.resultados_busca = buscar_veiculo_por_placa(
                df_veiculos_antt, df_veiculos_compl, df_fornecedores, df_condutores, termo_busca
            )
        else:
            st.warning("Digite um termo para buscar.")

    if st.session_state.resultados_busca:
        st.success(f"✅ {len(st.session_state.resultados_busca)} veículo(s) encontrado(s).")
        opcoes = []
        for v in st.session_state.resultados_busca:
            desc = f"{v['marca']} {v['modelo']}".strip() or "Sem descrição"
            opcoes.append(f"{v['placa']} - {desc} ({v['ano']}) | Proprietário: {v['prop_nome'] or 'N/I'}")
        selecao = st.selectbox("Selecione o veículo principal:", opcoes, key="selecao_veiculo")
        if st.button("✅ Confirmar veículo", key="confirma_veiculo"):
            idx = opcoes.index(selecao)
            st.session_state.veiculo_escolhido = st.session_state.resultados_busca[idx]
            st.session_state.etapa = "confirmar_proprietario"
            st.rerun()
    elif termo_busca:
        st.warning("Nenhum veículo encontrado com esse termo.")

# ============================================================
# ETAPA 2: CONFIRMAR PROPRIETÁRIO (TRANSPORTADOR)
# ============================================================
if st.session_state.etapa == "confirmar_proprietario":
    st.header("👤 2. Confirmar Transportador (Proprietário)")

    veic = st.session_state.veiculo_escolhido
    prop_sugerido = None
    cnpj_prop = veic.get('prop_cnpj')
    if cnpj_prop:
        p = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == cnpj_prop]
        if not p.empty:
            row = p.iloc[0]
            prop_sugerido = {
                'nome': safe_str(row['Nome Fornecedor']),
                'cpf_cnpj': safe_str(row['Cnpj/Cpf']),
                'endereco': safe_str(row.get('Endereço', '')),
                'bairro': safe_str(row.get('Bairro', '')),
                'cidade': safe_str(row.get('Cidade', '')),
                'uf': safe_str(row.get('UF', '')),
                'data_inclusao': safe_str(row.get('Data Inclusão', '')),
                'rntrc': safe_str(row.get('Rntrc', ''))
            }

    if prop_sugerido:
        st.write(f"**Nome:** {prop_sugerido['nome']}")
        st.write(f"**CPF/CNPJ:** {formatar_cpf_cnpj(prop_sugerido['cpf_cnpj'])}")
        st.write(f"**Endereço:** {prop_sugerido['endereco']}, {prop_sugerido['bairro']}, {prop_sugerido['cidade']}/{prop_sugerido['uf']}")
        st.write(f"**RNTRC:** {prop_sugerido['rntrc']}")
        if st.button("✅ Confirmar este proprietário", key="confirma_prop"):
            st.session_state.proprietario = prop_sugerido
            st.session_state.etapa = "selecionar_veiculos"
            # Obter todos os veículos do proprietário
            st.session_state.todos_veiculos = obter_veiculos_do_proprietario(
                cnpj_prop, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
            )
            if not st.session_state.todos_veiculos:
                st.session_state.todos_veiculos = [veic]  # fallback
            st.rerun()
    else:
        st.warning("Proprietário não encontrado na base. Selecione manualmente:")

    # Seleção manual via lista de fornecedores
    with st.expander("🔄 Selecionar proprietário manualmente"):
        lista_forn = []
        for _, row in df_fornecedores.iterrows():
            nome = safe_str(row.get('Nome Fornecedor', ''))
            cpf = safe_str(row.get('Cnpj/Cpf', ''))
            if nome and cpf:
                lista_forn.append({
                    'nome': nome,
                    'cpf_cnpj': cpf,
                    'row': row
                })
        if lista_forn:
            opcoes_forn = [f"{f['nome']} - {formatar_cpf_cnpj(f['cpf_cnpj'])}" for f in lista_forn]
            escolha_forn = st.selectbox("Selecione um fornecedor:", opcoes_forn, key="manual_forn")
            if st.button("Usar este fornecedor", key="usa_forn"):
                idx = opcoes_forn.index(escolha_forn)
                row = lista_forn[idx]['row']
                prop_manual = {
                    'nome': safe_str(row['Nome Fornecedor']),
                    'cpf_cnpj': safe_str(row['Cnpj/Cpf']),
                    'endereco': safe_str(row.get('Endereço', '')),
                    'bairro': safe_str(row.get('Bairro', '')),
                    'cidade': safe_str(row.get('Cidade', '')),
                    'uf': safe_str(row.get('UF', '')),
                    'data_inclusao': safe_str(row.get('Data Inclusão', '')),
                    'rntrc': safe_str(row.get('Rntrc', ''))
                }
                st.session_state.proprietario = prop_manual
                st.session_state.etapa = "selecionar_veiculos"
                cnpj_manual = limpar_cpf_cnpj(prop_manual['cpf_cnpj'])
                st.session_state.todos_veiculos = obter_veiculos_do_proprietario(
                    cnpj_manual, df_veiculos_antt, df_veiculos_compl, df_condutores, df_fornecedores
                )
                if not st.session_state.todos_veiculos:
                    st.session_state.todos_veiculos = [veic]
                st.rerun()
        else:
            st.info("Nenhum fornecedor cadastrado.")

# ============================================================
# ETAPA 3: SELECIONAR VEÍCULOS (com adição de extras)
# ============================================================
if st.session_state.etapa == "selecionar_veiculos":
    st.header("🚛 3. Selecionar Veículos para o Contrato")

    if not st.session_state.todos_veiculos:
        st.warning("Nenhum veículo disponível. Volte e confirme o proprietário.")
        if st.button("Voltar"):
            st.session_state.etapa = "busca"
            st.rerun()
        st.stop()

    # Exibir lista atual de veículos
    opcoes_veiculos = [f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})" for v in st.session_state.todos_veiculos]
    selecionados = st.multiselect(
        "Selecione os veículos que farão parte do contrato (o primeiro será o principal):",
        opcoes_veiculos,
        default=opcoes_veiculos[:1]
    )
    if selecionados:
        st.session_state.veiculos_selecionados = [st.session_state.todos_veiculos[opcoes_veiculos.index(s)] for s in selecionados]
    else:
        st.warning("Selecione pelo menos um veículo.")

    # Adicionar veículo extra da base completa
    with st.expander("➕ Adicionar veículo extra da base completa"):
        modo_extra = st.radio("Modo de busca:", ["Por placa", "Listar todos (filtrar)"], key="modo_extra")

        if modo_extra == "Por placa":
            placa_extra = st.text_input("Digite a placa (ou parte):", key="placa_extra")
            if st.button("Buscar veículo extra", key="buscar_extra"):
                if placa_extra:
                    extras = buscar_veiculo_por_placa(
                        df_veiculos_antt, df_veiculos_compl, df_fornecedores, df_condutores, placa_extra
                    )
                    if extras:
                        st.session_state._extras_temp = extras
                    else:
                        st.warning("Nenhum veículo encontrado.")
            if "_extras_temp" in st.session_state and st.session_state._extras_temp:
                extras = st.session_state._extras_temp
                opcoes_extra = [f"{v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})" for v in extras]
                escolha_extra = st.selectbox("Selecione o veículo:", opcoes_extra, key="escolha_extra")
                if st.button("➕ Adicionar este veículo", key="add_extra"):
                    idx = opcoes_extra.index(escolha_extra)
                    novo = extras[idx]
                    # Verificar duplicata
                    if any(limpar_placa(v['placa']) == limpar_placa(novo['placa']) for v in st.session_state.todos_veiculos):
                        st.warning("Veículo já está na lista.")
                    else:
                        st.session_state.todos_veiculos.append(novo)
                        st.success(f"✅ {novo['placa']} adicionado!")
                        st.rerun()

        else:  # Listar todos com filtro
            filtro = st.text_input("Filtrar por placa ou modelo (opcional):", key="filtro_veic")
            # Construir lista completa da base
            todos_da_base = []
            for _, row in df_veiculos_antt.iterrows():
                placa = safe_str(row['Placa'])
                marca = safe_str(row.get('Marca', ''))
                modelo = safe_str(row.get('Modelo', ''))
                ano = safe_str(row.get('Ano', ''))
                placa_limpa = limpar_placa(placa)
                n_equip = ''
                if not df_veiculos_compl.empty:
                    vc = df_veiculos_compl[df_veiculos_compl['Placa'].apply(limpar_placa) == placa_limpa]
                    if not vc.empty:
                        rv = vc.iloc[0]
                        marca = safe_str(rv.get('Marca', '')) or marca
                        modelo = safe_str(rv.get('Modelo', '')) or modelo
                        n_equip = safe_str(rv.get('Nº Equipamento', ''))
                # Obter proprietário para exibição
                prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
                prop_nome = ''
                if prop_cnpj:
                    p = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
                    if not p.empty:
                        prop_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
                desc = f"{placa} - {marca} {modelo} ({ano}) - Proprietário: {prop_nome or 'N/I'}"
                if filtro:
                    if filtro.upper() in placa.upper() or filtro.upper() in marca.upper() or filtro.upper() in modelo.upper():
                        todos_da_base.append({'desc': desc, 'veiculo': {'placa': placa, 'marca': marca, 'modelo': modelo, 'ano': ano, 'renavan': safe_str(row.get('Renavan', '')), 'n_equipamento': n_equip, 'mot_cpf': '', 'mot_nome': ''}})
                else:
                    todos_da_base.append({'desc': desc, 'veiculo': {'placa': placa, 'marca': marca, 'modelo': modelo, 'ano': ano, 'renavan': safe_str(row.get('Renavan', '')), 'n_equipamento': n_equip, 'mot_cpf': '', 'mot_nome': ''}})
            if todos_da_base:
                st.write(f"Total: {len(todos_da_base)} veículos")
                opcoes_lista = [item['desc'] for item in todos_da_base[:50]]
                escolha_lista = st.selectbox("Selecione um veículo para adicionar:", opcoes_lista, key="escolha_lista")
                if st.button("➕ Adicionar da lista", key="add_lista"):
                    idx = opcoes_lista.index(escolha_lista)
                    novo = todos_da_base[idx]['veiculo']
                    if any(limpar_placa(v['placa']) == limpar_placa(novo['placa']) for v in st.session_state.todos_veiculos):
                        st.warning("Veículo já está na lista.")
                    else:
                        st.session_state.todos_veiculos.append(novo)
                        st.success(f"✅ {novo['placa']} adicionado!")
                        st.rerun()
            else:
                st.info("Nenhum veículo encontrado com esse filtro.")

    if st.session_state.veiculos_selecionados:
        if st.button("➡️ Próximo (Motoristas)", key="prox_motoristas"):
            # Preparar motoristas disponíveis
            motoristas_vinculados = obter_motoristas_dos_veiculos(st.session_state.todos_veiculos, df_condutores, df_fornecedores)
            if not motoristas_vinculados:
                # Buscar todos os condutores da base
                st.info("Nenhum motorista vinculado aos veículos. Buscando todos os condutores da base...")
                motoristas_vinculados = obter_todos_condutores_da_base(df_condutores, df_fornecedores)
            st.session_state.motoristas_disponiveis = motoristas_vinculados
            st.session_state.etapa = "selecionar_motoristas"
            st.rerun()

# ============================================================
# ETAPA 4: SELECIONAR MOTORISTAS
# ============================================================
if st.session_state.etapa == "selecionar_motoristas":
    st.header("👤 4. Selecionar Motoristas")

    # Exibir motoristas disponíveis
    if st.session_state.motoristas_disponiveis:
        opcoes_mot = [f"{m['nome']} (CPF: {formatar_cpf(m['cpf'])})" for m in st.session_state.motoristas_disponiveis]
        selecionados_mot = st.multiselect(
            "Selecione os motoristas para o contrato (o primeiro será o principal):",
            opcoes_mot,
            default=opcoes_mot[:1] if opcoes_mot else []
        )
        if selecionados_mot:
            st.session_state.motoristas_selecionados = [st.session_state.motoristas_disponiveis[opcoes_mot.index(s)] for s in selecionados_mot]
        else:
            st.warning("Selecione pelo menos um motorista.")
    else:
        st.warning("Nenhum motorista disponível. Adicione manualmente ou da base.")

    # Adicionar motoristas extras
    with st.expander("➕ Adicionar motorista extra"):
        st.subheader("Adicionar da base geral")
        if st.button("Carregar condutores da base"):
            todos_cond = obter_todos_condutores_da_base(df_condutores, df_fornecedores)
            # Filtrar os que já estão na lista
            cpfs_existentes = set(limpar_cpf_cnpj(m['cpf']) for m in st.session_state.motoristas_disponiveis)
            novos = [c for c in todos_cond if limpar_cpf_cnpj(c['cpf']) not in cpfs_existentes]
            if novos:
                st.session_state._novos_condutores = novos
                st.success(f"{len(novos)} novos condutores encontrados.")
            else:
                st.info("Todos os condutores já estão na lista.")

        if "_novos_condutores" in st.session_state and st.session_state._novos_condutores:
            novos = st.session_state._novos_condutores
            opcoes_novos = [f"{m['nome']} (CPF: {formatar_cpf(m['cpf'])})" for m in novos]
            selecionar_novos = st.multiselect("Selecione os que deseja adicionar:", opcoes_novos)
            if st.button("➕ Adicionar selecionados", key="add_novos"):
                indices = [opcoes_novos.index(s) for s in selecionar_novos]
                for i in indices:
                    st.session_state.motoristas_disponiveis.append(novos[i])
                st.success(f"{len(indices)} motorista(s) adicionado(s).")
                st.rerun()

        st.subheader("Digitar manualmente")
        with st.form("form_motorista_manual"):
            nome = st.text_input("Nome:")
            cpf = st.text_input("CPF (somente números):")
            cnh = st.text_input("CNH:")
            rg = st.text_input("RG:")
            if st.form_submit_button("➕ Adicionar motorista manual"):
                if nome and cpf:
                    novo = {'nome': nome, 'cpf': cpf, 'cnh': cnh, 'rg': rg}
                    st.session_state.motoristas_disponiveis.append(novo)
                    st.success("Motorista adicionado!")
                    st.rerun()
                else:
                    st.error("Nome e CPF são obrigatórios.")

    if st.session_state.motoristas_selecionados:
        if st.button("➡️ Próximo (Gerar Contrato)", key="prox_contrato"):
            st.session_state.etapa = "gerar_contrato"
            st.rerun()

# ============================================================
# ETAPA 5: GERAR CONTRATO (com opção de número da casa)
# ============================================================
if st.session_state.etapa == "gerar_contrato":
    st.header("📄 5. Gerar Contrato")

    proprietario = st.session_state.proprietario
    veiculos_selecionados = st.session_state.veiculos_selecionados
    motoristas_selecionados = st.session_state.motoristas_selecionados

    if not proprietario or not veiculos_selecionados or not motoristas_selecionados:
        st.error("Dados incompletos. Volte e preencha todas as etapas.")
        st.stop()

    st.subheader("📝 Dados Complementares")

    col1, col2 = st.columns(2)
    with col1:
        antt = proprietario.get('rntrc', '')
        if not antt:
            antt = buscar_antt(proprietario, veiculos_selecionados, df_veiculos_antt, df_veiculos_compl)
        if not antt:
            antt = st.text_input("ANTT/RNTRC (não encontrado):", value="", key="antt_manual")
        else:
            st.write(f"✅ ANTT/RNTRC encontrado: {antt}")

        serial = ''
        for v in veiculos_selecionados:
            if v.get('n_equipamento') and v['n_equipamento'].lower() != 'nan':
                serial = v['n_equipamento']
                break
        if not serial:
            serial = st.text_input("Serial do rastreador:", value="A DEFINIR", key="serial_manual")
        else:
            st.write(f"✅ Serial encontrado: {serial}")

        valor = st.text_input("Valor por entrega (R$):", value="X (a definir)", key="valor")

    with col2:
        is_pf = len(limpar_cpf_cnpj(proprietario['cpf_cnpj'])) == 11
        if not is_pf:
            cidade = proprietario.get('cidade', '')
            uf = proprietario.get('uf', '')
            endereco_raw = proprietario.get('endereco', '')
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

    # === OPÇÃO DE NÚMERO DA CASA ===
    st.subheader("🏠 Endereço - Número da Casa")
    numero_casa = st.text_input(
        "Número da residência (deixe em branco se não houver):",
        value=st.session_state.get("numero_casa", ""),
        key="num_casa_input"
    )
    if numero_casa:
        st.session_state.numero_casa = numero_casa
    else:
        st.session_state.numero_casa = ""

    if is_pf:
        st.subheader("👤 Dados da Pessoa Física (Proprietário)")
        col3, col4 = st.columns(2)
        with col3:
            rg_prop = buscar_rg_proprietario(proprietario['cpf_cnpj'], df_condutores, df_fornecedores)
            if not rg_prop:
                rg_prop = st.text_input("RG do proprietário:", value=st.session_state.get("rg_prop", ""), key="rg_prop")
            else:
                st.write(f"✅ RG encontrado: {rg_prop}")
        with col4:
            estado_civil = st.text_input("Estado civil:", value=st.session_state.get("estado_civil", ""), key="estado_civil")
    else:
        rg_prop = ""
        estado_civil = ""

    # Mostrar resumo
    st.subheader("📋 Resumo do Contrato")
    st.write(f"**Transportador:** {proprietario['nome']} ({formatar_cpf_cnpj(proprietario['cpf_cnpj'])})")
    st.write(f"**ANTT/RNTRC:** {antt}")
    st.write("**Veículos selecionados:**")
    for v in veiculos_selecionados:
        st.write(f"  - {v['placa']} - {v['marca']} {v['modelo']} ({v['ano']})")
    st.write("**Motoristas selecionados:**")
    for m in motoristas_selecionados:
        st.write(f"  - {m['nome']} (CPF: {formatar_cpf(m['cpf'])}, CNH: {m['cnh']})")

    if st.button("🚀 Gerar Contrato", key="gerar"):
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
        ] or [{'nome': '', 'cpf': '', 'cnh': ''}]

        veiculos_lista = [
            {
                'modelo': f"{v.get('marca','')} {v.get('modelo','')}".strip(),
                'ano': v.get('ano', ''),
                'renavam': v.get('renavan', ''),
            }
            for v in veiculos_selecionados
        ] or [{'modelo': '', 'ano': '', 'renavam': ''}]

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
            rg_final = rg_prop or buscar_rg_proprietario(proprietario['cpf_cnpj'], df_condutores, df_fornecedores) or ''
            contexto.update({
                'nome_contratado_pf': proprietario['nome'],
                'estado_civil': estado_civil,
                'rg_contratado': rg_final,
                'cpf_contratado': formatar_cpf(proprietario['cpf_cnpj']),
                'numero_da_casa': st.session_state.numero_casa,
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
                'razao_social_pj': proprietario['nome'],
                'cnpj_pj': formatar_cnpj(proprietario['cpf_cnpj']),
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
            if key not in ["df_fornecedores", "df_veiculos_antt", "df_condutores", "df_veiculos_compl", "template_bytes"]:
                del st.session_state[key]
        st.rerun()
