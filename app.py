import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
from datetime import datetime

st.set_page_config(page_title="Gerador de Contratos", layout="wide")
st.title("🚛 Gerador de Contratos de Transporte")

# ------------------------------------------------------------
# FUNÇÕES AUXILIARES (do seu script)
# ------------------------------------------------------------
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

# ------------------------------------------------------------
# INTERFACE
# ------------------------------------------------------------
st.sidebar.header("📤 Upload dos Arquivos")
fornecedores_file = st.sidebar.file_uploader("Fornecedores", type="xlsx")
veiculos_antt_file = st.sidebar.file_uploader("Veículos ANTT", type="xlsx")
condutores_file = st.sidebar.file_uploader("Condutores", type="xlsx")
veiculos_compl_file = st.sidebar.file_uploader("Veículos Complementares", type="xlsx")
template_file = st.sidebar.file_uploader("Template do Contrato (DOCX)", type="docx")

if not (fornecedores_file and veiculos_antt_file and condutores_file and veiculos_compl_file and template_file):
    st.info("📂 Faça o upload de todos os arquivos na barra lateral.")
    st.stop()

# Ler planilhas
@st.cache_data
def carregar_dados(forn, veic_antt, cond, veic_compl):
    df_forn = pd.read_excel(forn, dtype=str, engine='openpyxl', header=4)
    df_veic = pd.read_excel(veic_antt, dtype=str, engine='openpyxl', header=4)
    df_cond = pd.read_excel(cond, dtype=str, engine='openpyxl', header=4)
    df_compl = pd.read_excel(veic_compl, dtype=str, engine='openpyxl', header=4)
    for df in [df_forn, df_veic, df_cond, df_compl]:
        df.dropna(how='all', inplace=True)
        df.columns = df.columns.str.strip()
    return df_forn, df_veic, df_cond, df_compl

df_fornecedores, df_veiculos_antt, df_condutores, df_veiculos_compl = carregar_dados(
    fornecedores_file, veiculos_antt_file, condutores_file, veiculos_compl_file
)

st.success("✅ Planilhas carregadas com sucesso!")

# ------------------------------------------------------------
# BUSCA POR PLACA
# ------------------------------------------------------------
st.header("🔍 Buscar Veículo")
placa_busca = st.text_input("Digite a placa (ou parte)").strip().upper()

veiculo_escolhido = None
if placa_busca:
    mask = df_veiculos_antt['Placa'].str.upper().str.contains(placa_busca, na=False)
    resultados = df_veiculos_antt[mask]
    if resultados.empty:
        st.warning("Nenhum veículo encontrado.")
    else:
        opcoes = []
        veiculos = []
        for _, row in resultados.iterrows():
            placa = safe_str(row['Placa'])
            marca = safe_str(row.get('Marca', ''))
            modelo = safe_str(row.get('Modelo', ''))
            ano = safe_str(row.get('Ano', ''))
            prop_cnpj = limpar_cpf_cnpj(row.get('Proprietário Cnpj', ''))
            prop_nome = ''
            if prop_cnpj:
                p = df_fornecedores[df_fornecedores['Cnpj/Cpf'].apply(limpar_cpf_cnpj) == prop_cnpj]
                if not p.empty:
                    prop_nome = safe_str(p.iloc[0]['Nome Fornecedor'])
            desc = f"{placa} - {marca} {modelo} ({ano}) - {prop_nome or 'N/I'}"
            opcoes.append(desc)
            veiculos.append({
                'placa': placa,
                'marca': marca,
                'modelo': modelo,
                'ano': ano,
                'prop_cnpj': prop_cnpj,
                'prop_nome': prop_nome,
                'row': row
            })
        selecionado = st.selectbox("Selecione o veículo:", opcoes)
        idx = opcoes.index(selecionado)
        veiculo_escolhido = veiculos[idx]
        st.success(f"✅ Veículo selecionado: {veiculo_escolhido['placa']}")

if not veiculo_escolhido:
    st.stop()

# ------------------------------------------------------------
# PROPRIETÁRIO
# ------------------------------------------------------------
st.header("👤 Transportador")
cnpj_prop = veiculo_escolhido['prop_cnpj']
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
else:
    st.warning("Proprietário não encontrado. Preencha manualmente:")
    proprietario = {
        'nome': st.text_input("Nome do proprietário:"),
        'cpf_cnpj': st.text_input("CPF/CNPJ:"),
        'endereco': st.text_input("Endereço:"),
        'bairro': st.text_input("Bairro:"),
        'cidade': st.text_input("Cidade:"),
        'uf': st.text_input("UF:"),
        'data_inclusao': st.text_input("Data de inclusão:"),
        'rntrc': st.text_input("RNTRC:")
    }

# ------------------------------------------------------------
# VEÍCULOS E MOTORISTAS
# ------------------------------------------------------------
st.header("🚛 Veículos e Motoristas")

# Veículos do proprietário
mask_prop = df_veiculos_antt['Proprietário Cnpj'].apply(limpar_cpf_cnpj) == limpar_cpf_cnpj(proprietario['cpf_cnpj'])
veiculos_prop = df_veiculos_antt[mask_prop]
if veiculos_prop.empty:
    st.info("Nenhum outro veículo encontrado para este proprietário.")
    veiculos_selecionados = [veiculo_escolhido]
else:
    opcoes_veic = []
    veiculos_lista = []
    for _, row in veiculos_prop.iterrows():
        placa = safe_str(row['Placa'])
        marca = safe_str(row.get('Marca', ''))
        modelo = safe_str(row.get('Modelo', ''))
        ano = safe_str(row.get('Ano', ''))
        desc = f"{placa} - {marca} {modelo} ({ano})"
        opcoes_veic.append(desc)
        veiculos_lista.append({
            'placa': placa,
            'marca': marca,
            'modelo': modelo,
            'ano': ano,
            'renavan': safe_str(row.get('Renavan', ''))
        })
    selecionados = st.multiselect("Selecione os veículos para o contrato (o primeiro será o principal):", opcoes_veic, default=opcoes_veic[:1])
    veiculos_selecionados = [veiculos_lista[opcoes_veic.index(s)] for s in selecionados]

if not veiculos_selecionados:
    st.warning("Selecione pelo menos um veículo.")
    st.stop()

# Motoristas
motoristas = []
# Buscar motoristas dos veículos selecionados
for v in veiculos_selecionados:
    placa_limpa = limpar_placa(v['placa'])
    mask_veic = df_veiculos_antt['Placa'].apply(limpar_placa) == placa_limpa
    if mask_veic.any():
        row = df_veiculos_antt[mask_veic].iloc[0]
        mot_cpf = limpar_cpf_cnpj(row.get('Motorista Cnpj', ''))
        if mot_cpf:
            m = df_condutores[df_condutores['CPF N°'].apply(limpar_cpf_cnpj) == mot_cpf]
            if not m.empty:
                motoristas.append({
                    'nome': safe_str(m.iloc[0]['Nome']),
                    'cpf': mot_cpf,
                    'cnh': safe_str(m.iloc[0].get('CNH N°', ''))
                })

# Se não encontrou, buscar da base geral
if not motoristas:
    # Todos os condutores
    for _, row in df_condutores.iterrows():
        nome = safe_str(row.get('Nome', ''))
        cpf = safe_str(row.get('CPF N°', ''))
        if nome and cpf:
            motoristas.append({
                'nome': nome,
                'cpf': limpar_cpf_cnpj(cpf),
                'cnh': safe_str(row.get('CNH N°', ''))
            })

if not motoristas:
    st.info("Nenhum motorista encontrado. Digite manualmente:")
    nome_mot = st.text_input("Nome do motorista:")
    cpf_mot = st.text_input("CPF:")
    cnh_mot = st.text_input("CNH:")
    motoristas = [{'nome': nome_mot, 'cpf': cpf_mot, 'cnh': cnh_mot}] if nome_mot else []

if motoristas:
    opcoes_mot = [f"{m['nome']} (CPF: {formatar_cpf(m['cpf'])})" for m in motoristas]
    mot_selecionados = st.multiselect("Selecione os motoristas para o contrato:", opcoes_mot, default=opcoes_mot[:1])
    motoristas_selecionados = [motoristas[opcoes_mot.index(s)] for s in mot_selecionados]
else:
    motoristas_selecionados = []

if not motoristas_selecionados:
    st.warning("Selecione pelo menos um motorista.")
    st.stop()

# ------------------------------------------------------------
# DADOS ADICIONAIS
# ------------------------------------------------------------
st.header("📄 Dados Complementares")
serial = st.text_input("Serial do rastreador:", value="A DEFINIR")
valor = st.text_input("Valor por entrega (R$):", value="X (a definir)")
antt = proprietario.get('rntrc', '')
if not antt or antt.lower() == 'nan':
    antt = st.text_input("ANTT/RNTRC:")

is_pf = len(limpar_cpf_cnpj(proprietario['cpf_cnpj'])) == 11
if is_pf:
    numero_casa = st.text_input("Número da residência:", value="")
    rg_prop = st.text_input("RG do proprietário:", value="")
    estado_civil = st.text_input("Estado civil:", value="")
else:
    numero_casa = ""
    rg_prop = ""
    estado_civil = ""

# ------------------------------------------------------------
# GERAR CONTRATO
# ------------------------------------------------------------
if st.button("🚀 Gerar Contrato"):
    try:
        # Preparar contexto
        dia_cad, mes_cad, ano_cad = formatar_data_cadastro(proprietario.get('data_inclusao', ''))
        cidade = proprietario.get('cidade', '')
        uf = proprietario.get('uf', '')
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
            rua, num, comp = extrair_endereco_pj(endereco_raw)
            if not rua: rua = endereco_raw
            if not num: num = 'S/N'
            if not comp: comp = bairro
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
                'cep_pj_1': '00000-000',
            })

        # Renderizar
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
        st.error(f"❌ Erro: {e}")