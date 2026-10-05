import streamlit as st
import pandas as pd
import database as db
import io
import re
import os
import base64
import streamlit.components.v1 as components

# Configuração da Página
st.set_page_config(
    page_title="CRM Call Center - Migração",
    page_icon="📞",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Declarar componente customizado de colar imagem (Ctrl+V)
PASTE_COMPONENT_PATH = os.path.join(os.path.dirname(__file__), "paste_component")
image_paste_component = components.declare_component("image_paste_box", path=PASTE_COMPONENT_PATH)

# Estilização CSS Customizada
st.markdown("""
<style>
    .main-header {
        font-size: 1.8rem;
        color: #0056b3;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .badge-gpon {
        background-color: #28a745;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: bold;
        display: inline-block;
    }
    .badge-no-gpon {
        background-color: #dc3545;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: bold;
        display: inline-block;
    }
    .badge-device {
        background-color: #17a2b8;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: bold;
        display: inline-block;
    }
    .card-box {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stButton>button {
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Função auxiliar para extrair texto de forma 100% segura contra valores nulos (nan)
def get_str(data_dict, key, default="N/I"):
    val = data_dict.get(key, "")
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "<na>", ""]:
        return default
    return val_str

# Carregar dados do banco/excel
df = db.load_dataset()
total_clientes = len(df)

# Inicialização dos estados de sessão
if "current_index" not in st.session_state:
    st.session_state.current_index = 0

if "consultor" not in st.session_state:
    st.session_state.consultor = "Consultor 1"

# Callback para quando o usuário altera o seletor na sidebar
def on_selectbox_change():
    if "select_client_widget" in st.session_state:
        st.session_state.current_index = st.session_state.select_client_widget

# Sidebar
st.sidebar.image("https://img.icons8.com/color/96/000000/headset.png", width=56)
st.sidebar.title(" Painel do Consultor")

consultor_input = st.sidebar.text_input("Seu Nome / ID:", value=st.session_state.consultor)
st.session_state.consultor = consultor_input

st.sidebar.divider()

# Progresso de Atendimentos
atendidos = (df["STATUS_CHAMADA"] != "Pendente").sum()
progresso = atendidos / total_clientes if total_clientes > 0 else 0

st.sidebar.subheader("📈 Progresso da Carteira")
st.sidebar.progress(progresso)
st.sidebar.write(f"**{atendidos}** de **{total_clientes}** clientes atendidos ({int(progresso*100)}%)")

st.sidebar.divider()

# Filtro e Navegação Direta
st.sidebar.subheader("🔍 Localizar Cliente")
modo_view = st.sidebar.radio("Filtrar lista por:", ["Todos", "Pendentes", "Atendidos"])

if modo_view == "Pendentes":
    lista_indices = df[df["STATUS_CHAMADA"] == "Pendente"].index.tolist()
elif modo_view == "Atendidos":
    lista_indices = df[df["STATUS_CHAMADA"] != "Pendente"].index.tolist()
else:
    lista_indices = df.index.tolist()

if not lista_indices:
    st.sidebar.warning("Nenhum cliente encontrado no filtro selecionado.")
    lista_indices = df.index.tolist()

# Garantir que o index atual seja válido dentro da lista_indices
if st.session_state.current_index not in lista_indices:
    st.session_state.current_index = lista_indices[0]

# Posição atual no filtro
selected_pos = lista_indices.index(st.session_state.current_index) if st.session_state.current_index in lista_indices else 0

# Seletor direto de clientes na Sidebar
st.sidebar.selectbox(
    "Ir direto para cliente:",
    options=lista_indices,
    index=selected_pos,
    format_func=lambda i: f"#{i+1} - {get_str(df.iloc[i].to_dict(), 'CLIENTE')[:25]}...",
    key="select_client_widget",
    on_change=on_selectbox_change
)

st.sidebar.divider()

# RECURSO: UPLOAD DE NOVAS PLANILHAS NA SIDEBAR
st.sidebar.subheader("📥 Importar Nova Planilha")
uploaded_sheet = st.sidebar.file_uploader(
    "Subir arquivo Excel (.xlsx) ou CSV:",
    type=["xlsx", "xls", "csv"],
    key="new_dataset_uploader"
)
if uploaded_sheet is not None:
    if st.sidebar.button("🔄 Atualizar Base de Dados", type="primary", use_container_width=True):
        file_bytes = uploaded_sheet.read()
        num_rows = db.import_new_dataset(file_bytes, uploaded_sheet.name)
        st.sidebar.success(f"✅ Planilha carregada! {num_rows} clientes importados.")
        st.session_state.current_index = 0
        st.rerun()

st.sidebar.divider()

# Exportar Relatório
st.sidebar.subheader("📊 Exportar Relatório")
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_export = db.export_results()
    df_export.to_excel(writer, index=False, sheet_name="Atendimentos")

st.sidebar.download_button(
    label="📥 Baixar Excel Completo",
    data=buffer.getvalue(),
    file_name=f"relatorio_atendimentos_migra_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# === CORPO PRINCIPAL DO APP ===

current_row_idx = st.session_state.current_index
row_dict = df.iloc[current_row_idx].to_dict()

cliente_nome = get_str(row_dict, "CLIENTE")
cnpj_val = get_str(row_dict, "CNPJ_CLIENTE")
contato_nome = get_str(row_dict, "CONTATO")
tel_raw = get_str(row_dict, "NR_TELEFONE", "")
municipio_val = get_str(row_dict, "DS_MUNICIPIO")
uf_val = get_str(row_dict, "UF")
endereco_val = get_str(row_dict, "NR_ENDERECO")
numero_val = get_str(row_dict, "Nº", "")
cep_val = get_str(row_dict, "NR_CEP", "")
plano_val = get_str(row_dict, "PLANO")
linhas_val = get_str(row_dict, "QTDE DE LINHAS ")
tipo_rede_val = get_str(row_dict, "TIPO_REDE", "SEM GPON").upper()
cobertura_val = get_str(row_dict, "TEM_COBERTURA_BANDA_LARGA", "Não")
mancha_val = get_str(row_dict, "TA NA MANCHA ", "0")
marca_val = get_str(row_dict, "APARELHO_TRAFEGO_MARCA", "")
modelo_val = get_str(row_dict, "APARELHO_TRAFEGO_MODELO", "")
recomendacao_val = get_str(row_dict, "RECOMENDACAO_APARELHO_LINHA", "Sem recomendação específica")
status_atual = get_str(row_dict, "STATUS_CHAMADA", "Pendente")
saved_atendeu = get_str(row_dict, "ATENDEU", "Não Registrado")
saved_obs = get_str(row_dict, "INTERACAO_CONSULTOR", "")
proposal_notes_saved = get_str(row_dict, "PROPOSAL_NOTES", "")

# Header do Cliente
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"<div class='main-header'>📞 Cliente #{current_row_idx + 1} de {total_clientes} - {cliente_nome}</div>", unsafe_allow_html=True)
with col_h2:
    if status_atual == "Pendente":
        st.warning(f"Status: **{status_atual}**")
    else:
        st.success(f"Status: **{status_atual}**")

st.divider()

# === BLOCO 1: DADOS DO CLIENTE (EMPRESA, CONTATO, GPON, APARELHOS) ===
col_info1, col_info2 = st.columns([1, 1])

with col_info1:
    st.markdown("####  Empresa & Contato")
    st.markdown(f"**CNPJ:** `{cnpj_val}`")
    st.markdown(f"**Contato:** {contato_nome}")
    
    # Tratamento de Telefone e WhatsApp
    tel_clean = re.sub(r'\D', '', tel_raw.replace('.0', ''))
    st.markdown(f"**Telefone:** `{tel_raw if tel_raw else 'N/I'}`")
    if tel_clean:
        tel_wa = tel_clean if tel_clean.startswith("55") else "55" + tel_clean
        st.markdown(f"[📱 **Iniciar conversa no WhatsApp**](https://wa.me/{tel_wa})", unsafe_allow_html=True)

    st.markdown(f"**Endereço:** {endereco_val}, Nº {numero_val} - {municipio_val}/{uf_val}")
    st.markdown(f"**Plano Atual:** `{plano_val}` ({linhas_val} linhas)")

with col_info2:
    st.markdown("#### ⚡ Viabilidade de Rede & Aparelhos")
    
    tem_gpon = "GPON" in tipo_rede_val or cobertura_val.lower() in ["sim", "1"]
    if tem_gpon:
        st.markdown(f"**Rede:** <span class='badge-gpon'>{tipo_rede_val} (Com Viabilidade)</span>", unsafe_allow_html=True)
    else:
        st.markdown(f"**Rede:** <span class='badge-no-gpon'>{tipo_rede_val if tipo_rede_val != 'N/I' else 'SEM GPON'} (Sem Cobertura)</span>", unsafe_allow_html=True)

    aparelho_trafego = f"{marca_val} {modelo_val}".strip()
    st.markdown(f"**Aparelho em Tráfego:** {aparelho_trafego if aparelho_trafego else 'Não informado'}")
    st.markdown(f"**Recomendação de Aparelho:** <span class='badge-device'>{recomendacao_val}</span>", unsafe_allow_html=True)

st.divider()

# === BLOCO 2: PAINEL DA PROPOSTA (LADO A LADO: NOTAS/COLAR DA ESQUERDA vs PREVIEW DA IMAGEM NA DIREITA) ===
st.markdown("### 📋 Painel da Proposta (Pré-Atendimento)")

col_prop_left, col_prop_right = st.columns([1, 1])

with col_prop_left:
    st.markdown("##### 📄 Notas e Condições da Proposta")
    prop_notes_input = st.text_area(
        "Anote propostas, valores ou condições negociadas:",
        value=proposal_notes_saved,
        height=100,
        key=f"prop_notes_{current_row_idx}",
        placeholder="Ex: Oferta de Migração 100GB por R$ 99/mês + aparelho Galaxy A17 grátis."
    )
    if st.button("💾 Salvar Notas da Proposta", key=f"btn_save_prop_notes_{current_row_idx}"):
        db.save_proposal(current_row_idx, prop_notes_input, cnpj=cnpj_val)
        st.toast("✅ Notas da proposta salvas!", icon="💾")
        st.rerun()

    st.markdown("---")
    st.markdown("##### 📸 Anexar / Colar Print da Proposta")
    
    # Colar via Ctrl+V
    st.caption("1. Clique na caixa abaixo e aperte **CTRL + V** (PrintScreen / Win+Shift+S):")
    pasted_b64 = image_paste_component(key=f"paste_box_{current_row_idx}")
    
    if pasted_b64 and isinstance(pasted_b64, str) and pasted_b64.startswith("data:image"):
        try:
            header, b64_data = pasted_b64.split(",", 1)
            img_bytes = base64.b64decode(b64_data)
            db.save_proposal(current_row_idx, prop_notes_input, img_bytes, "print_colado.png", cnpj=cnpj_val)
            st.toast("✅ Print colado com sucesso!", icon="📸")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao colar imagem: {e}")

    # Anexo de arquivo
    with st.expander("Ou selecione um arquivo do computador"):
        uploaded_file = st.file_uploader(
            "Escolha o arquivo de imagem:",
            type=["png", "jpg", "jpeg", "webp"],
            key=f"prop_img_uploader_{current_row_idx}"
        )
        if uploaded_file is not None:
            image_bytes = uploaded_file.read()
            db.save_proposal(current_row_idx, prop_notes_input, image_bytes, uploaded_file.name, cnpj=cnpj_val)
            st.toast("✅ Imagem da proposta enviada!", icon="🎉")
            st.rerun()

with col_prop_right:
    st.markdown("##### 🖼️ Visualização da Proposta (Imagem)")
    images_list = db.get_proposal_images(current_row_idx)
    
    if images_list:
        newest_img = images_list[0]
        st.success(f"📅 **Enviado em:** `{newest_img['date']}`")
        st.image(newest_img['path'], use_container_width=True)
        
        # Histórico de imagens anteriores (se houver mais de uma)
        if len(images_list) > 1:
            with st.expander(f"📜 Histórico ({len(images_list) - 1} imagens anteriores)"):
                for idx_img, img_item in enumerate(images_list[1:], start=1):
                    st.caption(f"**Enviado em: {img_item['date']}**")
                    st.image(img_item['path'], use_container_width=True)
                    st.divider()
    else:
        st.info("ℹ️ Nenhuma imagem de proposta enviada ainda. Tire um print (`Win + Shift + S`) e cole ao lado!")

st.divider()

# === BLOCO 3: REGISTRO DA CHAMADA E ATENDIMENTO ===
st.markdown("### 📝 Registro do Atendimento (Obrigatório)")

# Botões de Resposta Rápida
col_b1, col_b2, col_b3, col_b4 = st.columns(4)

if f"status_choice_{current_row_idx}" not in st.session_state:
    st.session_state[f"status_choice_{current_row_idx}"] = status_atual if status_atual != "Pendente" else "Atendeu"

if f"atendeu_choice_{current_row_idx}" not in st.session_state:
    st.session_state[f"atendeu_choice_{current_row_idx}"] = saved_atendeu if saved_atendeu != "Não Registrado" else "Sim"

with col_b1:
    if st.button(" Atendeu Chamada", use_container_width=True, type="primary" if st.session_state[f"atendeu_choice_{current_row_idx}"] == "Sim" else "secondary"):
        st.session_state[f"atendeu_choice_{current_row_idx}"] = "Sim"
        st.session_state[f"status_choice_{current_row_idx}"] = "Atendeu"
        st.rerun()

with col_b2:
    if st.button(" Não Atendeu", use_container_width=True, type="primary" if st.session_state[f"atendeu_choice_{current_row_idx}"] == "Não" else "secondary"):
        st.session_state[f"atendeu_choice_{current_row_idx}"] = "Não"
        st.session_state[f"status_choice_{current_row_idx}"] = "Não Atendeu"
        st.rerun()

with col_b3:
    if st.button("⏳ Agendar Retorno", use_container_width=True, type="primary" if st.session_state[f"status_choice_{current_row_idx}"] == "Agendado" else "secondary"):
        st.session_state[f"atendeu_choice_{current_row_idx}"] = "Sim"
        st.session_state[f"status_choice_{current_row_idx}"] = "Agendado"
        st.rerun()

with col_b4:
    if st.button("🚫 Sem Interesse / Inviável", use_container_width=True, type="primary" if st.session_state[f"status_choice_{current_row_idx}"] == "Sem Interesse" else "secondary"):
        st.session_state[f"atendeu_choice_{current_row_idx}"] = "Sim"
        st.session_state[f"status_choice_{current_row_idx}"] = "Sem Interesse"
        st.rerun()

st.caption(f"Seleção: Chamada **{st.session_state[f'atendeu_choice_{current_row_idx}']}** | Status: **{st.session_state[f'status_choice_{current_row_idx}']}**")

# Campo de texto obrigatório
obs_input = st.text_area(
    "Descreva o resultado da ligação / observações obrigatórias:",
    value=saved_obs,
    height=90,
    key=f"obs_input_{current_row_idx}",
    placeholder="Ex: Cliente interessado no plano 100GB. Solicitou proposta por WhatsApp para falar com a diretoria na quinta-feira."
)

obs_valida = bool(obs_input and obs_input.strip())

col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])

with col_nav1:
    if st.button("◀ Cliente Anterior", disabled=(current_row_idx == 0), use_container_width=True):
        curr_pos = lista_indices.index(current_row_idx) if current_row_idx in lista_indices else 0
        if curr_pos > 0:
            st.session_state.current_index = lista_indices[curr_pos - 1]
            st.rerun()

with col_nav2:
    if not obs_valida:
        st.error("⚠️ Preencha a observação para liberar o avanço.")
        st.button(" Gravar & Avançar para Próximo", disabled=True, use_container_width=True)
    else:
        if st.button(" Gravar & Avançar para Próximo Cliente", type="primary", use_container_width=True):
            db.save_interaction(
                row_id=current_row_idx,
                cnpj=cnpj_val,
                cliente=cliente_nome,
                status=st.session_state[f"status_choice_{current_row_idx}"],
                atendeu=st.session_state[f"atendeu_choice_{current_row_idx}"],
                observacao=obs_input.strip(),
                consultor=st.session_state.consultor
            )
            st.toast("✅ Atendimento registrado com sucesso!", icon="🎉")
            
            curr_pos = lista_indices.index(current_row_idx) if current_row_idx in lista_indices else 0
            if curr_pos < len(lista_indices) - 1:
                st.session_state.current_index = lista_indices[curr_pos + 1]
            st.rerun()

with col_nav3:
    if st.button("Próximo Sem Salvar ▶", disabled=(current_row_idx == lista_indices[-1]), use_container_width=True):
        curr_pos = lista_indices.index(current_row_idx) if current_row_idx in lista_indices else 0
        if curr_pos < len(lista_indices) - 1:
            st.session_state.current_index = lista_indices[curr_pos + 1]
            st.rerun()
