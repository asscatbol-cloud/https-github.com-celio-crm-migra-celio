import streamlit as st
import pandas as pd
import database as db
import io
import re
import os
import base64
import urllib.parse
from datetime import datetime, date, time
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
    .alert-callback {
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        padding: 10px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Função auxiliar resiliente para extração de campos do Excel/DB (Multi-Key e Normalizado)
def get_str(data_dict, keys, default="N/I"):
    if isinstance(keys, str):
        keys = [keys]
    
    # Dicionário normalizado (sem espaços nas chaves, em maiúsculas)
    norm_dict = {str(k).strip().upper().replace(" ", "_"): v for k, v in data_dict.items()}
    
    for key in keys:
        norm_target = str(key).strip().upper().replace(" ", "_")
        if norm_target in norm_dict:
            val = norm_dict[norm_target]
            if not pd.isna(val) and val is not None:
                val_str = str(val).strip()
                if val_str.lower() not in ["nan", "none", "<na>", "", "0.0", "nan.0"]:
                    return val_str
    return default

# Carregar dados do banco/excel
df_raw = db.load_dataset()

# Inicialização dos estados de sessão
if "current_index" not in st.session_state:
    st.session_state.current_index = 0

if "consultor" not in st.session_state:
    st.session_state.consultor = "Consultor"

# Função central para alterar cliente garantindo sincronia total dos widgets da sidebar
def set_current_client(index):
    st.session_state.current_index = index
    st.session_state.select_client_widget = index

# Callback do seletor da sidebar
def on_selectbox_change():
    if "select_client_widget" in st.session_state:
        st.session_state.current_index = st.session_state.select_client_widget

# Sidebar
st.sidebar.image("https://img.icons8.com/color/96/000000/headset.png", width=56)
st.sidebar.title(" Painel do Consultor")

consultor_input = st.sidebar.text_input("Seu Nome / ID:", value=st.session_state.consultor)
st.session_state.consultor = consultor_input

st.sidebar.divider()

# RECURSO: FILTRO POR VENDEDOR / CARTEIRA
st.sidebar.subheader("👤 Vendedor / Carteira")
lista_vendedores = set()
for v in df_raw.get("QUEM VENDEU", pd.Series(dtype=str)).dropna().tolist() + df_raw.get("CARTEIRA", pd.Series(dtype=str)).dropna().tolist():
    v_clean = str(v).strip()
    if v_clean and v_clean.lower() not in ["nan", "none", ""]:
        lista_vendedores.add(v_clean)

vendedor_selecionado = st.sidebar.selectbox("Filtrar por Carteira:", ["Todas"] + sorted(list(lista_vendedores)))

# Filtrar DataFrame pelo Vendedor se selecionado
if vendedor_selecionado != "Todas":
    col_vendeu = "QUEM VENDEU" if "QUEM VENDEU" in df_raw.columns else "QUEM_VENDEU"
    col_cart = "CARTEIRA" if "CARTEIRA" in df_raw.columns else "CARTEIRA"
    df = df_raw[(df_raw[col_vendeu] == vendedor_selecionado) | (df_raw[col_cart] == vendedor_selecionado)].copy()
    if df.empty:
        st.sidebar.warning("Nenhum cliente nesta carteira. Exibindo todos.")
        df = df_raw.copy()
else:
    df = df_raw.copy()

total_clientes = len(df)

# Progresso de Atendimentos
atendidos = (df["STATUS_CHAMADA"] != "Pendente").sum()
progresso = atendidos / total_clientes if total_clientes > 0 else 0

st.sidebar.subheader("📈 Progresso da Carteira")
st.sidebar.progress(progresso)
st.sidebar.write(f"**{atendidos}** de **{total_clientes}** clientes atendidos ({int(progresso*100)}%)")

st.sidebar.divider()

# RECURSO: ALERTAS DE RETORNO / AGENDAMENTOS (FOLLOW-UP)
st.sidebar.subheader("🔔 Agendamentos de Hoje")
df_agendados = df[df["STATUS_CHAMADA"] == "Agendado"]
today_str = datetime.now().strftime("%Y-%m-%d")

agendados_hoje = []
for idx, r in df_agendados.iterrows():
    cb_date = str(r.get("CALLBACK_DATE", "")).strip()
    cb_time = str(r.get("CALLBACK_TIME", "")).strip()
    if cb_date <= today_str:
        nome_c = get_str(r.to_dict(), ["CLIENTE", "RAZAO_SOCIAL"])
        agendados_hoje.append((idx, nome_c, cb_date, cb_time))

if agendados_hoje:
    st.sidebar.warning(f"⚠️ **{len(agendados_hoje)} retornos pendentes para hoje!**")
    for row_i, c_nome, c_d, c_t in agendados_hoje:
        if st.sidebar.button(f"📞 #{row_i+1} - {c_nome[:18]}... ({c_t})", key=f"btn_cb_{row_i}"):
            set_current_client(row_i)
            st.rerun()
else:
    st.sidebar.success(" Nenhuma ligação agendada pendente para hoje.")

st.sidebar.divider()

# Filtro e Navegação Direta
st.sidebar.subheader("🔍 Localizar Cliente")
modo_view = st.sidebar.radio("Filtrar lista por:", ["Todos", "Pendentes", "Atendidos", "Agendados"])

if modo_view == "Pendentes":
    lista_indices = df[df["STATUS_CHAMADA"] == "Pendente"].index.tolist()
elif modo_view == "Atendidos":
    lista_indices = df[df["STATUS_CHAMADA"] != "Pendente"].index.tolist()
elif modo_view == "Agendados":
    lista_indices = df[df["STATUS_CHAMADA"] == "Agendado"].index.tolist()
else:
    lista_indices = df.index.tolist()

if not lista_indices:
    st.sidebar.warning("Nenhum cliente no filtro selecionado.")
    lista_indices = df.index.tolist()

if st.session_state.current_index not in lista_indices:
    st.session_state.current_index = lista_indices[0]

# Garantir sincronia inicial da chave da sidebar
if "select_client_widget" not in st.session_state or st.session_state.select_client_widget not in lista_indices:
    st.session_state.select_client_widget = st.session_state.current_index

selected_pos = lista_indices.index(st.session_state.current_index) if st.session_state.current_index in lista_indices else 0

st.sidebar.selectbox(
    "Ir direto para cliente:",
    options=lista_indices,
    index=selected_pos,
    format_func=lambda i: f"#{i+1} - {get_str(df_raw.iloc[i].to_dict(), ['CLIENTE', 'RAZAO_SOCIAL'])[:25]}...",
    key="select_client_widget",
    on_change=on_selectbox_change
)

st.sidebar.divider()

# UPLOAD DE NOVAS PLANILHAS NA SIDEBAR
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
        set_current_client(0)
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

# === CORPO PRINCIPAL DO APP: SELEÇÃO DE MODO (ATENDIMENTO vs DASHBOARD) ===
modo_app = st.radio("Selecione o Modo:", ["📞 Atendimento Individual", "📊 Dashboard Estatístico da Carteira"], horizontal=True)

st.divider()

if modo_app == "📊 Dashboard Estatístico da Carteira":
    st.markdown("## 📊 Dashboard Gerencial da Carteira")
    
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    
    total_cart = len(df_raw)
    atendidos_cart = (df_raw["STATUS_CHAMADA"] != "Pendente").sum()
    atendeu_sucesso = (df_raw["STATUS_CHAMADA"] == "Atendeu").sum()
    prop_com_img = df_raw["PROPOSAL_IMAGE_PATH"].apply(lambda p: bool(p and str(p).strip())).sum()
    def check_viabilidade(r):
        t_rede = str(r.get("TIPO_REDE", "")).strip().upper()
        t_cob = str(r.get("TEM_COBERTURA_BANDA_LARGA", "")).strip().upper()
        t_man = str(r.get("TA NA MANCHA ", r.get("TA_NA_MANCHA_", ""))).strip().upper()
        return (t_rede == "GPON") or (t_cob == "SIM") or ("SIM" in t_man)

    gpon_count = df_raw.apply(check_viabilidade, axis=1).sum()
    
    with col_m1:
        st.metric("Total de Clientes", total_cart)
    with col_m2:
        st.metric("Atendidos (%)", f"{int((atendidos_cart/total_cart)*100)}%" if total_cart else "0%")
    with col_m3:
        st.metric("Chamadas Atendidas", atendeu_sucesso)
    with col_m4:
        st.metric("Propostas com Imagem", prop_com_img)
    with col_m5:
        st.metric("Cobertura GPON (%)", f"{int((gpon_count/total_cart)*100)}%" if total_cart else "0%")

    st.divider()
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.markdown("### 📞 Status das Chamadas")
        status_counts = df_raw["STATUS_CHAMADA"].value_counts()
        st.bar_chart(status_counts)
        
    with col_g2:
        st.markdown("### 👤 Atendimentos por Vendedor / Carteira")
        col_vend_name = "QUEM VENDEU" if "QUEM VENDEU" in df_raw.columns else "QUEM_VENDEU"
        if col_vend_name in df_raw.columns:
            vendedor_counts = df_raw[df_raw["STATUS_CHAMADA"] != "Pendente"][col_vend_name].value_counts()
            if not vendedor_counts.empty:
                st.bar_chart(vendedor_counts)
            else:
                st.info("Nenhum atendimento realizado ainda.")

else:
    # === MODO ATENDIMENTO INDIVIDUAL ===
    current_row_idx = st.session_state.current_index
    row_dict = df_raw.iloc[current_row_idx].to_dict()
    real_row_id = int(row_dict.get("ROW_ID", current_row_idx))

    cliente_nome = get_str(row_dict, ["CLIENTE", "RAZAO_SOCIAL", "NOME_CLIENTE"])
    cnpj_val = get_str(row_dict, ["CNPJ_CLIENTE", "CNPJ", "CGC"])
    contato_nome = get_str(row_dict, ["CONTATO", "NOME_CONTATO", "CONTATO_CLIENTE"])
    tel_raw = get_str(row_dict, ["NR_TELEFONE", "TELEFONE", "CELULAR", "FONE"], "")
    municipio_val = get_str(row_dict, ["DS_MUNICIPIO", "MUNICIPIO", "CIDADE"])
    uf_val = get_str(row_dict, ["UF", "ESTADO"])
    endereco_val = get_str(row_dict, ["NR_ENDERECO", "ENDERECO", "LOGRADOURO"])
    numero_val = get_str(row_dict, ["Nº", "N", "NUMERO"], "")
    cep_val = get_str(row_dict, ["NR_CEP", "CEP"], "")
    plano_val = get_str(row_dict, ["PLANO", "PLANO_ATUAL"])
    linhas_val = get_str(row_dict, ["QTDE DE LINHAS", "QTDE_DE_LINHAS", "LINHAS", "QT_PLANTA"], "1")
    tipo_rede_val = get_str(row_dict, ["TIPO_REDE", "REDE"], "SEM GPON").upper()
    cobertura_val = get_str(row_dict, ["TEM_COBERTURA_BANDA_LARGA", "COBERTURA", "MANCHA"], "Não")
    marca_val = get_str(row_dict, ["APARELHO_TRAFEGO_MARCA", "MARCA_APARELHO", "MARCA"], "")
    modelo_val = get_str(row_dict, ["APARELHO_TRAFEGO_MODELO", "MODELO_APARELHO", "MODELO"], "")
    
    # Recomendação com fallback seguro para a coluna APAREHOS
    rec_linha = get_str(row_dict, ["RECOMENDACAO_APARELHO_LINHA", "RECOMENDACAO"])
    if rec_linha != "N/I":
        recomendacao_val = rec_linha
    else:
        rec_ap = get_str(row_dict, "APAREHOS")
        recomendacao_val = rec_ap if rec_ap not in ["N/I", "0"] else "Sem recomendação específica"

    status_atual = get_str(row_dict, "STATUS_CHAMADA", "Pendente")
    saved_atendeu = get_str(row_dict, "ATENDEU", "Não Registrado")
    saved_obs = get_str(row_dict, "INTERACAO_CONSULTOR", "")
    proposal_notes_saved = get_str(row_dict, "PROPOSAL_NOTES", "")
    saved_cb_date = get_str(row_dict, "CALLBACK_DATE", "")
    saved_cb_time = get_str(row_dict, "CALLBACK_TIME", "")

    # Header do Cliente
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"<div class='main-header'>📞 Cliente #{current_row_idx + 1} de {total_clientes} - {cliente_nome}</div>", unsafe_allow_html=True)
    with col_h2:
        if status_atual == "Pendente":
            st.warning(f"Status: **{status_atual}**")
        elif status_atual == "Agendado":
            st.info(f"Status: **Agendado ({saved_cb_date})**")
        else:
            st.success(f"Status: **{status_atual}**")

    st.divider()

    # === BLOCO 1: DADOS DO CLIENTE (EMPRESA, CONTATO, GPON, APARELHOS) ===
    col_info1, col_info2 = st.columns([1, 1])

    with col_info1:
        st.markdown("####  Empresa & Contato")
        st.markdown(f"**CNPJ:** `{cnpj_val}`")
        st.markdown(f"**Contato:** **{contato_nome}**")
        
        # MENSAGEM AUTOMÁTICA DE WHATSAPP (BOM DIA/BOA TARDE)
        tel_clean = re.sub(r'\D', '', tel_raw.replace('.0', ''))
        st.markdown(f"**Telefone:** `{tel_raw if tel_raw else 'N/I'}`")
        
        # Calcular Saudação com base no horário atual
        hora_atual = datetime.now().hour
        saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
        
        # MENSAGEM EXATA EXIGIDA PELO USUÁRIO
        msg_txt = f"{saudacao} meu nome é {st.session_state.consultor} sou consultor vivo empresas, responsavel pelas linhas moveis da sua empresa, temos uma revisão das ofertas e tenho algumas opções que gostaria de discutir com você podemos conversar?"
        msg_encoded = urllib.parse.quote(msg_txt)
        
        if tel_clean:
            tel_wa = tel_clean if tel_clean.startswith("55") else "55" + tel_clean
            wa_link_url = f"https://web.whatsapp.com/send?phone={tel_wa}&text={msg_encoded}"
            
            st.markdown(f"""
                <a href="{wa_link_url}" target="whatsapp_web_window" style="
                    background-color: #25D366; 
                    color: white; 
                    padding: 8px 16px; 
                    text-decoration: none; 
                    border-radius: 6px; 
                    font-weight: bold; 
                    display: inline-block;
                    margin-top: 5px;
                ">📱 Iniciar Conversa no WhatsApp (Janela Única)</a>
            """, unsafe_allow_html=True)
            
        # Caixa com a mensagem pronta para copiar com 1 clique
        st.text_area("Mensagem Pronta para o WhatsApp:", value=msg_txt, height=85, key=f"wa_msg_box_{current_row_idx}")

        st.markdown(f"**Endereço:** {endereco_val}, Nº {numero_val} - {municipio_val}/{uf_val}")
        st.markdown(f"**Plano Atual:** `{plano_val}` ({linhas_val} linhas)")

    with col_info2:
        st.markdown("#### ⚡ Viabilidade de Rede & Aparelhos")
        
        mancha_val = get_str(row_dict, ["TA_NA_MANCHA_", "TA NA MANCHA"], "0")
        tem_gpon = (tipo_rede_val == "GPON") or (cobertura_val.upper() == "SIM") or ("SIM" in mancha_val.upper())
        if tem_gpon:
            label_gpon = tipo_rede_val if tipo_rede_val not in ["N/I", "SEM GPON"] else ("GPON" if cobertura_val.upper() == "SIM" else "MANCHA FIBRA")
            st.markdown(f"**Rede:** <span class='badge-gpon'>{label_gpon} (Com Viabilidade)</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"**Rede:** <span class='badge-no-gpon'>SEM GPON (Sem Cobertura)</span>", unsafe_allow_html=True)

        aparelho_trafego = f"{marca_val} {modelo_val}".strip()
        st.markdown(f"**Aparelho em Tráfego:** {aparelho_trafego if aparelho_trafego else 'Não informado'}")
        st.markdown(f"**Recomendação de Aparelho:** <span class='badge-device'>{recomendacao_val}</span>", unsafe_allow_html=True)

    st.divider()

    # === BLOCO 2: PAINEL DA PROPOSTA (NOTAS DA ESQUERDA vs IMAGEM NA DIREITA) ===
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
            db.save_proposal(real_row_id, prop_notes_input, cnpj=cnpj_val)
            st.toast("✅ Notas da proposta salvas!", icon="💾")
            st.rerun()

        st.markdown("---")
        st.markdown("##### 📸 Anexar / Colar Print da Proposta")
        
        st.caption("1. Clique na caixa abaixo e aperte **CTRL + V** (PrintScreen / Win+Shift+S):")
        pasted_b64 = image_paste_component(key=f"paste_box_{current_row_idx}")
        
        if pasted_b64 and isinstance(pasted_b64, str) and pasted_b64.startswith("data:image"):
            try:
                header, b64_data = pasted_b64.split(",", 1)
                img_bytes = base64.b64decode(b64_data)
                db.save_proposal(real_row_id, prop_notes_input, img_bytes, "print_colado.png", cnpj=cnpj_val)
                st.toast("✅ Print colado com sucesso!", icon="📸")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao colar imagem: {e}")

        with st.expander("Ou selecione um arquivo do computador"):
            uploaded_file = st.file_uploader(
                "Escolha o arquivo de imagem:",
                type=["png", "jpg", "jpeg", "webp"],
                key=f"prop_img_uploader_{current_row_idx}"
            )
            if uploaded_file is not None:
                image_bytes = uploaded_file.read()
                db.save_proposal(real_row_id, prop_notes_input, image_bytes, uploaded_file.name, cnpj=cnpj_val)
                st.toast("✅ Imagem da proposta enviada!", icon="🎉")
                st.rerun()

    with col_prop_right:
        st.markdown("##### 🖼️ Visualização da Proposta (Imagem)")
        images_list = db.get_proposal_images(real_row_id)
        
        if images_list:
            newest_img = images_list[0]
            
            col_img_date, col_img_del = st.columns([3, 1])
            with col_img_date:
                st.success(f"📅 **Enviado em:** `{newest_img['date']}`")
            with col_img_del:
                if st.button("🗑️ Excluir Imagem", key=f"del_newest_{current_row_idx}"):
                    db.delete_proposal_image(real_row_id, newest_img['path'])
                    st.toast("🗑️ Imagem excluída com sucesso!", icon="🚮")
                    st.rerun()
                    
            st.image(newest_img['path'], use_container_width=True)
            
            if len(images_list) > 1:
                with st.expander(f"📜 Histórico ({len(images_list) - 1} imagens anteriores)"):
                    for idx_img, img_item in enumerate(images_list[1:], start=1):
                        col_hist_dt, col_hist_del = st.columns([3, 1])
                        with col_hist_dt:
                            st.caption(f"**Enviado em: {img_item['date']}**")
                        with col_hist_del:
                            if st.button("🗑️ Excluir", key=f"del_hist_{current_row_idx}_{idx_img}"):
                                db.delete_proposal_image(real_row_id, img_item['path'])
                                st.toast("🗑️ Imagem antiga excluída com sucesso!", icon="🚮")
                                st.rerun()
                                
                        st.image(img_item['path'], use_container_width=True)
                        st.divider()
        else:
            st.info("ℹ️ Nenhuma imagem de proposta enviada ainda. Tire um print (`Win + Shift + S`) e cole ao lado!")

    st.divider()

    # === BLOCO 3: REGISTRO DA CHAMADA E ATENDIMENTO ===
    st.markdown("### 📝 Registro do Atendimento (Obrigatório)")

    col_b1, col_b2, col_b3, col_b4 = st.columns(4)

    if f"status_choice_{current_row_idx}" not in st.session_state:
        st.session_state[f"status_choice_{current_row_idx}"] = status_atual if status_atual != "Pendente" else "Atendeu"

    if f"atendeu_choice_{current_row_idx}" not in st.session_state:
        st.session_state[f"atendeu_choice_{current_row_idx}"] = saved_atendeu if saved_atendeu != "Não Registrado" else "Sim"

    with col_b1:
        if st.button(" Atendeu Chamada", use_container_width=True, type="primary" if st.session_state[f"atendeu_choice_{current_row_idx}"] == "Sim" and st.session_state[f"status_choice_{current_row_idx}"] == "Atendeu" else "secondary"):
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

    # DATA E HORA DE AGENDAMENTO DE RETORNO (SE FOR AGENDADO)
    cb_date_val = ""
    cb_time_val = ""
    if st.session_state[f"status_choice_{current_row_idx}"] == "Agendado":
        st.markdown("<div class='alert-callback'>📅 <b>Definir Data e Horário para Retorno da Ligação:</b></div>", unsafe_allow_html=True)
        col_cb1, col_cb2 = st.columns(2)
        with col_cb1:
            default_d = datetime.strptime(saved_cb_date, "%Y-%m-%d").date() if saved_cb_date else date.today()
            sel_date = st.date_input("Data do Retorno:", value=default_d, key=f"date_cb_{current_row_idx}")
            cb_date_val = sel_date.strftime("%Y-%m-%d")
        with col_cb2:
            default_t = datetime.strptime(saved_cb_time, "%H:%M").time() if saved_cb_time else time(10, 0)
            sel_time = st.time_input("Horário do Retorno:", value=default_t, key=f"time_cb_{current_row_idx}")
            cb_time_val = sel_time.strftime("%H:%M")

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
                set_current_client(lista_indices[curr_pos - 1])
                st.rerun()

    with col_nav2:
        if not obs_valida:
            st.error("⚠️ Preencha a observação para liberar o avanço.")
            st.button(" Gravar & Avançar para Próximo", disabled=True, use_container_width=True)
        else:
            if st.button(" Gravar & Avançar para Próximo Cliente", type="primary", use_container_width=True):
                db.save_interaction(
                    row_id=real_row_id,
                    cnpj=cnpj_val,
                    cliente=cliente_nome,
                    status=st.session_state[f"status_choice_{current_row_idx}"],
                    atendeu=st.session_state[f"atendeu_choice_{current_row_idx}"],
                    observacao=obs_input.strip(),
                    consultor=st.session_state.consultor,
                    callback_date=cb_date_val,
                    callback_time=cb_time_val
                )
                st.toast("✅ Atendimento registrado com sucesso!", icon="🎉")
                
                curr_pos = lista_indices.index(current_row_idx) if current_row_idx in lista_indices else 0
                if curr_pos < len(lista_indices) - 1:
                    set_current_client(lista_indices[curr_pos + 1])
                st.rerun()

    with col_nav3:
        if st.button("Próximo Sem Salvar ▶", disabled=(current_row_idx == lista_indices[-1]), use_container_width=True):
            curr_pos = lista_indices.index(current_row_idx) if current_row_idx in lista_indices else 0
            if curr_pos < len(lista_indices) - 1:
                set_current_client(lista_indices[curr_pos + 1])
                st.rerun()
