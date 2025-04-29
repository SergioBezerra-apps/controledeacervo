# Streamlit app for process control and reporting – v3 (melhorias solicitadas)  
# Autor: ChatGPT (OpenAI) – ajustado conforme requisicoes de Sérgio Luiz (3ª CAP)
# -----------------------------------------------------------------------------
# IMPORTS
# -----------------------------------------------------------------------------
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt
import io, ssl, smtplib, re
from email.message import EmailMessage

# -----------------------------------------------------------------------------
# CONFIG BASICA
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Controle de Acervo TCE‑RJ", layout="wide")
TODAY = dt.date.today()

TYPICAL_GROUPS = [
    "APOSENTADORIA",
    "CONCURSO PÚBLICO",
    "CONCURSO PÚBLICO (DOC)",
    "CONCURSO PÚBLICO (RETIFICAÇÃO)",
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO",
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO (RETIFICAÇÃO)",
    "PENSÃO",
    "PROMOÇÃO",
    "REFORMA",
    "RESPOSTA A OFÍCIO",
    "REVISÃO DE PENSÃO",
    "REVISÃO DE PROVENTOS",
    "TRANSFERÊNCIA PARA RESERVA REMUNERADA",
]

# -----------------------------------------------------------------------------
# UPLOADS
# -----------------------------------------------------------------------------
st.title("📑 Controle de Acervo & Relatórios – 3ª CAP / TCE‑RJ")

upload_acervo = st.file_uploader("⬆️ Carregue a planilha *acervo portal bi.xlsx*", type=["xlsx"])
upload_manter = st.file_uploader("⬆️ Carregue *processosmanter.xlsx*", type=["xlsx"])

if not upload_acervo or not upload_manter:
    st.info("Envie os dois arquivos para prosseguir.")
    st.stop()

# -----------------------------------------------------------------------------
# LOAD & PRE‑TREAT
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data(acervo_bytes, manter_bytes):
    acervo = pd.read_excel(acervo_bytes)
    manter = pd.read_excel(manter_bytes)

    manter_processos = set(manter.iloc[:, 0].astype(str).str.strip())
    acervo_cruzado = acervo[acervo["Processo"].astype(str).str.strip().isin(manter_processos)].copy()

    # Types
    acervo_cruzado["Data Cadastro"] = pd.to_datetime(acervo_cruzado["Data Cadastro"], errors="coerce")
    for col in ["Dias no Orgão", "Tempo TCERJ"]:
        acervo_cruzado[col] = pd.to_numeric(acervo_cruzado[col], errors="coerce")

    return acervo, acervo_cruzado  # retorna original e cruzado

acervo_raw, df = load_data(upload_acervo, upload_manter)

st.success(f"Dados carregados: {len(df)} processos após cruzamento com lista 'manter'.")

# -----------------------------------------------------------------------------
# PROCESSOS ATÍPICOS
# -----------------------------------------------------------------------------
atyp_df = df[~df["Grupo Natureza"].isin(TYPICAL_GROUPS)].copy()
with st.expander("🚨 Processos ATÍPICOS detectados"):
    st.write(f"Total: **{len(atyp_df)}** processos fora da lista típica.")
    if not atyp_df.empty:
        st.dataframe(atyp_df[["Processo", "Grupo Natureza", "Data Cadastro", "Orgão Origem"]], use_container_width=True)
        def to_excel_bytes(dframe):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                dframe.to_excel(writer, index=False, sheet_name="Atipicos")
            return buf.getvalue()
        st.download_button("💾 Baixar lista de atípicos", to_excel_bytes(atyp_df), file_name=f"processos_atipicos_{TODAY.isoformat()}.xlsx")
    else:
        st.info("Nenhum processo atípico encontrado.")

# -----------------------------------------------------------------------------
# FILTROS & PRIORIZAÇÃO
# -----------------------------------------------------------------------------
# (1) seleção de grupos – pode escolher "Todos"
all_option = "— TODOS —"
options_gn = [all_option] + sorted(df["Grupo Natureza"].unique())
selected_gn = st.selectbox("Grupo Natureza para priorização", options=options_gn, index=0)

# (2) filtros numeric sliders
with st.expander("⚙️ Filtros avançados"):
    colF1, colF2 = st.columns(2)
    with colF1:
        min_dias = st.number_input("Dias no Órgão – mínimo", value=0, step=10)
        max_dias = st.number_input("Dias no Órgão – máximo (0 = sem limite)", value=0, step=10)
    with colF2:
        min_tce = st.number_input("Tempo TCERJ – mínimo (dias)", value=0, step=30)
        max_tce = st.number_input("Tempo TCERJ – máximo (0 = sem limite)", value=0, step=30)

sessao_filter = st.radio("Filtrar pela coluna *Já foi a Sessão*?", options=["Todos", "SIM", "NÃO"], horizontal=True)
num_procs = st.slider("Quantidade de processos a listar", 1, 50, 10)

base = df[df["Tipo Processo"].str.upper() == "PRINCIPAL"].copy()
if selected_gn != all_option:
    base = base[base["Grupo Natureza"] == selected_gn]

# filtros numéricos
if min_dias:
    base = base[base["Dias no Orgão"] >= min_dias]
if max_dias:
    base = base[base["Dias no Orgão"] <= max_dias]
if min_tce:
    base = base[base["Tempo TCERJ"] >= min_tce]
if max_tce:
    base = base[base["Tempo TCERJ"] <= max_tce]

if sessao_filter != "Todos":
    base = base[base["Já foi a Sessão"].str.upper() == sessao_filter]

# regras especiais mantidas
special = selected_gn in [
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO",
    "CONCURSO PÚBLICO",
]
if special and selected_gn == "CONCURSO PÚBLICO":
    base = base[base["Natureza"].str.contains("ADMISSÃO DE CONCURSADO", case=False, na=False)]
if special:
    base = base.sort_values("Data Cadastro").drop_duplicates(subset=["Orgão Origem"], keep="first")

# (3) ordenação: 1º Tempo TCERJ > 1800 desc, depois Dias > 150 desc, senão Data Cadastro
base["pri_flag_tce"] = base["Tempo TCERJ"] > 1800
base["pri_flag_dias"] = base["Dias no Orgão"] > 150
base = base.sort_values([
    "pri_flag_tce",  # True primeiro
    "Tempo TCERJ",   # desc
    "pri_flag_dias", # True primeiro
    "Dias no Orgão", # desc
    "Data Cadastro"  # mais antigo primeiro
], ascending=[False, False, False, False, True])

result = base.head(num_procs).copy()

# -----------------------------------------------------------------------------
# HIGHLIGHT DE ALERTAS
# -----------------------------------------------------------------------------

def alert_row(row):
    alert = (row["Dias no Orgão"] > 180) or (row["Tempo TCERJ"] > 1825)
    if special:
        if (
            (row["Dias no Orgão"] >= 360) or (row["Dias no Orgão"] >= 720) or
            (row["Tempo TCERJ"] >= 360) or (row["Tempo TCERJ"] >= 720)
        ):
            alert = True
    return ["background-color:#ffcccc" if alert else "" for _ in row]

st.subheader("📋 Processos priorizados")
if result.empty:
    st.info("Nenhum processo encontrado com os filtros atuais.")
else:
    st.dataframe(result.drop(columns=["pri_flag_tce", "pri_flag_dias"]).style.apply(alert_row, axis=1), use_container_width=True)

    # download
    def to_excel_bytes(df_):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            df_.to_excel(w, index=False, sheet_name="Prioridade")
        return out.getvalue()
    excel_bytes = to_excel_bytes(result)
    st.download_button("💾 Baixar relatório em Excel", data=excel_bytes, file_name=f"processos_prioritarios_{TODAY.isoformat()}.xlsx")

    # envio email
    st.subheader("✉️ Enviar por e-mail")
    recip_default = "sergiollima2@hotmail.com"
    recip = st.text_input("Destinatários (separados por vírgula)", value=recip_default)
    if st.button("Enviar relatório"):
        if not recip:
            st.warning("Informe ao menos um e-mail válido.")
        else:
            try:
                creds = st.secrets["email"]  # {'user':..., 'pass':...}
                msg = EmailMessage()
                msg["Subject"] = f"Relatório Processos Prioritários – {TODAY.strftime('%d/%m/%Y')}"
                msg["From"] = creds["user"]
                msg["To"] = [r.strip() for r in recip.split(',')]
                msg.set_content("Segue em anexo o relatório gerado pelo app de controle de acervo.")
                msg.add_attachment(excel_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"processos_prioritarios_{TODAY.isoformat()}.xlsx")
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as server:
                    server.login(creds["user"], creds["pass"])
                    server.send_message(msg)
                st.success("E-mail enviado com sucesso! ✅")
            except Exception as e:
                st.error(f"Falha ao enviar e-mail: {e}")

# -----------------------------------------------------------------------------
# (4) ANALISE – DOCUMENTOS NA COLUNA OBS
# -----------------------------------------------------------------------------
with st.expander("🔍 Identificação de DOCS não juntados" ):
    # filtra somente DOCUMENTO
    docs_df = acervo_raw[acervo_raw["Tipo Processo"].str.upper() == "DOCUMENTO"].copy()
    docs_df["processo_observado"] = docs_df["Observação"].astype(str).str.extract(r"(\d{6}-\d+/\d{4})", expand=False)

    # separa encontrados e não encontrados
    manter_set = set(df["Processo"].astype(str).str.strip())
    docs_df["encontrado_na_3cap"] = docs_df["processo_observado"].isin(manter_set)

    docs_com = docs_df[docs_df["encontrado_na_3cap"] == True]
    docs_sem = docs_df[docs_df["encontrado_na_3cap"] == False]

    colC1, colC2 = st.columns(2)
    with colC1:
        st.subheader("DOCS não juntados COM proc. principal na 3ª CAP")
        st.write(f"Total: **{len(docs_com)}**")
        st.dataframe(docs_com[["Processo", "Observação", "processo_observado"]], use_container_width=True)
    with colC2:
        st.subheader("DOCS não juntados SEM proc. principal na 3ª CAP")
        st.write(f"Total: **{len(docs_sem)}**")
        st.dataframe(docs_sem[["Processo", "Observação", "processo_observado"]], use_container_width=True)

    # opção download
    def _tozip():
        mem = io.BytesIO()
        with pd.ExcelWriter(mem, engine="xlsxwriter") as writer:
            docs_com.to_excel(writer, index=False, sheet_name="COM_principal")
            docs_sem.to_excel(writer, index=False, sheet_name="SEM_principal")
        return mem.getvalue()
    st.download_button("💾 Baixar resultado DOCS (Excel)", _tozip(), file_name=f"docs_nao_juntados_{TODAY.isoformat()}.xlsx")

# -----------------------------------------------------------------------------
# DASHBOARD RÁPIDO
# -----------------------------------------------------------------------------
with st.expander("📊 Dashboard de apoio"):
    col1, col2 = st.columns([2, 1])
    with col1:
        fig, ax = plt.subplots()
        df.groupby("Grupo Natureza").size().sort_values().plot.barh(ax=ax)
        ax.set_xlabel("Quantidade")
        st.pyplot(fig)
    with col2:
        total = len(df)
        venc_180 = len(df[df["Dias no Orgão"] > 180])
        venc_360 = len(df[df["Dias no Orgão"] > 360])
        venc_720 = len(df[df["Dias no Orgão"] > 720])
        st.metric("Total processos", total)
        st.metric("> 180 dias", venc_180)
        st.metric("> 360 dias", venc_360)
        st.metric("> 720 dias", venc_720)

# -----------------------------------------------------------------------------
# RODAPÉ
# -----------------------------------------------------------------------------
st.caption(f"Desenvolvido para 3ª CAP / TCE‑RJ · {TODAY.year}")
