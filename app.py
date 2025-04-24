# Streamlit app for process control and reporting
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import datetime as dt
import io, ssl, smtplib
from email.message import EmailMessage

TODAY = dt.date.today()

st.set_page_config(page_title="Controle de Acervo 3ª CAP", layout="wide")

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

st.title("📑 Controle de Acervo & Relatórios – 3ª CAP / TCE‑RJ")

# ------------------------- FILE UPLOADS -------------------------
upload_acervo = st.file_uploader("⬆️ Carregue a planilha *acervo portal bi.xlsx*", type=["xlsx"])
upload_manter = st.file_uploader("⬆️ Carregue *processosmanter.xlsx*", type=["xlsx"])

if not upload_acervo or not upload_manter:
    st.info("Envie os dois arquivos para prosseguir.")
    st.stop()

@st.cache_data(show_spinner=False)
def load_data(acervo_bytes, manter_bytes):
    acervo = pd.read_excel(acervo_bytes)
    manter = pd.read_excel(manter_bytes)
    manter_processos = set(manter[manter.columns[0]].astype(str).str.strip())
    acervo = acervo[acervo["Processo"].astype(str).str.strip().isin(manter_processos)].copy()
    # Tipos
    acervo["Data Cadastro"] = pd.to_datetime(acervo["Data Cadastro"], errors="coerce")
    # Garante dias como int
    acervo["Dias no Orgão"] = pd.to_numeric(acervo["Dias no Orgão"], errors="coerce")
    acervo["Tempo TCERJ"] = pd.to_numeric(acervo["Tempo TCERJ"], errors="coerce")
    return acervo

df = load_data(upload_acervo, upload_manter)

st.success(f"Dados carregados: {len(df)} processos após cruzamento com lista 'manter'.")

# ------------------------- FILTROS -------------------------
gn_options = sorted(df["Grupo Natureza"].unique())
selected_gn = st.selectbox("Grupo Natureza", options=gn_options, index=0)

sessao_filter = st.radio("Filtrar pela coluna *Já foi a Sessão*?", options=["Todos", "SIM", "NÃO"], horizontal=True)

num_procs = st.slider("Quantidade de processos a listar", 1, 20, 5)

# Aplica filtros comuns
base = df[df["Tipo Processo"].str.upper() == "PRINCIPAL"].copy()
if sessao_filter != "Todos":
    base = base[base["Já foi a Sessão"].str.upper() == sessao_filter]

base = base[base["Grupo Natureza"] == selected_gn]

# regra especial contratação / concurso
special = selected_gn in [
    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO",
    "CONCURSO PÚBLICO",
]

if special and selected_gn == "CONCURSO PÚBLICO":
    base = base[base["Natureza"].str.contains("ADMISSÃO DE CONCURSADO", case=False, na=False)]

if special:
    # pega mais antigo por Órgão Origem
    base = (
        base.sort_values("Data Cadastro")
            .drop_duplicates(subset=["Orgão Origem"], keep="first")
    )

# Seleciona N mais antigos
result = base.sort_values("Data Cadastro").head(num_procs)

# ------------------------- ALERTAS -------------------------
def alert_row(row):
    # condição básica
    alert = (
        (row["Dias no Orgão"] > 180) |
        (row["Tempo TCERJ"] > 1825)
    )
    # regras extras para Contratação / Concurso
    if special:
        if (
            (row["Dias no Orgão"] >= 360) |
            (row["Dias no Orgão"] >= 720) |
            (row["Tempo TCERJ"] >= 360) |
            (row["Tempo TCERJ"] >= 720)
        ):
            alert = True

    # um estilo para cada coluna da linha
    return ["background-color:#ffcccc" if alert else "" for _ in row]

styled = result.style.apply(alert_row, axis=1)

st.subheader("📋 Resultado dos processos priorizados")
st.dataframe(styled, use_container_width=True)

# ------------------------- DASHBOARD -------------------------
with st.expander("📊 Dashboard de apoio"):
    col1, col2 = st.columns([2, 1])
    with col1:
        fig, ax = plt.subplots()
        (df.groupby("Grupo Natureza").size().sort_values().plot.barh(ax=ax))
        st.pyplot(fig)
    with col2:
        total = len(df)
        venc_180 = len(df[df["Dias no Orgão"] > 180])
        venc_360 = len(df[df["Dias no Orgão"] > 360])
        venc_720 = len(df[df["Dias no Orgão"] > 720])
        st.metric("Total de processos", total)
        st.metric(
            "Dias no Orgão > 180", venc_180,
            delta=(venc_180 / total * 100 if total else 0).__round__(1),
            delta_color="inverse",
        )
        st.metric("Dias no Orgão > 360", venc_360)
        st.metric("Dias no Orgão > 720", venc_720)

# ------------------------- DOWNLOAD -------------------------

def to_excel_bytes(df_):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df_.to_excel(writer, index=False, sheet_name="Prioridade")
    return out.getvalue()

if not result.empty:
    excel_bytes = to_excel_bytes(result)
    default_fname = f"processos_prioritarios_{TODAY.isoformat()}.xlsx"
    st.download_button("💾 Baixar relatório em Excel", data=excel_bytes, file_name=default_fname)

    # --------------------- EMAIL ---------------------
    st.subheader("✉️ Enviar por e‑mail")
    dest_default = "sergiollima2@hotmail.com"
    recip = st.text_input("Destinatários (separados por vírgula)", value=dest_default)
    if st.button("Enviar relatório"):
        if not recip:
            st.warning("Informe ao menos um e‑mail de destino.")
        else:
            try:
                creds = st.secrets["email"]  # {'user':..., 'pass':...}
                msg = EmailMessage()
                msg["Subject"] = f"Relatório Processos Prioritários – {TODAY.strftime('%d/%m/%Y')}"
                msg["From"] = creds["user"]
                msg["To"] = [r.strip() for r in recip.split(',')]
                msg.set_content("Segue em anexo o relatório gerado pelo app de controle de acervo.")
                msg.add_attachment(excel_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=default_fname)
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                    server.login(creds["user"], creds["pass"])
                    server.send_message(msg)
                st.success("E‑mail enviado com sucesso! ✅")
            except Exception as e:
                st.error(f"Falha ao enviar e‑mail: {e}")
else:
    st.info("Nenhum processo encontrado com os filtros atuais.")
