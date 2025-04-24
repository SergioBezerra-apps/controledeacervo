# Streamlit app for process control and reporting – v2 (inclui processos atípicos)

import streamlit as st

import pandas as pd

import matplotlib.pyplot as plt

import datetime as dt

import io, ssl, smtplib

from email.message import EmailMessage



# -----------------------------------------------------------------------------

# CONSTANTES E CONFIG BASICA

# -----------------------------------------------------------------------------

st.set_page_config(page_title="Controle de Acervo TCE-RJ", layout="wide")

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

# TÍTULO E UPLOADS

# -----------------------------------------------------------------------------



st.title("📑 Controle de Acervo & Relatórios – 3ª CAP / TCE-RJ")



upload_acervo = st.file_uploader("⬆️ Carregue a planilha *acervo portal bi.xlsx*", type=["xlsx"])

upload_manter = st.file_uploader("⬆️ Carregue *processosmanter.xlsx*", type=["xlsx"])



if not upload_acervo or not upload_manter:

    st.info("Envie os dois arquivos para prosseguir.")

    st.stop()



# -----------------------------------------------------------------------------

# CARREGAMENTO E PRÉ-TRATAMENTO

# -----------------------------------------------------------------------------



@st.cache_data(show_spinner=False)

def load_data(acervo_bytes, manter_bytes):

    acervo = pd.read_excel(acervo_bytes)

    manter = pd.read_excel(manter_bytes)



    manter_processos = set(manter[manter.columns[0]].astype(str).str.strip())

    acervo = acervo[acervo["Processo"].astype(str).str.strip().isin(manter_processos)].copy()



    # Tipos corretos

    acervo["Data Cadastro"] = pd.to_datetime(acervo["Data Cadastro"], errors="coerce")

    for col in ["Dias no Orgão", "Tempo TCERJ"]:

        acervo[col] = pd.to_numeric(acervo[col], errors="coerce")



    return acervo



df = load_data(upload_acervo, upload_manter)



st.success(f"Dados carregados: {len(df)} processos após cruzamento com lista 'manter'.")



# -----------------------------------------------------------------------------

# DETECÇÃO DE PROCESSOS ATÍPICOS – SEMPRE VISÍVEL

# -----------------------------------------------------------------------------



atyp_df = df[~df["Grupo Natureza"].isin(TYPICAL_GROUPS)].copy()



with st.expander("🚨 Processos ATÍPICOS detectados"):

    st.write(f"Total: **{len(atyp_df)}** processos fora da lista típica.")

    if not atyp_df.empty:

        st.dataframe(atyp_df[["Processo", "Grupo Natureza", "Data Cadastro", "Orgão Origem"]], use_container_width=True)

        # botões de download

        def to_excel_bytes(dframe):

            buf = io.BytesIO()

            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:

                dframe.to_excel(writer, index=False, sheet_name="Atipicos")

            return buf.getvalue()



        st.download_button("💾 Baixar lista de atípicos", to_excel_bytes(atyp_df), file_name=f"processos_atipicos_{TODAY.isoformat()}.xlsx")

    else:

        st.info("Nenhum processo atípico encontrado.")



# -----------------------------------------------------------------------------

# FILTROS DE PRIORIDADE

# -----------------------------------------------------------------------------



gn_options = sorted(df["Grupo Natureza"].unique())

selected_gn = st.selectbox("Grupo Natureza para priorização", options=gn_options, index=0)



sessao_filter = st.radio("Filtrar pela coluna *Já foi a Sessão*?", options=["Todos", "SIM", "NÃO"], horizontal=True)

num_procs = st.slider("Quantidade de processos a listar", 1, 20, 5)



base = df[df["Tipo Processo"].str.upper() == "PRINCIPAL"].copy()

if sessao_filter != "Todos":

    base = base[base["Já foi a Sessão"].str.upper() == sessao_filter]

base = base[base["Grupo Natureza"] == selected_gn]



special = selected_gn in [

    "CONTRATAÇÃO DE PESSOAL POR PRAZO DETERMINADO",

    "CONCURSO PÚBLICO",

]

if special and selected_gn == "CONCURSO PÚBLICO":

    base = base[base["Natureza"].str.contains("ADMISSÃO DE CONCURSADO", case=False, na=False)]



if special:

    base = base.sort_values("Data Cadastro").drop_duplicates(subset=["Orgão Origem"], keep="first")



result = base.sort_values("Data Cadastro").head(num_procs)



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

    st.dataframe(result.style.apply(alert_row, axis=1), use_container_width=True)



    # -------------------------------------------------------------------------

    # DOWNLOAD DO RELATÓRIO

    # -------------------------------------------------------------------------



    def to_excel_bytes(df_):

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            df_.to_excel(writer, index=False, sheet_name="Prioridade")

        return output.getvalue()



    excel_bytes = to_excel_bytes(result)

    default_fname = f"processos_prioritarios_{TODAY.isoformat()}.xlsx"



    st.download_button("💾 Baixar relatório em Excel", data=excel_bytes, file_name=default_fname)



    # -------------------------------------------------------------------------

    # ENVIO DE E-MAIL

    # -------------------------------------------------------------------------

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

                msg.add_attachment(excel_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=default_fname)

                context = ssl.create_default_context()

                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:

                    server.login(creds["user"], creds["pass"])

                    server.send_message(msg)

                st.success("E-mail enviado com sucesso! ✅")

            except Exception as e:

                st.error(f"Falha ao enviar e-mail: {e}")



# -----------------------------------------------------------------------------

# DASHBOARD RÁPIDO

# -----------------------------------------------------------------------------

with st.expander("📊 Dashboard de apoio"):

    col1, col2 = st.columns([2, 1])

    with col1:

        fig, ax = plt.subplots()

        (df.groupby("Grupo Natureza").size().sort_values().plot.barh(ax=ax))

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

