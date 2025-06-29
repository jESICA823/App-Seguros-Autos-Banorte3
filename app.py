import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage
import io

# Configura tu correo
EMAIL_EMISOR = "beristain.jessy.1e@gmail.com"
CLAVE_APP = "eyowisidnacjjzeo"
EMAIL_RECEPTOR = "beristain.jessy.1e@gmail.com"

# Título
st.title("📋 Control de Pólizas de Autos Financiados - Soluciones Operativas")

# Subida de archivo Excel
archivo = st.file_uploader("📎 Carga tu archivo Excel con las hojas 'AUTO' y 'Alianzas'", type=["xlsx"])

if archivo:
    try:
        auto_df = pd.read_excel(archivo, sheet_name="AUTO")
        alianzas_df = pd.read_excel(archivo, sheet_name="Alianzas ")
        st.success("✅ Archivo cargado correctamente")

        # Buscar por campo
        st.subheader("🔍 Buscar póliza")
        criterio = st.selectbox("Buscar por:", ["CREDITO", "NOMBRE_COMPLETO", "SERIE"])
        valor = st.text_input("Introduce el valor a buscar:")
        if valor:
            resultado = auto_df[auto_df[criterio].astype(str).str.contains(valor, case=False, na=False)]
            st.dataframe(resultado)

        # Duplicados
        duplicados = auto_df[auto_df["SERIE"].isin(alianzas_df["Número de Serie"])]
        st.subheader("⚠️ Pólizas Duplicadas")
        st.dataframe(duplicados)

        # Fechas
        hoy = pd.Timestamp(datetime.today().date())
        auto_df["FECHA"] = pd.to_datetime(auto_df["FECHA"], errors="coerce")
        auto_df["DIAS_RESTANTES"] = (auto_df["FECHA"] + pd.DateOffset(years=1) - hoy).dt.days

        proximas_a_vencer = auto_df[auto_df["DIAS_RESTANTES"] <= 30]
        vencidas = auto_df[auto_df["DIAS_RESTANTES"] < 0]
        sin_aseguradora = auto_df[auto_df["ASEGURADORA"].isna()]

        st.subheader("🔔 Pólizas por vencer")
        st.dataframe(proximas_a_vencer[["CREDITO", "NOMBRE_COMPLETO", "DIAS_RESTANTES"]])

        st.subheader("❌ Pólizas vencidas")
        st.dataframe(vencidas[["CREDITO", "NOMBRE_COMPLETO", "DIAS_RESTANTES"]])

        st.subheader("🚩 Pólizas sin aseguradora")
        st.dataframe(sin_aseguradora[["CREDITO", "NOMBRE_COMPLETO", "CLASIFICACION"]])

        # Exportar a Excel
        resumen = io.BytesIO()
        with pd.ExcelWriter(resumen, engine="openpyxl") as writer:
            auto_df.to_excel(writer, sheet_name="AUTO", index=False)
            alianzas_df.to_excel(writer, sheet_name="Alianzas", index=False)
            duplicados.to_excel(writer, sheet_name="Duplicados", index=False)
            proximas_a_vencer.to_excel(writer, sheet_name="Por Vencer", index=False)
            vencidas.to_excel(writer, sheet_name="Vencidas", index=False)
            sin_aseguradora.to_excel(writer, sheet_name="Sin Aseguradora", index=False)
        resumen.seek(0)

        # Enviar correo
        if st.button("📧 Enviar alerta por correo"):
            msg = EmailMessage()
            msg["Subject"] = "🚨 Alerta de pólizas"
            msg["From"] = EMAIL_EMISOR
            msg["To"] = EMAIL_RECEPTOR
            msg.set_content(
                f"""
Hola Jesica,

Este es el resumen actualizado de tu sistema de pólizas:

- Por vencer: {len(proximas_a_vencer)}
- Vencidas: {len(vencidas)}
- Sin aseguradora: {len(sin_aseguradora)}
- Duplicadas: {len(duplicados)}

Se adjunta el archivo Excel con el detalle.

"""
            )
            msg.add_attachment(resumen.read(), maintype="application", subtype="octet-stream", filename="Resumen_Polizas.xlsx")
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login(EMAIL_EMISOR, CLAVE_APP)
                    smtp.send_message(msg)
                st.success("✅ Correo enviado exitosamente a beristain.jessy.1e@gmail.com")
            except Exception as e:
                st.error(f"❌ Error al enviar correo: {e}")
    except Exception as e:
        st.error(f"⚠️ Ocurrió un error al leer el archivo: {e}")
