# -*- coding: utf-8 -*-
"""
Simulador APNR, Solvencia (pbs) y Pricing / Rentabilidad (RAROC, EVA, Breakeven) — V2.

Marco normativo:
- Decreto 1477 de 2018 (compilado en Decreto 2555 de 2010, Art. 2.1.1.3.2 a 2.1.1.3.5)
- Decreto 573 de 2025 (Ponderación cartera de libranza consumo)
- Capítulo 1 de la Parte 2 de la Circular Básica Financiera (CBF) de la SFC

Alcance: Simulador de originación y pricing (excluye cartera en incumplimiento).
"""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from apnr_motor_v2 import (
    DEFAULT_PARAM_FILENAME,
    agregar_cliente_en_memoria,
    calcular_operacion_v2,
    cargar_parametros_desde_bytes,
    cargar_parametros_desde_ruta,
    evaluar_lote_operaciones,
    exportar_parametros_a_bytes,
)

# ==============================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
# ==============================================================================
st.set_page_config(
    page_title="Simulador APNR, Solvencia & Pricing (V2) | SFC",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #002B49 0%, #004B87 100%);
        padding: 20px 25px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 43, 73, 0.18);
    }
    .main-header h1 { color: white !important; margin: 0; font-size: 24px; font-weight: 700; }
    .main-header p { color: #CBD5E1 !important; margin: 4px 0 0 0; font-size: 13.5px; }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 5px solid #004B87;
        padding: 12px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.06);
    }
    .section-title {
        color: #002B49;
        border-bottom: 2px solid #004B87;
        padding-bottom: 4px;
        margin-top: 16px;
        margin-bottom: 12px;
        font-weight: 600;
        font-size: 15.5px;
    }
    .formula-box {
        background-color: #EBF4FF;
        border-left: 4px solid #004B87;
        padding: 12px;
        border-radius: 6px;
        font-family: monospace;
        font-size: 13px;
        color: #0A192F;
        margin-bottom: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_PARAM_PATH = BASE_DIR / DEFAULT_PARAM_FILENAME

# ==============================================================================
# 2. GESTIÓN DE PARÁMETROS EN SESIÓN (REFRESCO MENSUAL SIN MUTAR INSUMOS)
# ==============================================================================
if "params_v2" not in st.session_state:
    params_iniciales, validacion_inicial = cargar_parametros_desde_ruta(DEFAULT_PARAM_PATH)
    st.session_state["params_v2"] = params_iniciales
    st.session_state["validacion_v2"] = validacion_inicial
    st.session_state["origen_params"] = f"Archivo base ({DEFAULT_PARAM_FILENAME})"

params = st.session_state["params_v2"]
validacion = st.session_state["validacion_v2"]
meta_str = params["meta_str"]
meta_num = params["meta_num"]

df_gar = params["Garantias"]
df_ltv = params["LTV_Ponderacion"]
df_tipos_cp = params["Tipos_Contraparte"]
df_matriz = params["Matriz_Contraparte"]
df_matriz_cp = params["Matriz_Calif_CP"]
df_ccf = params["Factores_Conversion_CCF"]
df_finesp = params["Fin_Especializada"]
df_cli = params["Clientes"]
df_pricing_ref = params["Parametros_Riesgo_Pricing"]

# ==============================================================================
# 3. CASOS PRECARGADOS (PLANTILLA EXCEL + CASOS NORMATIVOS V2)
# ==============================================================================
CASOS_ESTUDIO = {
    "1. Plantilla SFC - Garantía Banco Ext USD (Colateral Financiero 8% Fc)": {
        "id": "010023000030", "cartera": "COMERCIAL", "unidad": "COP Millones (MM)",
        "saldo": 40000.0, "prov": 0.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "BE", "val_gar1": 40000.0, "mon_ob": "COP", "mon_gar1": "USD",
        "usar_gar2": False, "gar2": "AN", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.132,
    },
    "2. Compañía Mundial de Seguros (Piso 20% Art. 2.1.1.3.4)": {
        "id": "860037013", "cartera": "COMERCIAL", "unidad": "COP Millones (MM)",
        "saldo": 80000.0, "prov": 0.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "AN", "val_gar1": 79000.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": False, "gar2": "AN", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.128,
    },
    "3. Rendifin S.A.S. (Pignoración Acciones + Provisiones)": {
        "id": "900111222", "cartera": "COMERCIAL", "unidad": "COP Pesos ($)",
        "saldo": 56316000000.0, "prov": 504283108.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "PA", "val_gar1": 156550127707.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": False, "gar2": "AN", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.145,
    },
    "4. Inverinmobiliarias SAS (Pyme Hipoteca - Mitiga vía LTV sin reducir E)": {
        "id": "900555666", "cartera": "COMERCIAL", "unidad": "COP Pesos ($)",
        "saldo": 11241050182.0, "prov": 68159911.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "HP", "val_gar1": 36000000000.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": False, "gar2": "AN", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.140,
    },
    "5. Garantía Mixta Parágrafo 2 (Hipoteca + FNG Concurrente)": {
        "id": "900888999", "cartera": "COMERCIAL", "unidad": "COP Millones (MM)",
        "saldo": 10000.0, "prov": 0.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "HP", "val_gar1": 15000.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": True, "gar2": "FG", "val_gar2": 5000.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.142,
    },
    "6. Consumo Libranza (Decreto 573 de 2025 - RW 60%)": {
        "id": "1020304050", "cartera": "CONSUMO", "unidad": "COP Millones (MM)",
        "saldo": 80.0, "prov": 0.0, "cupo_no_util": 0.0, "ccf": 1.00,
        "gar1": "LB", "val_gar1": 0.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": False, "gar2": "SF", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": True, "corto_plazo": False, "tasa": 0.165,
    },
    "7. Cupo Rotativo Corporativo con Tramo No Desembolsado (CCF 10%)": {
        "id": "900333444", "cartera": "COMERCIAL", "unidad": "COP Millones (MM)",
        "saldo": 20000.0, "prov": 0.0, "cupo_no_util": 30000.0, "ccf": 0.10,
        "gar1": "FG", "val_gar1": 10000.0, "mon_ob": "COP", "mon_gar1": "COP",
        "usar_gar2": False, "gar2": "SF", "val_gar2": 0.0, "mon_gar2": "COP",
        "libranza": False, "corto_plazo": False, "tasa": 0.138,
    },
}

# ==============================================================================
# 4. ENCABEZADO Y BARRA LATERAL (GESTIÓN DE REFRESCO MENSUAL)
# ==============================================================================
st.markdown(
    f"""
    <div class="main-header">
        <h1>🏦 Simulador APNR, Solvencia y Pricing (RAROC / EVA) — V2</h1>
        <p>Motor Paramétrico Normativo | {meta_str.get('Version_Normativa', 'Decreto 1477/2018 & CBF SFC')} | Corte Activo: <b>{meta_str.get('Fecha_Corte', 'N/D')}</b></p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### ⚡ Casos de Estudio Normativos")
    caso_sel = st.selectbox("Cargar caso preconfigurado:", list(CASOS_ESTUDIO.keys()))
    preset = CASOS_ESTUDIO[caso_sel]

    st.divider()
    st.markdown("### 🔄 Refresco Mensual de Parámetros")
    st.caption(f"**Origen activo:** {st.session_state['origen_params']}")
    st.caption(f"**Fecha de corte:** `{meta_str.get('Fecha_Corte', 'N/D')}`")

    bytes_descarga = exportar_parametros_a_bytes(params)
    fecha_tag = datetime.now().strftime("%Y%m%d")
    st.download_button(
        label="📥 Descargar Excel Paramétrico V2",
        data=bytes_descarga,
        file_name=f"parametros_apnr_sfc_v2_{fecha_tag}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    archivo_subido = st.file_uploader("📤 Cargar Excel mensual actualizado (.xlsx)", type=["xlsx"])
    if archivo_subido is not None:
        nuevos_params, val_subida = cargar_parametros_desde_bytes(archivo_subido.getvalue())
        if val_subida.valido and nuevos_params is not None:
            st.session_state["params_v2"] = nuevos_params
            st.session_state["validacion_v2"] = val_subida
            st.session_state["origen_params"] = f"Cargado en sesión ({archivo_subido.name})"
            st.success("Parámetros mensuales validados y activados en sesión.")
            st.rerun()
        else:
            st.error("El archivo subido falló la validación:")
            for err in val_subida.errores:
                st.write(f"- {err}")

    if st.button("↩️ Restaurar parámetros base del disco", width="stretch"):
        p_base, v_base = cargar_parametros_desde_ruta(DEFAULT_PARAM_PATH)
        st.session_state["params_v2"] = p_base
        st.session_state["validacion_v2"] = v_base
        st.session_state["origen_params"] = f"Archivo base ({DEFAULT_PARAM_FILENAME})"
        st.rerun()

    st.divider()
    st.markdown("### ⚙️ Parámetros de Solvencia y Pricing")
    st.caption(f"• **Solvencia Actual:** `{meta_num['Solvencia_Total_Actual']*100:.2f}%`")
    st.caption(f"• **Solvencia Objetivo Pricing:** `{meta_num['Solvencia_Objetivo_Pricing']*100:.2f}%`")
    st.caption(f"• **Ke (Hurdle Rate):** `{meta_num['Costo_Capital_Ke_EA']*100:.2f}% E.A.`")
    st.caption(f"• **FTP + Liquidez Base:** `{(meta_num['FTP_Base_EA']+meta_num['Prima_Liquidez_EA'])*100:.2f}% E.A.`")
    st.caption(f"• **Impuesto Renta:** `{meta_num['Tasa_Impuesto_Renta']*100:.1f}%`")
    st.caption(f"• **Tope 0.2% Minorista:** `{meta_num['Cartera_Minorista_Total_MM']*meta_num['Tope_Concentracion_Minorista_Pct']:,.0f} MM`")

# ==============================================================================
# 5. PESTAÑAS PRINCIPALES
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "🧮 Simulador APNR & Pricing (RAROC/EVA)",
        "📦 Simulador Masivo de Portafolio (Lote)",
        "📊 Explorador y Validador Paramétrico",
        "📖 Auditoría Normativa (Decreto 1477/2018 & V1→V2)",
    ]
)

with tab1:
    col_in, col_out = st.columns([1.08, 1.0])

    with col_in:
        st.markdown('<div class="section-title">1. Cliente y Clasificación de Contraparte</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1.1, 1])
        with c1:
            id_input = st.text_input("Número Identificación Cliente (NIT / ID)", value=preset["id"])

        id_clean = id_input.strip()
        id_no_zeros = id_clean.lstrip("0") or "0"
        match_cli = df_cli[
            (df_cli["Nro_ID"].astype(str).str.strip() == id_clean)
            | (df_cli["Nro_ID_9"].astype(str).str.strip() == id_no_zeros)
        ]

        if not match_cli.empty:
            fila_c = match_cli.iloc[0]
            nombre_cli = fila_c["Nombre_Cliente"]
            sigla_cp = fila_c["tipo_contraparte"]
            calif_lp = fila_c["Calificacion_LP"]
            calif_cp = fila_c["Calificacion_CP"]
            row_cp_cat = df_tipos_cp[df_tipos_cp["SIGLA"] == sigla_cp]
            desc_cp = row_cp_cat.iloc[0]["DESCRIPCION"] if not row_cp_cat.empty else sigla_cp
            with c2:
                st.success(f"✔️ Cliente localizado (`{sigla_cp}`)")
            st.info(
                f"**Cliente:** {nombre_cli} | **Contraparte:** `{sigla_cp}` ({desc_cp}) | "
                f"**Calif. LP:** `{calif_lp}` | **Calif. CP:** `{calif_cp}`"
            )
        else:
            with c2:
                st.warning("⚠️ ID no encontrado. Ingreso manual:")
            m1, m2, m3, m4 = st.columns([1.2, 1.3, 1, 0.8])
            with m1:
                nombre_cli = st.text_input("Nombre Entidad / Cliente", value="NUEVO CLIENTE SAS")
            with m2:
                opciones_cp = [f"{r['SIGLA']} - {r['DESCRIPCION']}" for _, r in df_tipos_cp.iterrows()]
                cp_sel = st.selectbox("Tipo Contraparte", opciones_cp, index=19)
                sigla_cp = cp_sel.split(" - ")[0]
                desc_cp = cp_sel.split(" - ")[1]
            with m3:
                califs_pos = df_matriz[df_matriz["TIPO_CONTRAPARTE"] == sigla_cp]["CALIFICACION"].unique().tolist()
                if not califs_pos:
                    califs_pos = ["SIN CALIFICACION", "NA"]
                calif_lp = st.selectbox("Calificación LP", califs_pos)
            with m4:
                califs_cp_opts = ["NA"] + df_matriz_cp["CALIFICACION_CP"].tolist()
                calif_cp = st.selectbox("Calif. CP", califs_cp_opts)

            if st.button("💾 Agregar cliente a la sesión (exportable en Excel)", width="stretch"):
                st.session_state["params_v2"] = agregar_cliente_en_memoria(
                    params, id_clean, nombre_cli, sigla_cp, "CO", calif_lp, calif_cp
                )
                st.success("Cliente agregado en memoria. Descarga el Excel desde el panel izquierdo para conservarlo.")
                st.rerun()

        st.markdown('<div class="section-title">2. Obligación, Cartera y Contingencias (Art. 2.1.1.3.5)</div>', unsafe_allow_html=True)
        u1, u2, u3 = st.columns(3)
        with u1:
            clase_cartera = st.selectbox(
                "Clase Cartera",
                ["COMERCIAL", "CONSUMO", "VIVIENDA"],
                index=["COMERCIAL", "CONSUMO", "VIVIENDA"].index(preset["cartera"]),
            )
        with u2:
            unidad_mon = st.selectbox(
                "Unidad Monetaria",
                ["COP Millones (MM)", "COP Pesos ($)"],
                index=0 if preset["unidad"] == "COP Millones (MM)" else 1,
            )
        with u3:
            moneda_ob = st.selectbox(
                "Moneda Obligación",
                ["COP", "USD", "EUR", "UVR"],
                index=["COP", "USD", "EUR", "UVR"].index(preset["mon_ob"]),
            )

        s1, s2 = st.columns(2)
        with s1:
            saldo_desembolsado = st.number_input(
                f"Saldo Desembolsado ({unidad_mon})",
                min_value=0.0,
                value=float(preset["saldo"]),
                step=100.0,
                format="%.2f",
            )
        with s2:
            provisiones = st.number_input(
                f"Provisiones Individuales ({unidad_mon})",
                min_value=0.0,
                max_value=float(saldo_desembolsado),
                value=min(float(preset["prov"]), float(saldo_desembolsado)),
                step=10.0,
                format="%.2f",
            )

        c_ccf1, c_ccf2 = st.columns([1, 1.3])
        with c_ccf1:
            cupo_no_utilizado = st.number_input(
                f"Cupo No Utilizado / Contingencia ({unidad_mon})",
                min_value=0.0,
                value=float(preset["cupo_no_util"]),
                step=100.0,
                format="%.2f",
            )
        with c_ccf2:
            opciones_ccf = [f"{r['Tipo_Exposicion']} ({r['CCF']*100:.0f}%)" for _, r in df_ccf.iterrows()]
            idx_ccf = 3 if preset["cupo_no_util"] > 0 and preset["ccf"] == 0.10 else 0
            ccf_sel = st.selectbox("Factor Conversión Crediticia (CCF - Art. 2.1.1.3.5)", opciones_ccf, index=idx_ccf)
            ccf_val = float(df_ccf.iloc[opciones_ccf.index(ccf_sel)]["CCF"])

        j1, j2, j3, j4 = st.columns(4)
        with j1:
            es_constructor = st.checkbox("🏗️ Crédito Constructor", value=False)
        with j2:
            es_libranza = st.checkbox(
                "💳 Libranza (Dec. 573/25)",
                value=preset["libranza"],
                disabled=(clase_cartera != "CONSUMO"),
            )
        with j3:
            es_corto_plazo = st.checkbox("⏱️ Corto Plazo (usa CP)", value=preset["corto_plazo"])
        with j4:
            inmueble_ok = st.checkbox("🏠 Avalúo / Inm. Construido", value=True)

        finesp_sel = st.selectbox(
            "🌉 Financiación Especializada / Subordinada (Art. 2.1.1.3.2 & 2.1.1.3.3)",
            df_finesp["Subtipo"].tolist(),
            index=0,
            disabled=es_constructor,
        )

        st.markdown('<div class="section-title">3. Garantías y Mitigación (Art. 2.1.1.3.4)</div>', unsafe_allow_html=True)
        lista_gar_str = [
            f"{r['Tipo de Garantía']} - {r['Descripción de garantía']} [{r['Clase Garantía']} | {r['TIPO_GARANTIA_SFC']}]"
            for _, r in df_gar.iterrows()
        ]
        idx_gar1 = next((i for i, r in df_gar.iterrows() if r["Tipo de Garantía"] == preset["gar1"]), 0)
        gar1_sel = st.selectbox("Garantía Principal", lista_gar_str, index=idx_gar1)
        cod_gar1 = gar1_sel.split(" - ")[0]

        g1_c1, g1_c2 = st.columns(2)
        with g1_c1:
            val_gar1 = st.number_input(
                f"Valor Garantía 1 ({unidad_mon})",
                min_value=0.0,
                value=float(preset["val_gar1"]),
                step=100.0,
                format="%.2f",
            )
        with g1_c2:
            mon_gar1 = st.selectbox(
                "Moneda Garantía 1",
                ["COP", "USD", "EUR", "UVR"],
                index=["COP", "USD", "EUR", "UVR"].index(preset["mon_gar1"]),
            )

        usar_gar2 = st.checkbox("➕ Agregar 2da Garantía Concurrente (incluye Garantías Mixtas Parágrafo 2)", value=preset["usar_gar2"])
        if usar_gar2:
            idx_gar2 = next((i for i, r in df_gar.iterrows() if r["Tipo de Garantía"] == preset["gar2"]), 0)
            gar2_sel = st.selectbox("Garantía 2", lista_gar_str, index=idx_gar2)
            cod_gar2 = gar2_sel.split(" - ")[0]
            g2_c1, g2_c2 = st.columns(2)
            with g2_c1:
                val_gar2 = st.number_input(
                    f"Valor Garantía 2 ({unidad_mon})",
                    min_value=0.0,
                    value=float(preset["val_gar2"]),
                    step=100.0,
                    format="%.2f",
                )
            with g2_c2:
                mon_gar2 = st.selectbox(
                    "Moneda Garantía 2",
                    ["COP", "USD", "EUR", "UVR"],
                    index=["COP", "USD", "EUR", "UVR"].index(preset["mon_gar2"]),
                )
        else:
            cod_gar2, val_gar2, mon_gar2 = "SF", 0.0, "COP"

        with st.expander("⚙️ Opciones Avanzadas (Modo Normativo Estricto vs Conciliación Plantilla)", expanded=False):
            aplicar_piso_20 = st.checkbox(
                "Aplicar piso normativo E >= 20% de A (Art. 2.1.1.3.4 — Obligatorio por norma)",
                value=True,
            )
            hipoteca_mitiga_exp = st.checkbox(
                "Modo Comparativo Plantilla: permitir que HIPOTECA reduzca E (Norma Parágrafo 1 indica NO marcar)",
                value=False,
            )
            fa_activo = st.number_input("Factor Ajuste Activo (Fa)", min_value=0.0, max_value=1.0, value=float(meta_num["Factor_Ajuste_Activo_Fa_Default"]), step=0.01)

        st.markdown('<div class="section-title">4. Módulo de Pricing y Rentabilidad (P&L de la Operación)</div>', unsafe_allow_html=True)
        row_ref_p = df_pricing_ref[df_pricing_ref["Clase_Cartera"] == clase_cartera]
        pd_def = float(row_ref_p.iloc[0]["PD_Referencia"]) if not row_ref_p.empty else 0.015
        lgd_def = float(row_ref_p.iloc[0]["LGD_Referencia"]) if not row_ref_p.empty else 0.45
        gasto_def = float(row_ref_p.iloc[0]["Costo_Operativo_Ref_Pct"]) if not row_ref_p.empty else meta_num["Gasto_Operativo_Pct"]

        p1, p2, p3 = st.columns(3)
        with p1:
            tasa_activa_ea = st.number_input("Tasa Activa Cliente (% E.A.)", min_value=0.0, max_value=100.0, value=float(preset["tasa"] * 100), step=0.25) / 100.0
            comision_des_pct = st.number_input("Comisión Desembolso (%)", min_value=0.0, max_value=20.0, value=0.0, step=0.10) / 100.0
        with p2:
            ftp_ea = st.number_input("Costo Fondeo FTP (% E.A.)", min_value=0.0, max_value=50.0, value=float(meta_num["FTP_Base_EA"] * 100), step=0.25) / 100.0
            prima_liq_ea = st.number_input("Prima Liquidez (% E.A.)", min_value=0.0, max_value=20.0, value=float(meta_num["Prima_Liquidez_EA"] * 100), step=0.10) / 100.0
        with p3:
            pd_anual = st.number_input("PD Anual Esperada (%)", min_value=0.0, max_value=100.0, value=pd_def * 100, step=0.10) / 100.0
            lgd_pct = st.number_input("LGD / Severidad (%)", min_value=0.0, max_value=100.0, value=lgd_def * 100, step=1.0) / 100.0

        p4, p5, p6, p7 = st.columns(4)
        with p4:
            comision_disp_pct = st.number_input("Comisión Cupo No Util. (%)", min_value=0.0, max_value=10.0, value=0.50 if cupo_no_utilizado > 0 else 0.0, step=0.10) / 100.0
        with p5:
            gasto_op_pct = st.number_input("Gasto Operativo (% Saldo)", min_value=0.0, max_value=20.0, value=gasto_def * 100, step=0.10) / 100.0
        with p6:
            solvencia_obj_pct = st.number_input("Solvencia Obj. (%)", min_value=4.5, max_value=30.0, value=float(meta_num["Solvencia_Objetivo_Pricing"] * 100), step=0.25) / 100.0
        with p7:
            ke_ea = st.number_input("Hurdle Rate Ke (% E.A.)", min_value=1.0, max_value=50.0, value=float(meta_num["Costo_Capital_Ke_EA"] * 100), step=0.50) / 100.0

    # ==========================================================================
    # MOTOR DE CÁLCULO (COLUMNA DERECHA)
    # ==========================================================================
    res = calcular_operacion_v2(
        params,
        sigla_cp=sigla_cp,
        calif_lp=calif_lp,
        calif_cp=calif_cp,
        clase_cartera=clase_cartera,
        unidad_mon=unidad_mon,
        moneda_ob=moneda_ob,
        saldo_desembolsado=saldo_desembolsado,
        provisiones=provisiones,
        cupo_no_utilizado=cupo_no_utilizado,
        ccf=ccf_val,
        fa_activo=fa_activo,
        es_corto_plazo=es_corto_plazo,
        es_constructor=es_constructor,
        finesp_subtipo=finesp_sel,
        es_libranza=es_libranza,
        gar1_cod=cod_gar1,
        gar1_val=val_gar1,
        gar1_mon=mon_gar1,
        usar_gar2=usar_gar2,
        gar2_cod=cod_gar2,
        gar2_val=val_gar2,
        gar2_mon=mon_gar2,
        inmueble_cumple_requisitos=inmueble_ok,
        aplicar_piso_20=aplicar_piso_20,
        hipoteca_mitiga_exp=hipoteca_mitiga_exp,
        tasa_activa_ea=tasa_activa_ea,
        comision_desembolso_pct=comision_des_pct,
        comision_disponibilidad_pct=comision_disp_pct,
        ftp_ea=ftp_ea,
        prima_liquidez_ea=prima_liq_ea,
        pd_anual=pd_anual,
        lgd_pct=lgd_pct,
        gasto_operativo_pct=gasto_op_pct,
        tasa_impuesto=meta_num["Tasa_Impuesto_Renta"],
        solvencia_objetivo=solvencia_obj_pct,
        ke_ea=ke_ea,
    )

    with col_out:
        st.markdown('<div class="section-title">📊 Resultados Regulatorios, Solvencia e Indicadores de Rentabilidad</div>', unsafe_allow_html=True)

        if res["excede_tope_minorista"]:
            st.warning("⚠️ La exposición excede el 0.2% de la cartera minorista total: según Art. 2.1.1.3.2 pondera al 100% como Corporativo Sin Calificación.")

        st.markdown(
            f"""
            <div class="formula-box">
                <b>Exposición Art. 2.1.1.3.4 & 2.1.1.3.5:</b> A = {res['activo_neto_balance']:,.2f} (balance) + {res['exposicion_contingente_eq']:,.2f} (contingente × CCF) = <b>{res['activo_neto_A']:,.2f}</b><br>
                <b>Mitigación en E:</b> E = ({res['activo_neto_A']:,.2f} × (1 + {fa_activo:.2f})) - ({res['g_computable']:,.2f} × (1 - {res['fg_efectivo']:.4f} - {res['fc_efectivo']:.4f})) = <b>{res['valor_exposicion']:,.2f}</b><br>
                <b>APNR:</b> {res['valor_exposicion']:,.2f} × {res['ponderacion']*100:.1f}% = <b>{res['apnr']:,.2f}</b> | <b>Ahorro total APNR por gtía:</b> {res['ahorro_apnr']:,.2f}
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Fila 1 KPIs Regulatorios
        k1, k2, k3 = st.columns(3)
        k1.metric("Valor Exposición (E)", f"${res['valor_exposicion']:,.1f}", f"-${res['efecto_mitigacion_e']:,.1f} en E")
        k2.metric("Ponderación (RW)", f"{res['ponderacion']*100:.1f}%", f"LTV: {res['ltv_texto']}")
        k3.metric("APNR Calculado", f"${res['apnr']:,.1f}", f"Sin gtía: ${res['apnr_sin_gtia']:,.1f}")

        # Fila 2 KPIs Solvencia y Capital
        k4, k5, k6 = st.columns(3)
        k4.metric("Impacto Neto Solvencia", f"{res['impacto_neto_pbs']:+.3f} pbs", f"APNR: {res['impacto_apnr_pbs']:.3f} | PT: +{res['impacto_pt_pbs']:.3f}")
        k5.metric("Capital Económico Req.", f"${res['capital_req']:,.1f}", f"CET1+Colchón: ${res['capital_cet1_req']:,.1f}")
        k6.metric("Solvencia Post-Op.", f"{res['solvencia_post']*100:.4f}%", f"Base: {meta_num['Solvencia_Total_Actual']*100:.2f}%")

        # Fila 3 KPIs Rentabilidad y Pricing
        k7, k8, k9 = st.columns(3)
        delta_raroc = (res["raroc"] - ke_ea) * 100
        k7.metric("RAROC Operación", f"{res['raroc']*100:.2f}%", f"{delta_raroc:+.2f}% vs Ke ({ke_ea*100:.1f}%)")
        k8.metric("EVA (Valor Económico)", f"${res['eva']:,.2f}", f"ROA: {res['roa']*100:.2f}% | RORWA: {res['rorwa']*100:.2f}%")
        k9.metric("Tasa Mínima Breakeven", f"{res['tasa_breakeven_ea']*100:.2f}% E.A.", f"Spread mín FTP: {res['spread_min_sobre_ftp_pbs']:,.0f} pbs")

        st.success(f"⚖️ **Regla Normativa Aplicada:** {res['regla_pond']}")

        # Gráficos interactivos de Rentabilidad: Cascada P&L y Sensibilidad Tasa vs RAROC
        g_col1, g_col2 = st.columns(2)
        with g_col1:
            fig_wf = go.Figure(
                go.Waterfall(
                    orientation="v",
                    measure=["relative", "relative", "relative", "total", "relative", "relative", "relative", "total"],
                    x=["Intereses", "Comisiones", "Fondeo+Liq", "Margen Bruto", "Pérdida Esp.", "Gastos Op.", "Impuestos", "Utilidad Neta"],
                    y=[
                        res["ingreso_intereses"],
                        res["ingreso_comisiones"],
                        -res["costo_fondeo"],
                        0,
                        -res["perdida_esperada_el"],
                        -res["gastos_operativos"],
                        -res["impuestos"],
                        0,
                    ],
                    connector={"line": {"color": "#94A3B8"}},
                    increasing={"marker": {"color": "#0E7490"}},
                    decreasing={"marker": {"color": "#B91C1C"}},
                    totals={"marker": {"color": "#002B49"}},
                )
            )
            fig_wf.update_layout(title="Cascada P&L Anual de la Operación", height=320, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_wf, width="stretch")

        with g_col2:
            tasas_grid = np.linspace(max(0.02, res["tasa_breakeven_ea"] - 0.05), max(tasa_activa_ea + 0.05, res["tasa_breakeven_ea"] + 0.06), 25)
            rarocs_grid = []
            for t_i in tasas_grid:
                r_i = calcular_operacion_v2(
                    params,
                    sigla_cp=sigla_cp, calif_lp=calif_lp, calif_cp=calif_cp,
                    clase_cartera=clase_cartera, unidad_mon=unidad_mon, moneda_ob=moneda_ob,
                    saldo_desembolsado=saldo_desembolsado, provisiones=provisiones,
                    cupo_no_utilizado=cupo_no_utilizado, ccf=ccf_val, fa_activo=fa_activo,
                    es_corto_plazo=es_corto_plazo, es_constructor=es_constructor,
                    finesp_subtipo=finesp_sel, es_libranza=es_libranza,
                    gar1_cod=cod_gar1, gar1_val=val_gar1, gar1_mon=mon_gar1,
                    usar_gar2=usar_gar2, gar2_cod=cod_gar2, gar2_val=val_gar2, gar2_mon=mon_gar2,
                    inmueble_cumple_requisitos=inmueble_ok, aplicar_piso_20=aplicar_piso_20,
                    hipoteca_mitiga_exp=hipoteca_mitiga_exp, tasa_activa_ea=float(t_i),
                    comision_desembolso_pct=comision_des_pct, comision_disponibilidad_pct=comision_disp_pct,
                    ftp_ea=ftp_ea, prima_liquidez_ea=prima_liq_ea, pd_anual=pd_anual, lgd_pct=lgd_pct,
                    gasto_operativo_pct=gasto_op_pct, tasa_impuesto=meta_num["Tasa_Impuesto_Renta"],
                    solvencia_objetivo=solvencia_obj_pct, ke_ea=ke_ea,
                )
                rarocs_grid.append(r_i["raroc"] * 100.0)

            fig_sens = go.Figure()
            fig_sens.add_trace(go.Scatter(x=tasas_grid * 100, y=rarocs_grid, mode="lines+markers", name="RAROC (%)", line=dict(color="#004B87", width=3)))
            fig_sens.add_hline(y=ke_ea * 100, line_dash="dash", line_color="#DC2626", annotation_text=f"Ke ({ke_ea*100:.1f}%)")
            fig_sens.add_vline(x=res["tasa_breakeven_ea"] * 100, line_dash="dot", line_color="#059669", annotation_text=f"Breakeven ({res['tasa_breakeven_ea']*100:.2f}%)")
            fig_sens.update_layout(
                title="Sensibilidad Tasa Activa vs. RAROC",
                xaxis_title="Tasa Activa (% E.A.)",
                yaxis_title="RAROC (%)",
                height=320,
                margin=dict(l=10, r=10, t=35, b=10),
            )
            st.plotly_chart(fig_sens, width="stretch")

        # Tabla Ficha Completa y Descarga Multi-hoja
        df_ficha = pd.DataFrame(
            {
                "Concepto": [
                    "Identificación Cliente", "Nombre Cliente", "Tipo Contraparte", "Clase Cartera",
                    f"Saldo Desembolsado ({unidad_mon})", f"Provisiones Individuales ({unidad_mon})",
                    f"Cupo No Utilizado ({unidad_mon})", "Factor Conversión Crediticia (CCF)",
                    f"Activo Neto Equivalente A ({unidad_mon})", f"Valor Garantía Bruta ({unidad_mon})",
                    f"Garantía Computable en E ({unidad_mon})", "Fc Efectivo", "Fg Efectivo",
                    f"Valor Exposición E ({unidad_mon})", "LTV", "Ponderación Aplicada (RW)",
                    f"APNR Calculado ({unidad_mon})", f"APNR Sin Garantía ({unidad_mon})",
                    "Impacto APNR (pbs)", "Impacto Utilidad en PT (pbs)", "Impacto Neto Solvencia (pbs)",
                    "Solvencia Estimada Post-Operación", f"Capital Económico Req. ({unidad_mon})",
                    f"Ingresos Totales ({unidad_mon})", f"Costo Fondeo + Liquidez ({unidad_mon})",
                    f"Margen Financiero Bruto ({unidad_mon})", f"Pérdida Esperada EL ({unidad_mon})",
                    f"Gastos Operativos ({unidad_mon})", f"Impuestos ({unidad_mon})",
                    f"Utilidad Neta ({unidad_mon})", "ROA (%)", "RORWA (%)", "RAROC (%)",
                    f"EVA ({unidad_mon})", "Tasa Mínima Breakeven (% E.A.)", "Spread Mínimo sobre FTP (pbs)",
                ],
                "Valor": [
                    id_clean, nombre_cli, f"{sigla_cp} - {res['desc_cp']}", clase_cartera,
                    f"{res['saldo_desembolsado']:,.2f}", f"{res['provisiones']:,.2f}",
                    f"{res['cupo_no_utilizado']:,.2f}", f"{res['ccf']*100:.0f}%",
                    f"{res['activo_neto_A']:,.2f}", f"{res['val_gar_total_bruto']:,.2f}",
                    f"{res['g_computable']:,.2f}", f"{res['fc_efectivo']*100:.2f}%", f"{res['fg_efectivo']*100:.2f}%",
                    f"{res['valor_exposicion']:,.2f}", res["ltv_texto"], f"{res['ponderacion']*100:.2f}%",
                    f"{res['apnr']:,.2f}", f"{res['apnr_sin_gtia']:,.2f}",
                    f"{res['impacto_apnr_pbs']:.4f} pbs", f"{res['impacto_pt_pbs']:.4f} pbs", f"{res['impacto_neto_pbs']:.4f} pbs",
                    f"{res['solvencia_post']*100:.4f}%", f"{res['capital_req']:,.2f}",
                    f"{res['ingresos_totales']:,.2f}", f"{res['costo_fondeo']:,.2f}",
                    f"{res['margen_fin_bruto']:,.2f}", f"{res['perdida_esperada_el']:,.2f}",
                    f"{res['gastos_operativos']:,.2f}", f"{res['impuestos']:,.2f}",
                    f"{res['utilidad_neta']:,.2f}", f"{res['roa']*100:.2f}%", f"{res['rorwa']*100:.2f}%", f"{res['raroc']*100:.2f}%",
                    f"{res['eva']:,.2f}", f"{res['tasa_breakeven_ea']*100:.2f}%", f"{res['spread_min_sobre_ftp_pbs']:,.1f} pbs",
                ],
            }
        )
        st.dataframe(df_ficha, width="stretch", hide_index=True, height=340)

        buf_exp = BytesIO()
        with pd.ExcelWriter(buf_exp, engine="openpyxl") as writer:
            df_ficha.to_excel(writer, sheet_name="Liquidacion_APNR_Pricing", index=False)
        st.download_button(
            label="📥 Exportar Liquidación APNR & Pricing a Excel",
            data=buf_exp.getvalue(),
            file_name=f"Liquidacion_APNR_Pricing_V2_{id_clean}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )

# ==============================================================================
# 6. PESTAÑA 2: SIMULADOR MASIVO DE PORTAFOLIO (LOTE)
# ==============================================================================
with tab2:
    st.markdown("### 📦 Simulador Masivo de Portafolio / Lote de Operaciones")
    st.write("Descarga la plantilla de lote, diligencia tus operaciones simuladas y evalúa APNR, RAROC, EVA e impacto agregado en solvencia.")

    df_tpl_lote = pd.DataFrame(
        [
            {
                "ID_Operacion": "SIM-001", "ID_Cliente": "900111222", "Tipo_Contraparte": "GEL", "Calificacion_LP": "SIN CALIFICACION",
                "Clase_Cartera": "COMERCIAL", "Unidad": "COP Millones (MM)", "Moneda_Obligacion": "COP",
                "Saldo_Desembolsado": 25000.0, "Provisiones": 0.0, "Cupo_No_Utilizado": 0.0, "CCF": 1.0,
                "Es_Constructor": False, "Es_Libranza": False,
                "Codigo_Garantia_1": "PA", "Valor_Garantia_1": 30000.0, "Moneda_Garantia_1": "COP",
                "Tasa_Activa_EA": 0.145, "PD_Anual": 0.015, "LGD": 0.45,
            },
            {
                "ID_Operacion": "SIM-002", "ID_Cliente": "900555666", "Tipo_Contraparte": "MEL", "Calificacion_LP": "NA",
                "Clase_Cartera": "COMERCIAL", "Unidad": "COP Millones (MM)", "Moneda_Obligacion": "COP",
                "Saldo_Desembolsado": 11241.0, "Provisiones": 68.0, "Cupo_No_Utilizado": 0.0, "CCF": 1.0,
                "Es_Constructor": False, "Es_Libranza": False,
                "Codigo_Garantia_1": "HP", "Valor_Garantia_1": 36000.0, "Moneda_Garantia_1": "COP",
                "Tasa_Activa_EA": 0.140, "PD_Anual": 0.015, "LGD": 0.40,
            },
            {
                "ID_Operacion": "SIM-003", "ID_Cliente": "1020304050", "Tipo_Contraparte": "PNL", "Calificacion_LP": "NA",
                "Clase_Cartera": "CONSUMO", "Unidad": "COP Millones (MM)", "Moneda_Obligacion": "COP",
                "Saldo_Desembolsado": 120.0, "Provisiones": 0.0, "Cupo_No_Utilizado": 0.0, "CCF": 1.0,
                "Es_Constructor": False, "Es_Libranza": True,
                "Codigo_Garantia_1": "LB", "Valor_Garantia_1": 0.0, "Moneda_Garantia_1": "COP",
                "Tasa_Activa_EA": 0.168, "PD_Anual": 0.025, "LGD": 0.45,
            },
        ]
    )
    buf_tpl = BytesIO()
    with pd.ExcelWriter(buf_tpl, engine="openpyxl") as writer:
        df_tpl_lote.to_excel(writer, sheet_name="Operaciones", index=False)

    st.download_button(
        label="📥 Descargar Plantilla de Lote (.xlsx)",
        data=buf_tpl.getvalue(),
        file_name="plantilla_simulacion_lote_v2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    lote_subido = st.file_uploader("📤 Subir archivo Excel de lote para simular", type=["xlsx"], key="lote_uploader")
    df_a_evaluar = pd.read_excel(lote_subido, dtype={"ID_Cliente": str, "ID_Operacion": str}) if lote_subido else df_tpl_lote

    try:
        df_res_lote = evaluar_lote_operaciones(df_a_evaluar, params)
        tot_saldo = df_res_lote["Saldo_Desembolsado"].sum()
        tot_apnr = df_res_lote["APNR"].sum()
        tot_cap = df_res_lote["Capital_Requerido"].sum()
        tot_util = df_res_lote["Utilidad_Neta"].sum()
        tot_eva = df_res_lote["EVA"].sum()
        raroc_port = (tot_util / tot_cap) if tot_cap > 0 else 0.0
        imp_pbs_port = df_res_lote["Impacto_Neto_Solvencia_pbs"].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Saldo Total Lote", f"${tot_saldo:,.1f}")
        m2.metric("APNR Agregado", f"${tot_apnr:,.1f}")
        m3.metric("RAROC Portafolio", f"{raroc_port*100:.2f}%")
        m4.metric("EVA Total Lote", f"${tot_eva:,.2f}")
        m5.metric("Impacto Solvencia Lote", f"{imp_pbs_port:+.3f} pbs")

        st.dataframe(df_res_lote, width="stretch", hide_index=True)

        buf_res_lote = BytesIO()
        with pd.ExcelWriter(buf_res_lote, engine="openpyxl") as writer:
            df_res_lote.to_excel(writer, sheet_name="Resultados_Portafolio", index=False)
        st.download_button(
            label="📥 Exportar Resultados del Portafolio (.xlsx)",
            data=buf_res_lote.getvalue(),
            file_name=f"resultados_lote_apnr_v2_{fecha_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except ValueError as exc:
        st.error(f"Error al procesar el lote: {exc}")

# ==============================================================================
# 7. PESTAÑA 3: EXPLORADOR Y VALIDADOR PARAMÉTRICO
# ==============================================================================
with tab3:
    st.markdown("### 🗂️ Tablas Paramétricas Activas y Reporte de Validación")
    if validacion.valido:
        st.success(f"✔️ Validación de esquema aprobada ({len(validacion.resumen_hojas)} hojas verificadas).")
    for adv in validacion.advertencias:
        st.warning(adv)

    hojas_vis = list(validacion.resumen_hojas.keys())
    hoja_sel = st.selectbox("Seleccionar hoja paramétrica para inspeccionar:", hojas_vis)
    st.caption(f"Registros cargados: **{len(params[hoja_sel])}**")
    st.dataframe(params[hoja_sel], width="stretch", hide_index=True, height=400)

# ==============================================================================
# 8. PESTAÑA 4: AUDITORÍA NORMATIVA Y CAMBIOS V1 -> V2
# ==============================================================================
with tab4:
    st.markdown(
        """
        ### 📖 Auditoría Normativa: Decreto 1477 de 2018, Decreto 573 de 2025, CBF SFC y Evolución V1 → V2

        1. **Art. 2.1.1.3.4 (Decreto 2555 de 2010) — Mitigación en Exposición ($E$):**
           - **Piso del 20% activado por norma:** $0.20 \\cdot A \\le E \\le A$ cuando aplican garantías admisibles no inmobiliarias.
           - **Parágrafo 1 y 2 (Garantías Inmobiliarias):** Las garantías inmobiliarias/hipotecarias (`BRCR`, `BDLI`) **no** reducen $E$ mediante $G(1 - F_g - F_c)$; su mitigación opera vía tablas **LTV** en el Art. 2.1.1.3.2. Cuando concurren una hipoteca y otra garantía admisible no inmobiliaria (Parágrafo 2), solo esta última reduce $E$.
           - **Idoneidad estricta en concurrencia:** Las garantías `NO IDONEA` nunca se suman al cálculo de $G_{\\text{comp}}$ ni alteran el $F_g$ efectivo.

        2. **Factor Cambiario ($F_c$) y Monedas (CBF Parte 2 Cap. 1):**
           - Se diferencia **$F_c = 8\\%$** para colateral financiero admisible y **$F_c = 30\\%$** para otras garantías idóneas cuando existe descalce de monedas.
           - Se incorpora **UVR** junto a `COP`, `USD` y `EUR`, y moneda independiente para la 2da garantía.

        3. **Art. 2.1.1.3.2 — Correcciones en Ponderaciones (RW) y Tablas LTV:**
           - **Calificaciones faltantes en V1:** Se agregan `B+`, `B`, `B-`, `CC`, `C`, `D` (en Grandes Empresas `< BB-` pondera **150%**; en Bancos/Soberanos `BB+ a B-` pondera **100%** y `< B-` **150%**).
           - **Calificaciones de Corto Plazo (`Calificacion_CP`):** Integradas para operaciones de corto plazo (`20%`, `50%`, `100%`, `150%`).
           - **Intervalos LTV sin brechas numéricas:** Se usan intervalos semiabiertos $[\\text{Lím. Inf}, \\text{Lím. Sup})$ eliminando el hueco entre `0.50` y `0.500001` que existía en V1.
           - **Límite minorista del 0.2% (Numeral 9):** Si una exposición Pyme/Persona Natural excede el 0.2% de la cartera minorista total (configurable mensualmente en `Metadata_Mensual`), pasa automáticamente al **100%**.
           - **Decreto 573 de 2025 (Libranza Consumo):** Ponderación preferencial parametrizable (**60%**, ajustable según transición 70%/65%/60%).

        4. **Art. 2.1.1.3.5 — Contingencias y Cupos No Utilizados (CCF):**
           - Soporte explícito para cupos aprobados no desembolsados (**10%**), garantías de cumplimiento (**50%**) y sustitutos directos/avales/cupos irrevocables (**100%**).

        5. **Módulo Completo de Rentabilidad y Pricing:**
           - Estado de resultados completo de la operación (**Intereses, Comisiones, FTP + Liquidez, Pérdida Esperada `PD×LGD`, Gastos Operativos e Impuestos**).
           - Cálculo de **RAROC**, **RORWA**, **ROA**, **EVA**, **Tasa Mínima Breakeven (RAROC = Ke)**, **Spread mínimo sobre FTP** e **Impacto Neto en Solvencia (APNR + Patrimonio Técnico en pbs)**.
        """
    )
