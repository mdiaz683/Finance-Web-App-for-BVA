### Version 4: incluir Open AI functions  y conectar con One Drive

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import plotly.graph_objects as go
from io import BytesIO
from collections import namedtuple

from openai import OpenAI
import json


# =========================
# Configuración de la página
# =========================
st.set_page_config(
    page_title="Budget Variance Analysis",
    page_icon="📊",
    layout="wide"
)
    

BUInfo = namedtuple("BUInfo", ["business_unit", "responsible", "department"])

# =========================
# Helper Functions
# =========================

def _find_sheet_case_insensitive(sheet_names: list, target: str) -> str:
    """Busca una sheet ignorando mayúsculas/minúsculas"""
    t = target.strip().lower()
    # Búsqueda exacta
    for s in sheet_names:
        if s and s.strip().lower() == t:
            return s
    # Búsqueda parcial
    for s in sheet_names:
        if s and t in s.strip().lower():
            return s
    return None

# =========================
# Cached Loading Functions - NIVEL 1 (MÁS RÁPIDO)
# =========================

@st.cache_data(show_spinner=False)
def _list_relevant_sheets(file_bytes: bytes):
    """
    RÁPIDO: Solo lista nombres de sheets sin cargar contenido.
    Devuelve (bva_sheets, jde_sheet_name)
    """
    xls = pd.ExcelFile(BytesIO(file_bytes))
    names = [s for s in xls.sheet_names if isinstance(s, str)]
    bva_sheets = [s for s in names if s.startswith("BvA")]
    jde_sheet = _find_sheet_case_insensitive(names, "JDE GL Act")
    return bva_sheets, jde_sheet


@st.cache_data(show_spinner=False)
def _list_database(file_bytes: bytes):
    """
    RÁPIDO: Solo lista nombres de sheets sin cargar contenido.
    Devuelve (db_sheet)
    """
    xls = pd.ExcelFile(BytesIO(file_bytes))
    names = [s for s in xls.sheet_names if isinstance(s, str)]
    db_sheet = _find_sheet_case_insensitive(names, "Data Base")
    return db_sheet

# =========================
# Cached Loading Functions - NIVEL 2 (CARGA SELECTIVA)
# =========================

@st.cache_data(show_spinner=False)
def _load_sheet_subset(file_bytes: bytes, sheet_name: str, 
                       min_row: int = None, max_row: int = None,
                       min_col: int = None, max_col: int = None):
    """
    Carga SOLO un subconjunto específico de celdas de una sheet.
    Mucho más rápido que cargar todo el workbook.
    """
    wb = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb[sheet_name]
    
    # Obtener valores del rango especificado
    result = {}
    for r_idx, row in enumerate(
        ws.iter_rows(min_row=min_row, max_row=max_row,
                     min_col=min_col, max_col=max_col,
                     values_only=True),              # <- clave
        start=min_row
    ):
        for c_idx, value in enumerate(row, start=min_col):
            result[(r_idx, c_idx)] = value
    wb.close()
    return result

@st.cache_data(show_spinner=False)
def _get_business_unit(file_bytes: bytes, sheet_name: str) -> BUInfo:
    """Lee C6 (BU), D6 (Responsible), C7 (Department) del sheet BvA."""
    # Solo cargar las celdas necesarias (filas 6-7, columnas C-D)
    cells = _load_sheet_subset(
        file_bytes, sheet_name,
        min_row=6, max_row=7,
        min_col=column_index_from_string("C"),
        max_col=column_index_from_string("D")
    )
    
    def get_val(row, col_letter):
        col = column_index_from_string(col_letter)
        v = cells.get((row, col))
        if v is None:
            return None
        s = str(v).strip()
        return s or None
    
    bu = get_val(6, "C")
    res = get_val(6, "D")
    dep = get_val(7, "C")
    
    return BUInfo(bu, res, dep)

@st.cache_data(show_spinner=False)
def _get_meses_y_colmap(file_bytes: bytes, sheet_name: str):
    """Extrae meses disponibles (fila 12, ET..FE) - SOLO esa fila"""
    col_inicio = column_index_from_string("ET")
    col_fin = column_index_from_string("FE")
    
    # Solo cargar fila 12 entre columnas ET y FE
    cells = _load_sheet_subset(
        file_bytes, sheet_name,
        min_row=12, max_row=12,
        min_col=col_inicio, max_col=col_fin
    )
    
    meses_disponibles = []
    col_map = {}
    
    for col in range(col_inicio, col_fin + 1):
        mes_nombre = cells.get((12, col))
        if mes_nombre:
            mes_nombre = str(mes_nombre).strip()
            meses_disponibles.append(mes_nombre)
            col_map[mes_nombre] = col
    
    return meses_disponibles, col_map

@st.cache_data(show_spinner=False)
def _get_bva_groups(file_bytes: bytes, sheet_name: str):
    """Obtiene lista única de BVA groups desde la columna ES (filas 13+)"""
    col_bva = column_index_from_string("ES")
    
    # Cargar solo columna ES desde fila 13 hasta 200 (suficiente margen)
    cells = _load_sheet_subset(
        file_bytes, sheet_name,
        min_row=13, max_row=29,
        min_col=col_bva, max_col=col_bva
    )
    
    filas_excluir = {24, 26, 29}
    bva_groups = []
    
    for row in range(13, 29):
        if row in filas_excluir:
            continue
        
        val = cells.get((row, col_bva))
        if val is None:
            break
        
        bva_groups.append(str(val).strip())
    
    # Únicos, manteniendo el orden
    return list(dict.fromkeys(bva_groups))


# =========================
# Data Processing Functions
# =========================

@st.cache_data(show_spinner=False)
def _crear_df_actual_agg(file_bytes: bytes, sheet_name: str, mes: str):
    """Crea DataFrame agregado de valores actuales - carga solo lo necesario"""
    _, col_map = _get_meses_y_colmap(file_bytes, sheet_name)
    
    if mes not in col_map:
        raise ValueError(f"Mes '{mes}' no encontrado en el archivo.")
    
    col_actual = col_map[mes]
    col_bva = column_index_from_string("ES")
    
    # Cargar solo columnas ES y la del mes seleccionado
    cells = _load_sheet_subset(
        file_bytes, sheet_name,
        min_row=13, max_row=29,
        min_col=min(col_bva, col_actual),
        max_col=max(col_bva, col_actual)
    )
    
    filas_excluir = {24, 26, 29}
    bva_groups = []
    actual_values = []
    
    for row in range(13, 29):
        if row in filas_excluir:
            continue
        
        bva_val = cells.get((row, col_bva))
        if bva_val is None:
            break
        
        bva = str(bva_val).strip()
        valor = cells.get((row, col_actual))
        
        bva_groups.append(bva)
        actual_values.append(valor)
    
    df = pd.DataFrame({
        "BVA group": bva_groups,
        "Actual Amount (USD)": actual_values
    })
    col_actual_name = f"Actual Amount (USD)"
    df[col_actual_name] = pd.to_numeric(df[col_actual_name], errors="coerce").fillna(0.0)
    
    return df

@st.cache_data(show_spinner=False)
def _crear_df_plan(file_bytes: bytes, sheet_name: str, mes: str):
    """Crea DataFrame de plan detallado - carga solo filas 32-104"""
    _, col_map = _get_meses_y_colmap(file_bytes, sheet_name)
    
    if mes not in col_map:
        raise ValueError(f"Mes '{mes}' no encontrado en el archivo.")
    
    col_actual = col_map[mes]
    col_eo = column_index_from_string("EO")
    col_eq = column_index_from_string("EQ")
    col_er = column_index_from_string("ER")
    col_es = column_index_from_string("ES")
    
    # Cargar solo filas 32-104 y columnas necesarias
    cells = _load_sheet_subset(
        file_bytes, sheet_name,
        min_row=32, max_row=104,
        min_col=col_eo,
        max_col=max(col_es, col_actual)
    )
    descriptions = []
    jde_accounts = []
    account_names = []
    bva_groups = []
    plan_values = []
    
    for row in range(32, 105):
        jde = cells.get((row, col_eq))
        if not jde:
            continue
        desc = cells.get((row, col_eo))
        acc = cells.get((row, col_er))
        bva = cells.get((row, col_es))
        val = cells.get((row, col_actual))
        
        descriptions.append(desc)
        jde_accounts.append(jde)
        account_names.append(acc)
        bva_groups.append(bva)
        plan_values.append(val)
    
    df = pd.DataFrame({
        "JDE Account": jde_accounts,
        "Account Name": account_names,
        "BVA group": bva_groups,
        "Short Description": descriptions,
        "Plan Amount (USD)": plan_values
    })
    df["Plan Amount (USD)"] = (pd.to_numeric(df['Plan Amount (USD)'], errors='coerce')
    .fillna(0))
    
    text_columns = ["JDE Account", "Account Name", "BVA group", "Short Description"]
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '').replace('None', '')

    df = df[["BVA group"] + [c for c in df.columns if c != "BVA group"]]
    return df


@st.cache_data(show_spinner=False)
def _crear_df_plan_agg(df_plan: pd.DataFrame, mes: str):
    """Crea DataFrame agregado de plan"""
    col_plan = "Plan Amount (USD)"
    df_plan[col_plan] = pd.to_numeric(df_plan[col_plan], errors="coerce")
    df_agg = df_plan.groupby("BVA group", as_index=False)[col_plan].sum()
    return df_agg


# =========================
# Cached Loading Functions - NIVEL 3 (PANDAS - RÁPIDO PARA TABLAS)
# =========================

@st.cache_data(show_spinner=False)
def _load_database_df(file_bytes: bytes) -> pd.DataFrame:
    """
    Carga 'Data Base' completa con PANDAS (mucho más rápido que openpyxl).
    Solo se ejecuta una vez y se cachea.
    """
    db_sheet = _list_database(file_bytes)
    
    if not db_sheet:
        raise ValueError("Unable to find sheet 'Data Base'")
    
    # Usar pandas para leer - ES MUCHO MÁS RÁPIDO
    xls = pd.ExcelFile(BytesIO(file_bytes))
    
    # Leer con columnas específicas por letra
    letters = ["C", "E", "I", "J", "BB", "BD", "BZ", "CD", "CE"]
    df = pd.read_excel(
        xls,
        sheet_name=db_sheet,
        header=5,  # fila 6 como header
        usecols=",".join(letters),
        dtype=object
    )
    df.columns = df.columns.str.strip()
    # Obtener nombres reales de columnas y mapear a canónicos
    actual_cols = df.columns.tolist()
    canonical = ["Business Unit", "JDE Account", "Amount", "Explanation Alpha Name", "Invoice Number", "Purchase Order", "Month", "Account Name", "BVA group"]
    rename_map = dict(zip(actual_cols, canonical))
    df.rename(columns=rename_map, inplace=True)
    
    # Normalizar tipos
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    text_columns = ["Business Unit", "BVA group", "JDE Account", "Account Name", 
                    "Explanation Alpha Name", "Purchase Order", "Invoice Number", "Month"]
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace('nan', '').replace('None', '')
    
    cols = ["Month", "Business Unit", "BVA group", "JDE Account", "Account Name", "Explanation Alpha Name", "Purchase Order", "Invoice Number", "Amount"]
    return df[[c for c in cols if c in df.columns]]

@st.cache_data(show_spinner=False)
def _db_actual_det(file_bytes: bytes, business_unit: str, mes: str) -> tuple[pd.DataFrame, bool]:
    """
    Obtiene detalles de actual desde Data Base del otro book.
    Retorna: (DataFrame, tiene_datos)
    """
    df_db = _load_database_df(file_bytes)
    bu = str(business_unit).strip()
    m = str(mes).strip()
    cols = ["BVA group", "JDE Account", "Account Name", "Explanation Alpha Name", "Purchase Order", "Invoice Number", "Amount"]
    
    mask = (df_db["Business Unit"] == bu) & (df_db["Month"] == m)
    out = df_db.loc[mask, cols].copy()
    
    tiene_datos = len(out) > 0
    
    if tiene_datos:
        out = out[cols]
        out.rename(columns={"Amount": "Actual Amount (USD)"}, inplace=True)
        return out, tiene_datos
    else:
        # Retornar DataFrame vacío con estructura correcta
        empty_df = pd.DataFrame(columns=cols)
        return empty_df, tiene_datos
    
def get_variance(actual, plan):
    variance = actual - plan
    if variance < 0:
        variance_label = "Variance — Underspent"
    elif variance > 0:
        variance_label = "Variance — Overspent"
    else:
        variance_label = "Variance — On plan"
    if plan != 0:
        variance_pct = (variance / plan) * 100
        delta_str = f"{variance_pct:.2f}%"
    else:
        delta_str = "—"

    metric = st.metric(
        label=variance_label,
        value=f"${variance:,.2f}",
        delta=delta_str,
        delta_color="inverse",  # positivo = rojo, negativo = verde
        help="Actual - Plan (negative = underspent, positive = overspent)"
    )
    return metric


# =========================
# OpenAI Report Generation
# =========================

def generate_analysis_report(df_actual: pd.DataFrame, df_plan: pd.DataFrame, 
                            bva_group: str, jde_account: str, 
                            total_actual: float, total_plan: float) -> dict:
    """
    Genera un reporte de análisis usando OpenAI con los datos de Actual vs Plan.
    Retorna un diccionario con el análisis y recomendaciones estructuradas.
    """
    
    # Convertir DataFrames a CSV string
    actual_csv = df_actual.to_csv(index=False) if not df_actual.empty else "No actual data available"
    plan_csv = df_plan.to_csv(index=False) if not df_plan.empty else "No plan data available"
    
    # Calcular varianza
    variance = total_actual - total_plan
    variance_pct = ((variance / total_plan) * 100) if total_plan != 0 else 0
    
    # Determinar estado
    if variance < 0:
        status = "UNDERSPENT"
    elif variance > 0:
        status = "OVERSPENT"
    else:
        status = "ON TARGET"
    
    # Construir contexto
    context = f"""
    BVA Group: {bva_group}
    JDE Account Filter: {jde_account}
    
    FINANCIAL SUMMARY:
    - Total Actual: ${total_actual:,.2f}
    - Total Plan: ${total_plan:,.2f}
    - Variance: ${variance:,.2f} ({variance_pct:.2f}%)
    - Status: {status}
    """
    
    # Prompt del sistema
    system_prompt = """You are an expert financial analyst specializing in budget variance analysis. 
Your role is to analyze actual expenses versus planned budgets and provide actionable insights.

Your analysis should:
1. Identify specific line items that contribute most to variances
2. Match actual expenses to their corresponding budget line items
3. Highlight discrepancies and potential misclassifications
4. Provide recommendations for budget adjustments or expense reallocations
5. Flag unusual patterns or outliers

Focus on being specific, data-driven, and actionable."""

    # Prompt del usuario
    user_prompt = f"""
Please analyze the following Budget vs Actual data:

{context}

ACTUAL EXPENSES DATA (CSV):
{actual_csv}

PLANNED BUDGET DATA (CSV):
{plan_csv}

Please provide a structured analysis including:

1. **Executive Summary**: Brief overview of the variance situation

2. **Detailed Variance Analysis**: 
   - Which specific line items are driving the variance?
   - Are there expenses in Actual that don't have corresponding budget items?
   - Are there budget items with no actual spending?

3. **Matching Recommendations**: 
   - Suggest which actual expenses should match with which budget line items
   - Identify potential misclassifications
   - Highlight any JDE account discrepancies

4. **Root Cause Analysis**:
   - Why is this category {status}?
   - What are the main contributing factors?

5. **Actionable Recommendations**:
   - Should the budget be adjusted?
   - Are there expenses that should be reclassified?
   - What actions should the budget owner take?

Be specific and reference actual amounts and line items from the data.
"""

    try:
        # Llamada a OpenAI con structured outputs
        response = client.chat.completions.create(
            model="gpt-4o-2024-08-06",  # Modelo con structured outputs
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            response_format={
                "type": "json_schema",
                "json_schema": {
                    "name": "budget_analysis",
                    "strict": True,
                    "schema": {
                        "type": "object",
                        "properties": {
                            "executive_summary": {
                                "type": "string",
                                "description": "Brief overview of the variance situation"
                            },
                            "variance_drivers": {
                                "type": "array",
                                "description": "Main items driving the variance",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "item": {"type": "string"},
                                        "impact": {"type": "string"},
                                        "amount": {"type": "number"}
                                    },
                                    "required": ["item", "impact", "amount"],
                                    "additionalProperties": False
                                }
                            },
                            "matching_recommendations": {
                                "type": "array",
                                "description": "Suggested matches between actual and plan",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "actual_item": {"type": "string"},
                                        "suggested_plan_item": {"type": "string"},
                                        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                                        "reasoning": {"type": "string"}
                                    },
                                    "required": ["actual_item", "suggested_plan_item", "confidence", "reasoning"],
                                    "additionalProperties": False
                                }
                            },
                            "root_causes": {
                                "type": "array",
                                "description": "Identified root causes for variance",
                                "items": {
                                    "type": "string"
                                }
                            },
                            "actionable_recommendations": {
                                "type": "array",
                                "description": "Specific actions to take",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "action": {"type": "string"},
                                        "priority": {"type": "string", "enum": ["high", "medium", "low"]},
                                        "impact": {"type": "string"}
                                    },
                                    "required": ["action", "priority", "impact"],
                                    "additionalProperties": False
                                }
                            },
                            "anomalies": {
                                "type": "array",
                                "description": "Unusual patterns or outliers detected",
                                "items": {
                                    "type": "string"
                                }
                            }
                        },
                        "required": [
                            "executive_summary",
                            "variance_drivers",
                            "matching_recommendations",
                            "root_causes",
                            "actionable_recommendations",
                            "anomalies"
                        ],
                        "additionalProperties": False
                    }
                }
            }
        )
        
        # Extraer contenido JSON
        analysis = json.loads(response.choices[0].message.content)
        
        return {
            "success": True,
            "analysis": analysis,
            "tokens_used": response.usage.total_tokens
        }
        
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


# =========================
# UI - Sidebar
# =========================

st.title("📊 Budget Variance Analysis")
st.markdown("---")

# Initialize variables
uploaded_file = None
uploaded_file2 = None
file_bytes = None
file2_bytes = None

with st.sidebar:
    st.header("⚙️ Configuration")
    
    with st.expander("📁 Select Files", expanded=True):
        file_source = st.radio(
            "File source:",
            ["Upload manually", "Load from SharePoint"],
            key="file_source"
        )
        
        if file_source == "Upload manually":
            uploaded_file = st.file_uploader("📗 'FY26 QRA BvAs' File", type=['xlsx'], key="file1")
            uploaded_file2 = st.file_uploader("📗 'JDE Details QRA (YTD)' File", type=['xlsx'], key="file2")
            
            if uploaded_file and uploaded_file2:
                file_bytes = uploaded_file.getvalue()
                file2_bytes = uploaded_file2.getvalue()

    # Process files if available from manual upload
    if file_source == "Upload manually" and uploaded_file and uploaded_file2:
        pass  # file_bytes already set above
    
    # Process files if available from either source
    if (uploaded_file and uploaded_file2) or st.session_state.get('files_loaded', False):
        # Get file_bytes from session state if loaded from SharePoint
        if st.session_state.get('files_loaded', False) and not file_bytes:
            file_bytes = st.session_state.file_bytes
            file2_bytes = st.session_state.file2_bytes
            uploaded_file = True
            uploaded_file2 = True
        
        try:
            # PASO 1: Solo listar sheets (súper rápido)
            with st.spinner("Scanning first file..."):
                bva_sheets, jde_sheet = _list_relevant_sheets(file_bytes)
            
            with st.spinner("Scanning second file..."):
                db_sheet = _list_database(file2_bytes)

            if not bva_sheets:
                st.error("❌ No pages were found beginning with 'BvA'")
                st.stop()
            
            if not jde_sheet:
                st.error("❌ The required 'JDE GL Act' sheet was not found")
                st.stop()

            if not db_sheet:
                st.error("❌ The 'Data Base' sheet was not found in the current data file")
                st.stop()
            
            # Selección de sheet
            st.markdown("---")
            selected_sheet = st.selectbox("Select BvA Sheet", bva_sheets)
            
            # PASO 2: Cargar solo info del BU (celdas específicas)
            NOINFO = "No information"
            fmt = lambda x: x if (x and str(x).strip()) else NOINFO

            with st.spinner("Loading business unit info..."):
                bu_info = _get_business_unit(file_bytes, selected_sheet)
                business_unit = bu_info.business_unit
                responsible = bu_info.responsible
                department = bu_info.department

            st.info(
                f"**Business Unit:** {fmt(business_unit)}  \n"
                f"**CC Responsible:** {fmt(responsible)}  \n"
                f"**Department:** {fmt(department)}"
            )
            
            # PASO 3: Cargar solo fila de meses
            st.markdown("---")
            with st.spinner("Loading months..."):
                meses_disponibles, _ = _get_meses_y_colmap(file_bytes, selected_sheet)
            
            default_index = meses_disponibles.index("Aug") if "Aug" in meses_disponibles else 0
            mes_seleccionado = st.selectbox("Select Month", meses_disponibles, index=default_index)
            
            # Tipo de vista
            st.markdown("---")
            st.subheader("View Type")
            tipo_vista = st.radio(
                "Select Report Type",
                ["Summary Comparison", "Bva Group Details", "Side-by-Side Analysis"],
                label_visibility="collapsed"
            )
        
        except Exception as e:
            st.error(f"Failed processing the files: {str(e)}")
            st.exception(e)
            st.stop()


# =========================
# UI - Main Content
# =========================

if uploaded_file and uploaded_file2:
    try:
        # Cargar datos solo cuando sean necesarios
        with st.spinner("Loading analysis data..."):
            df_actual_agg = _crear_df_actual_agg(file_bytes, selected_sheet, mes_seleccionado)
            df_plan = _crear_df_plan(file_bytes, selected_sheet, mes_seleccionado)
            df_plan_agg = _crear_df_plan_agg(df_plan, mes_seleccionado)
            df_actual, tiene_datos_actual = _db_actual_det(file2_bytes, business_unit, mes_seleccionado)

        col_actual_name = "Actual Amount (USD)"
        col_plan_name = "Plan Amount (USD)"

        if not tiene_datos_actual:
            from datetime import datetime
            st.warning(
                f"⚠️ **No actual expense data available for {mes_seleccionado}**  \n"
                f"This month may not have been closed yet or no transactions have been recorded.  \n"
                f"*Current date: {datetime.now().strftime('%B %d, %Y')}*"
            )
        
        # =========================
        # VISTA: Summary Comparison
        # =========================
        if tipo_vista == "Summary Comparison":
            st.header(f"📈 Summary Comparison - {mes_seleccionado} | BU: {business_unit}")
            
            # Métricas principales
            col1, col2, col3 = st.columns(3)
            
            total_actual = pd.to_numeric(df_actual_agg[col_actual_name], errors="coerce").sum()
            total_plan = pd.to_numeric(df_plan_agg[col_plan_name], errors="coerce").sum()

            with col1:
                st.metric("Total Actual", f"${total_actual:,.2f}")
            with col2:
                st.metric("Total Plan", f"${total_plan:,.2f}")
            with col3:
                get_variance(total_actual, total_plan)
            
            st.markdown("---")
            
            # Tabla comparativa
            df_comp = pd.merge(
                df_actual_agg,
                df_plan_agg,
                on='BVA group',
                how='left'  
            ).fillna(0)
            
            df_comp[col_actual_name] = pd.to_numeric(df_comp[col_actual_name], errors="coerce")
            df_comp[col_plan_name] = pd.to_numeric(df_comp[col_plan_name], errors="coerce")
            df_comp['Variance'] = df_comp[col_actual_name] - df_comp[col_plan_name]
            #df_comp['Variance %'] = ((df_comp['Variance'] / df_comp[col_plan_name].replace(0, 1)) * 100).round(2)
            
            st.subheader("Comparison by BVA Group")
            st.dataframe(df_comp, width='stretch', height=500)
            df_comp['Total'] = df_comp[col_actual_name] + df_comp[col_plan_name]
            df_comp = df_comp.sort_values('Total', ascending=True)

            # Gráfico comparativo
            st.subheader("Visual Comparison")
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='Actual',
                x=df_comp[col_actual_name],
                y=df_comp['BVA group'],
                marker_color='#ff7f0e',
                orientation='h'
            ))
            fig.add_trace(go.Bar(
                name='Plan',
                x=df_comp[col_plan_name],
                y=df_comp['BVA group'],
                marker_color='#1f77b4',
                orientation='h'
            ))
            fig.update_layout(
                barmode='group',
                xaxis_title='Amount (USD)',
                yaxis_title='BVA group',
                height=600,
                width = 500,
            )
            st.plotly_chart(fig, config={'displayModeBar': False})

        # =========================
        # VISTA: Full Detail
        # =========================
        elif tipo_vista == "Bva Group Details":
            st.header(f"📋 Bva Group Details - {mes_seleccionado} | BU: {business_unit}")
            
            tab1, tab2 = st.tabs(["💰 Actual Detail", "📝 Plan Detail"])
            
            with tab1:
                st.subheader("Actual Transactions")
                st.dataframe(df_actual, width='stretch', height=400)
                
                st.subheader("Actual Summary by BVA Group")
                df_actual_summary = df_actual.groupby('BVA group', as_index=False)[col_actual_name].sum()
                st.dataframe(df_actual_summary, width='stretch')
                
                st.metric("Total Actual", f"${df_actual[col_actual_name].sum():,.2f}")
            
            with tab2:
                st.subheader("Plan Budget")
                st.dataframe(df_plan, width='stretch', height=400)
                
                st.subheader("Plan Summary by BVA Group")
                st.dataframe(df_plan_agg, width='stretch')
                
                st.metric("Total Plan", f"${df_plan_agg[col_plan_name].sum():,.2f}")
        
        # =========================
        # VISTA: BVA Group Detail
        # =========================
        else:
            st.header(f"🔍 Side-by-Side Analysis - {mes_seleccionado} | BU: {business_unit}")
            
            # FILTROS EN LA PANTALLA PRINCIPAL
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                # Obtener BVA groups disponibles
                bva_groups = _get_bva_groups(file_bytes, selected_sheet)
                if not bva_groups:
                    st.warning("No BVA groups were found in the selected sheet.")
                    st.stop()
                
                bva_seleccionado = st.selectbox(
                    "🏢 Select BVA Group",
                    bva_groups,
                    help="Choose a BVA group to analyze"
                )
            
            with col_filter2:
                # Filtrar datos para el BVA group seleccionado
                df_plan_filtered = df_plan[df_plan['BVA group'] == bva_seleccionado]
                df_actual_filtered = df_actual[df_actual['BVA group'] == bva_seleccionado]
                
                # Crear mapeo JDE Account -> Account Name
                # Combinamos ambos dataframes para tener todos los mapeos posibles
                df_combined = pd.concat([
                    df_plan_filtered[['JDE Account', 'Account Name']],
                    df_actual_filtered[['JDE Account', 'Account Name']]
                ]).drop_duplicates()
                
                jde_account_map = dict(zip(df_combined['JDE Account'], df_combined['Account Name']))
                
                # Obtener JDE Accounts que aparecen en CUALQUIERA de las dos tablas (unión)
                jde_accounts_plan = set(df_plan_filtered['JDE Account'].unique())
                jde_accounts_actual = set(df_actual_filtered['JDE Account'].unique())
                jde_accounts_todas = sorted(list(jde_accounts_plan.union(jde_accounts_actual)))
                
                # Crear opciones con formato "JDE - Account Name"
                jde_options_formatted = ["All Accounts"] + [
                    f"{jde} - {jde_account_map.get(jde, 'Unknown')}" 
                    for jde in jde_accounts_todas
                ]
                
                jde_seleccionado_formatted = st.selectbox(
                    "🔢 Select JDE Account",
                    jde_options_formatted,
                    help="Filter by specific JDE Account or view all (showing accounts from both Actual and Plan)"
                )
                
                # Extraer solo el código JDE de la selección
                if jde_seleccionado_formatted != "All Accounts":
                    jde_seleccionado = jde_seleccionado_formatted.split(" - ")[0]
                else:
                    jde_seleccionado = "All Accounts"

            st.markdown("---")
            
            # Filtrar datos según selección
            df_actual_bva = df_actual[df_actual['BVA group'] == bva_seleccionado].copy()
            df_plan_bva = df_plan_filtered.copy()
            
            # Aplicar filtro de JDE Account si no es "All Accounts"
            if jde_seleccionado != "All Accounts":
                df_actual_bva = df_actual_bva[df_actual_bva['JDE Account'] == jde_seleccionado]
                df_plan_bva = df_plan_bva[df_plan_bva['JDE Account'] == jde_seleccionado]
            
            # Mostrar título con filtros aplicados
            filter_text = f"{bva_seleccionado}"
            if jde_seleccionado != "All Accounts":
                # Usar el formato completo para el título
                filter_text += f" | JDE: {jde_seleccionado_formatted}"
            

            st.subheader(f"📊 {filter_text}")
            
            # CSS para tablas compactas
            st.markdown("""
                <style>
                /* Hacer las tablas más compactas */
                .stDataFrame {
                    font-size: 0.7rem !important;
                }
                /* Reducir padding de celdas */
                .stDataFrame td, .stDataFrame th {
                    padding: 2px 4px !important;
                }
                </style>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns([1, 1], gap="medium")
            
            with col1:
                st.markdown("#### 💰 Actual")
                if not df_actual_bva.empty:
                    # Crear una copia con nombres de columna más cortos - SIN Account Name
                    df_display_actual = df_actual_bva[['JDE Account', "Explanation Alpha Name", "Purchase Order", "Invoice Number", col_actual_name]].copy()
                    df_display_actual.columns = ['JDE', 'Explanation Alpha Name', 'PO', 'Invoice', '$']
                    
                    st.dataframe(
                        df_display_actual,
                        width='stretch',
                        height=400,
                        hide_index=True
                    )
                    total_actual_bva = df_actual_bva[col_actual_name].sum()
                    st.metric("Total Actual", f"${total_actual_bva:,.2f}")
                else:
                    st.info("No actual data for this selection")
                    total_actual_bva = 0
            
            with col2:
                st.markdown("#### 📄 Plan")
                if not df_plan_bva.empty:
                    # Crear una copia con nombres de columna más cortos - SIN Account Name
                    df_display_plan = df_plan_bva[['JDE Account', 'Short Description', col_plan_name]].copy()
                    df_display_plan.columns = ['JDE', 'Description', '$']
                    
                    st.dataframe(
                        df_display_plan,
                        width='stretch',
                        height=400,
                        hide_index=True
                    )
                    total_plan_bva = df_plan_bva[col_plan_name].sum()
                    st.metric("Total Plan", f"${total_plan_bva:,.2f}")
                else:
                    st.info("No plan data for this selection")
                    total_plan_bva = 0


            # Comparativa
            st.markdown("---")
            # Botón para generar reporte con OpenAI
            col_button = st.columns([1, 2, 1])[1]
            with col_button:
                if st.button("🤖 Generate AI Analysis Report", type="primary", use_container_width=True):
                    with st.spinner("🔄 Analyzing data with AI..."):
                        report_result = generate_analysis_report(
                            df_actual=df_actual_bva,
                            df_plan=df_plan_bva,
                            bva_group=bva_seleccionado,
                            jde_account=jde_seleccionado,
                            total_actual=total_actual_bva,
                            total_plan=total_plan_bva
                        )

                    if report_result["success"]:
                        st.success(f"✅ Analysis completed! (Tokens used: {report_result['tokens_used']})")

                        analysis = report_result["analysis"]

                        # Mostrar análisis en expanders
                        with st.expander("📊 Executive Summary", expanded=True):
                            st.markdown(analysis["executive_summary"])

                        with st.expander("📈 Variance Drivers", expanded=True):
                            if analysis["variance_drivers"]:
                                for driver in analysis["variance_drivers"]:
                                    st.markdown(f"""
                                    **{driver['item']}**  
                                    Impact: {driver['impact']}  
                                    Amount: ${driver['amount']:,.2f}
                                    """)
                                    st.markdown("---")
                            else:
                                st.info("No significant variance drivers identified")

                        with st.expander("🔗 Matching Recommendations", expanded=True):
                            if analysis["matching_recommendations"]:
                                # Crear DataFrame para mejor visualización
                                matches_df = pd.DataFrame(analysis["matching_recommendations"])

                                # Colorear por confidence
                                def highlight_confidence(row):
                                    colors = {
                                        'high': 'background-color: #d4edda',
                                        'medium': 'background-color: #fff3cd', 
                                        'low': 'background-color: #f8d7da'
                                    }
                                    return [colors.get(row['confidence'], '')] * len(row)

                                st.dataframe(
                                    matches_df.style.apply(highlight_confidence, axis=1),
                                    use_container_width=True,
                                    hide_index=True
                                )
                            else:
                                st.info("No matching recommendations available")

                        with st.expander("🔍 Root Cause Analysis"):
                            if analysis["root_causes"]:
                                for i, cause in enumerate(analysis["root_causes"], 1):
                                    st.markdown(f"{i}. {cause}")
                            else:
                                st.info("No root causes identified")

                        with st.expander("✅ Actionable Recommendations"):
                            if analysis["actionable_recommendations"]:
                                # Agrupar por prioridad
                                high = [r for r in analysis["actionable_recommendations"] if r["priority"] == "high"]
                                medium = [r for r in analysis["actionable_recommendations"] if r["priority"] == "medium"]
                                low = [r for r in analysis["actionable_recommendations"] if r["priority"] == "low"]

                                if high:
                                    st.markdown("#### 🔴 High Priority")
                                    for rec in high:
                                        st.markdown(f"- **{rec['action']}**  \n  Impact: {rec['impact']}")

                                if medium:
                                    st.markdown("#### 🟡 Medium Priority")
                                    for rec in medium:
                                        st.markdown(f"- **{rec['action']}**  \n  Impact: {rec['impact']}")

                                if low:
                                    st.markdown("#### 🟢 Low Priority")
                                    for rec in low:
                                        st.markdown(f"- **{rec['action']}**  \n  Impact: {rec['impact']}")
                            else:
                                st.info("No recommendations available")

                        with st.expander("⚠️ Anomalies Detected"):
                            if analysis["anomalies"]:
                                for anomaly in analysis["anomalies"]:
                                    st.warning(anomaly)
                            else:
                                st.success("No anomalies detected")

                        # Opción para descargar el reporte completo
                        st.markdown("---")
                        col_download = st.columns([1, 2, 1])[1]
                        with col_download:
                            json_str = json.dumps(analysis, indent=2)
                            st.download_button(
                                label="📥 Download Full Report (JSON)",
                                data=json_str,
                                file_name=f"analysis_report_{bva_seleccionado}_{mes_seleccionado}.json",
                                mime="application/json",
                                use_container_width=True
                            )

                    else:
                        st.error(f"❌ Error generating report: {report_result['error']}")

            st.markdown("---")

            col3, col4, col5 = st.columns([1, 1, 1.5])
            
            with col3:
                get_variance(total_actual_bva, total_plan_bva)

            with col4:
                if total_actual_bva > 0 or total_plan_bva > 0:
                    fig_pie = go.Figure(data=[go.Pie(
                        labels=['Actual', 'Plan'],
                        values=[total_actual_bva, total_plan_bva],
                        hole=.3,
                        marker_colors=['#ff7f0e', '#1f77b4']
                    )])
                    fig_pie.update_layout(
                        title=f'Plan vs Actual by BVA Group - {filter_text}',
                        height=300,
                        margin=dict(t=40, r=0, b=0, l=0)
                    )
                    st.plotly_chart(fig_pie, config={'displayModeBar': False})
                else:
                    st.info("No data available to display distribution chart")
            
            with col5:
                # Gráfico de barras comparativo por JDE Account
                if not df_actual_bva.empty or not df_plan_bva.empty:
                    # Agrupar por JDE Account
                    actual_by_account = df_actual_bva.groupby('JDE Account')[col_actual_name].sum().reset_index()
                    actual_by_account.columns = ['JDE Account', 'Actual']
                    
                    plan_by_account = df_plan_bva.groupby('JDE Account')[col_plan_name].sum().reset_index()
                    plan_by_account.columns = ['JDE Account', 'Plan']
                    
                    # Merge para tener ambas columnas
                    comparison_df = pd.merge(
                        actual_by_account, 
                        plan_by_account, 
                        on='JDE Account', 
                        how='outer'
                    ).fillna(0)
                    comparison_df['JDE Account'] = comparison_df['JDE Account'].astype(str).str.strip()

                    # Ordenar por total (Actual + Plan) descendente
                    comparison_df['Total'] = comparison_df['Actual'] + comparison_df['Plan']
                    comparison_df = comparison_df.sort_values('Total', ascending=True)

                    # Calcular altura dinámica basada en el número de cuentas
                    num_accounts = len(comparison_df)
                    # Mínimo 300px, añadir 40px por cada cuenta adicional después de 5
                    dynamic_height = max(300, 300 + (num_accounts - 5) * 40) if num_accounts > 5 else 300

                    # Crear gráfico de barras horizontal
                    fig_bar = go.Figure()
                    
                    fig_bar.add_trace(go.Bar(
                        name='Plan',
                        y=comparison_df['JDE Account'],
                        x=comparison_df['Plan'],
                        orientation='h',
                        marker_color='#1f77b4',
                        text=comparison_df['Plan'].apply(lambda x: f'${x:,.0f}' if x > 0 else ''),
                        textposition='auto',
                        hovertemplate='<b>%{y}</b><br>Plan: $%{x:,.2f}<extra></extra>'
                    ))
                    
                    fig_bar.add_trace(go.Bar(
                        name='Actual',
                        y=comparison_df['JDE Account'],
                        x=comparison_df['Actual'],
                        orientation='h',
                        marker_color='#ff7f0e',
                        text=comparison_df['Actual'].apply(lambda x: f'${x:,.0f}' if x > 0 else ''),
                        textposition='auto',
                        hovertemplate='<b>%{y}</b><br>Actual: $%{x:,.2f}<extra></extra>'
                    ))
                    
                    fig_bar.update_layout(
                        title='Plan vs Actual by JDE Account',
                        barmode='group',
                        height=dynamic_height,
                        xaxis_title='USD',
                        yaxis_title='JDE Account',
                        margin=dict(t=40, r=0, b=40, l=80),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        ),
                        yaxis=dict(
                            type='category',  # Forzar que el eje Y trate los valores como categorías
                            automargin=True
                        )
                    )
                    
                    st.plotly_chart(fig_bar, config={'displayModeBar': False})
                else:
                    st.info("No data to display")


    except Exception as e:
        st.error(f"❌ Error processing data: {str(e)}")
        st.exception(e)

else:
    st.info("👈 Please upload or select files from SharePoint to begin the analysis")
    
    with st.expander("ℹ️ Instructions"):
        st.markdown("""
        ### How to use this application:
        
        1. **Upload files**: 
           - **Budget File**: Excel with BvA sheets containing plan data
           - **Actual Data File**: Excel with 'Data Base' sheet containing actual transactions
        2. **Select BvA sheet**: Choose the Business Unit sheet to analyze
        3. **Select month**: Choose the month for the analysis
        4. **Choose view**:
           - **Summary Comparison**: Overview with metrics and charts
           - **Bva Group Details**: All transactions for Actual and Plan
           - **Side-by-Side Analysis**: Specific analysis of one BVA group
        
        ### Expected file structure:
        
        **Budget File:**
        - Must contain sheets starting with "BvA" (e.g., "BvA Ster Assurance 1110551")
        - Business Unit in cell C6 of the BvA sheet
        - Months in row 12 (columns ET to FE)
        - BVA groups starting at row 13 (column ES)
        
        **Actual Data File:**
        - Must contain sheet "Data Base" with transaction details
        - Columns: Business Unit, JDE Account, Amount, Month, Account Name, BVA group, etc.
        """)

st.markdown("---")
st.markdown("*Budget Analysis Dashboard v1.2 - Optimized*")