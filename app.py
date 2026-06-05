import streamlit as st
import requests
import base64
import os
import pandas as pd
import numpy as np
import plotly.express as px
import io
import xlsxwriter
import sys
import time
import re
import plotly.offline as opy
from fpdf import FPDF
from datetime import datetime

# --- PASO 1: DEFINIR LA FUNCIÓN DE RUTA (EL "GPS") ---
# Esta función DEBE ir antes que cualquier otra cosa
def obtener_ruta_recurso(nombre_archivo):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nombre_archivo)


# --- CONFIGURACIÓN DE REGLAS DE NEGOCIO DINAMICA---
from datetime import datetime

# 1. Definimos las reglas (puedes agregar más deptos aquí en el futuro)
DICCIONARIO_EXCLUSIONES = {
    "304": "04/2026"
}

# 1.2 Definimos la función lógica
def calcular_excluidos_del_mes(mes_consulta, dict_reglas):
    lista_final = []
    try:
        # Convertimos el mes seleccionado (ej: "04/2026") a fecha para comparar
        fecha_sel = datetime.strptime(mes_consulta, "%m/%Y")
        
        for unidad, mes_limite in dict_reglas.items():
            fecha_limite = datetime.strptime(mes_limite, "%m/%Y")
            
            # Si el mes consultado es menor o igual al límite, se queda en la lista
            if fecha_sel <= fecha_limite:
                lista_final.append(unidad)
    except:
        pass # En caso de error de formato, lista vacía
    return lista_final

# 1.3 Calculamos la lista para el periodo actual
# IMPORTANTE: Esta línea va en el lugar donde se define el 'mes_seleccionado'
# UNIDADES_EXCLUIDAS = calcular_excluidos_del_mes(mes_seleccionado, DICCIONARIO_EXCLUSIONES)

# 1.4 Columnas financieras (sin cambios)
COLS_FINANCIERAS = [
    'Pago Total del Mes', 'Monto Base', 'Monto Actual', 
    'Pagos Adelantados', 'Cuotas Extras', 'Otros Ingresos', 'Pago Mensual'
]


# --- 2. DEFINICIÓN DE LA FUNCIÓN PARA HTML ---
def generar_portal_web(df_resumen, df_maestra, metricas, df_mora_mes, df_mora_hist, m_salud, 
                       div_tendencia, div_torta, salud_data, mes_actual, fig_gauge, 
                       df_salud_final, explicacion_mil, df_control, pdf_b64=None, pdf_gestion_b64=None, 
                       unidades_excluidas=None, fecha_actualizacion="No disponible",
                       # Agregamos valores por defecto para que la función acepte llamadas con solo 18 datos
                       ver_balance=True, ver_resumen_fin=True, ver_morosidad=True, 
                       ver_control=True, ver_pdfs=True):
    import plotly.offline as opy
    import pandas as pd
    import re
    from datetime import datetime

    # --- CONFIGURACIÓN DE ANALÍTICA (Counter.dev) ---
    id_counter = "061a02f0-8a43-4280-a56b-1e5cf920b409"  
    script_analitica = f'<script src="https://cdn.counter.dev/script.js" data-id="{id_counter}" data-utcoffset="-4"></script>'
    
    # --- MEJORA CRÍTICA: Asegurar que detecte la exclusión ---
    if unidades_excluidas is None:
        try:
            lista_final = UNIDADES_EXCLUIDAS
        except NameError:
            lista_final = []
    else:
        lista_final = unidades_excluidas

    # 1. Preparación de todas las tablas HTML
    t_resumen = df_resumen.to_html(classes='table table-hover mb-0', index=False, border=0, justify='left')
    
    # --- [CAMBIO 1] CORRECCIÓN PARA DETALLE DE MOVIMIENTOS (PRIVACIDAD NOMINAL) ---
    df_maestra_copy = df_maestra.copy()

    # --- NUEVO: BLOQUE DE ORDENAMIENTO PARA EL PORTAL ---
    if 'Unidad' in df_maestra_copy.columns:
        # Convertimos a numérico para ordenar correctamente (1, 2, 10 en lugar de 1, 10, 2)
        df_maestra_copy['Unidad_Sort'] = pd.to_numeric(df_maestra_copy['Unidad'], errors='coerce')
        # Ordenamos: Números de unidad primero y NaNs (Gastos Generales) al final
        df_maestra_copy = df_maestra_copy.sort_values(by='Unidad_Sort', ascending=True, na_position='last')
        
        # Limpieza estética: Nos aseguramos de que no aparezcan decimales (.0) en el HTML
        df_maestra_copy['Unidad'] = df_maestra_copy['Unidad_Sort'].fillna(0).astype(int).astype(str).replace('0', '-')
        
        # Eliminamos la columna auxiliar
        df_maestra_copy = df_maestra_copy.drop(columns=['Unidad_Sort'])
    # ---------------------------------------------------
    
    def anonimizar_detalle(row):
        detalle_orig = str(row.get('Detalle', ''))
        sub_cat = str(row.get('Sub-Categoría', '')).strip()
        flujo = str(row.get('Flujo', ''))

        # Si es un ingreso y la sub-categoría es "Departamento"
        if "(+) INGRESO" in flujo and sub_cat == "Departamento":
            # Extraemos el número de la unidad (ej: de "201 - Daniel Guerrero" sacamos "201")
            unidad_match = re.search(r'(\d+)', detalle_orig)
            if unidad_match:
                return f"Departamento - {unidad_match.group(1)}"
            return "Departamento"
        
        return detalle_orig

    if 'Detalle' in df_maestra_copy.columns:
        df_maestra_copy['Detalle'] = df_maestra_copy.apply(anonimizar_detalle, axis=1)
    
    df_maestra_copy.columns.name = None  
    
    t_detalle = df_maestra_copy.to_html(
        classes='table table-sm table-striped mb-0', 
        index=False, 
        border=0,
        justify='left'
    )
    
    # --- [CAMBIO 2] TABLA MOROSIDAD MES ACTUAL (PRIVACIDAD NOMINAL) ---
    df_mora_mes_web = df_mora_mes.copy()
    
    # Identificamos cómo se llama la columna de la unidad (puede ser 'unidad' o 'Unidad')
    col_unidad_actual = 'unidad' if 'unidad' in df_mora_mes_web.columns else 'Unidad'
    
    # 1. Creamos la columna 'Tipo de Unidad'
    df_mora_mes_web['Tipo de Unidad'] = "Departamento"
    
    # 2. Definimos el orden asegurando que la unidad NO se pierda
    columnas_finales_mes = ['Tipo de Unidad', col_unidad_actual, 'Monto Actual', 'Estado']
    
    # 3. Filtramos (esto elimina 'residente' automáticamente)
    df_mora_mes_web = df_mora_mes_web[columnas_finales_mes]
    
    # 4. Renombramos la columna de la unidad a "Unidad" para que se vea bien el encabezado
    df_mora_mes_web = df_mora_mes_web.rename(columns={col_unidad_actual: 'Unidad'})

    t_mora_mes = df_mora_mes_web.to_html(
        classes='table table-hover mb-0 text-center', 
        index=False, 
        border=0
    )
    
    # --- [CAMBIO 3] TABLA MOROSIDAD HISTÓRICA (PRIVACIDAD NOMINAL Y LENGUAJE AMIGABLE) ---
    df_mora_hist_web = df_mora_hist.copy()
    
    # Eliminamos Unidad_Num y residente
    columnas_eliminar = ['residente', 'Unidad_Num']
    df_mora_hist_web = df_mora_hist_web.drop(columns=[c for c in columnas_eliminar if c in df_mora_hist_web.columns])
    
    # Insertamos columna Tipo al inicio
    df_mora_hist_web.insert(0, 'Tipo de Unidad', "Departamento")
    
    # --- AJUSTE DE COLUMNA GRAVEDAD A ESTADO ---
    # Cambiamos el nombre de la columna para que sea más neutro
    if 'Gravedad' in df_mora_hist_web.columns:
        # Reemplazamos todos los valores (Riesgo, Crítico, etc.) por "🔴 Pendiente"
        df_mora_hist_web['Gravedad'] = "🔴 Pendiente"
        # Renombramos la columna a "Situación" o "Estado"
        df_mora_hist_web = df_mora_hist_web.rename(columns={'Gravedad': 'Estado'})

    # Aseguramos el orden solicitado (usando el nuevo nombre 'Estado')
    orden_hist = ['Tipo de Unidad', 'Unidad', 'meses deuda', 'detalle meses', 'Deuda Total Est.', 'Estado']
    df_mora_hist_web = df_mora_hist_web[[c for c in orden_hist if c in df_mora_hist_web.columns]]

    # Generamos el HTML (Igualamos clases a la tabla Mes Actual para misma fuente)
    t_mora_hist_html = df_mora_hist_web.to_html(
        classes='table table-hover mb-0 text-center', 
        index=False, 
        border=0
    )
    # TRUCO FINAL: Inyectamos un pequeño ajuste de estilo manual para que el texto no se achique
    t_mora_hist_html = t_mora_hist_html.replace('<table ', '<table style="font-size: 1rem !important; width: 100%;" ')


    # --- LÓGICA DEL VISOR PDF 1: ESTADO DE CUENTA ---
    periodo_limpio = str(mes_actual).replace(" ", "_")
    if pdf_b64:
        nombre_archivo_estado = f"Estado_Cuenta_Creativa_II_{periodo_limpio}.pdf"
        html_pdf_estado = f"""
        <div class="alert alert-primary d-flex justify-content-between align-items-center mb-3">
            <span>📄 <b>Estado de Cuenta Mensual:</b> Resumen detallado de ingresos y egresos.</span>
            <a href="data:application/pdf;base64,{pdf_b64}" download="{nombre_archivo_estado}" class="btn btn-sm btn-primary">Descargar PDF</a>
        </div>
        <iframe src="data:application/pdf;base64,{pdf_b64}" width="100%" height="800px" style="border-radius: 8px; border: 1px solid #ddd;"></iframe>
        """
    else:
        html_pdf_estado = "<p class='text-center p-5 text-muted'>El Estado de Cuenta no está disponible.</p>"

    # --- LÓGICA DEL VISOR PDF 2: INFORME DE GESTIÓN ---
    if pdf_gestion_b64:
        nombre_archivo_gestion = f"Informe_Gestion_Creativa_II_{periodo_limpio}.pdf"
        html_pdf_gestion = f"""
        <div class="alert d-flex justify-content-between align-items-center mb-3 shadow-sm" 
             style="background-color: #2c3e50; color: #ffffff; border-left: 6px solid #f39c12; border-radius: 8px;">
            <span>
                <i class="me-2" style="font-size: 1.2rem;">📄</i> 
                <b style="color: #f39c12;">Informe de Gestión Administrativa:</b> 
                <span style="opacity: 0.9;">Análisis del Desempeño Financiero.</span>
            </span>
            <a href="data:application/pdf;base64,{pdf_gestion_b64}" 
               download="{nombre_archivo_gestion}" 
               class="btn btn-sm fw-bold" 
               style="background-color: #f39c12; color: #2c3e50; border: none; padding: 5px 15px;">
                Descargar PDF
            </a>
        </div>
        <iframe src="data:application/pdf;base64,{pdf_gestion_b64}" width="100%" height="800px" style="border-radius: 8px; border: 1px solid #ddd;"></iframe>
        """
    else:
        html_pdf_gestion = "<p class='text-center p-5 text-muted'>El Informe de Gestión no está disponible.</p>"

    # --- LÓGICA PARA TABLA CONTROL DE PAGOS ---
    if df_control is not None:
        cols_total = list(df_control.columns)
        col_unid = cols_total[0]
        cols_meses = cols_total[1:][-16:] 
        df_c_visual = df_control[[col_unid] + cols_meses].copy()
        
        def limpiar_depto(val):
            try: return str(int(float(val)))
            except: return str(val).strip()

        html_control = '<div class="table-responsive"><table class="table table-sm table-bordered table-control text-center">'
        
        # --- CAMBIO AQUÍ: Encabezado en dos filas y forzamos ancho mínimo ---
        html_control += '<thead><tr>'
        html_control += f'<th class="sticky-col" style="min-width: 80px; line-height: 1.2;">Departamento<br>N°</th>'

        for c in cols_meses: 
            html_control += f'<th style="font-size: 0.8rem; vertical-align: middle;">{c}</th>'
        html_control += '</tr></thead><tbody>'
        
        for _, row in df_c_visual.iterrows():
            depto_limpio = limpiar_depto(row[col_unid])
            html_control += f'<tr><td class="sticky-col fw-bold bg-light">{depto_limpio}</td>'
            for m in cols_meses:
                val = str(row[m]).strip().lower()
                style, display = "", ""
                if val in ["1", "1.0"]:
                    style = 'style="background-color: #d4edda; color: #155724;"'
                    display = "✅"
                elif val in ["0", "0.0", "nan", "none", ""]: display = ""
                else:
                    style = 'style="background-color: #fff3cd;"'
                    display = f"❓ {row[m]}"
                html_control += f'<td {style}>{display}</td>'
            html_control += '</tr>'
        html_control += '</tbody></table></div>'
    else:
        html_control = "<p class='text-center'>No hay datos de control disponibles.</p>"

    # --- CORRECCIÓN DE DATOS Y ALINEACIÓN PARA SALUD ---
    columnas_salud = ['Período', 'Recaudación Comunidad', 'Gastos Reales', 'Diferencia']
    
    if all(c in df_salud_final.columns for c in columnas_salud):
        df_salud_limpio = df_salud_final[columnas_salud].copy()
        
        # Construcción de la tabla
        html_salud = '<div class="table-responsive"><table class="table table-sm table-hover mb-0">'
        html_salud += '<thead><tr class="text-center">'
        for col in columnas_salud:
            html_salud += f'<th>{col}</th>'
        html_salud += '</tr></thead><tbody>'

        for _, row in df_salud_limpio.iterrows():
            html_salud += '<tr>'
            # Columnas centradas
            html_salud += f'<td class="text-center">{row["Período"]}</td>'
            html_salud += f'<td class="text-center">{row["Recaudación Comunidad"]}</td>'
            html_salud += f'<td class="text-center">{row["Gastos Reales"]}</td>'
            
            # Lógica de iconos para Diferencia
            valor_raw = str(row['Diferencia'])
            valor_limpio = valor_raw.replace('.', '').replace('$', '').replace(' ', '').strip()
            
            icon = ""
            try:
                num_diff = float(valor_limpio)
                if num_diff > 0: icon = "🟢 "
                elif num_diff < 0: icon = "🔴 "
            except: pass

            # Columna DIFERENCIA: Justificada a la izquierda, número negro, negrita
            style_diff = 'style="text-align: center; padding-left: 25px; color: #000;"'
            html_salud += f'<td {style_diff}>{icon}{valor_raw}</td>'
            html_salud += '</tr>'
        
        html_salud += '</tbody></table></div>'
        t_salud_mensual = html_salud
    else:
        t_salud_mensual = "<p class='text-center'>No hay datos detallados disponibles.</p>"

    
    config_res = {'responsive': True}
    div_gauge = opy.plot(fig_gauge, auto_open=False, output_type='div', include_plotlyjs='cdn', config=config_res) if fig_gauge else ""
    
    html_mil = "".join([f"<li class='list-group-item d-flex justify-content-between align-items-center'><b>{k}</b> <span class='badge bg-primary rounded-pill'>${v}</span></li>" for k, v in explicacion_mil.items()])

    valor_deuda_h = m_salud.get('total_historico', '0')
    if not str(valor_deuda_h).startswith('$'): valor_deuda_h = f"$ {valor_deuda_h}"

    try:
        suma_mes = pd.to_numeric(df_mora_mes['Monto Actual'].replace('[\$,.]', '', regex=True), errors='coerce').sum()
        valor_mora_mes = f"$ {suma_mes:,.0f}".replace(",", ".")
    except:
        valor_mora_mes = "$ 0"

    html_exclusion = ""
    if lista_final and len(lista_final) > 0:
        u_str = ", ".join([str(x) for x in lista_final])
        sujeto = "La unidad" if len(lista_final) == 1 else "Las unidades"
        verbo = "está excluida" if len(lista_final) == 1 else "están excluidas"
        html_exclusion = f"""
        <div class="alert-info-custom mb-4">
            ℹ️ <strong>Nota de Excepción:</strong> {sujeto} <b>{u_str}</b> {verbo} del cálculo de morosidad por acuerdos administrativos previos.
        </div>
        """

    # --- ESTILOS CONTABLES ---
    try:
        idx_monto_resumen = df_resumen.columns.get_loc("Monto") + 1
        estilo_resumen = f"""
            .resumen-monto td:nth-child({idx_monto_resumen}) {{ 
                text-align: right !important; 
                padding-right: 25px !important; 
                font-family: 'Consolas', 'Monaco', monospace; 
            }}
            .resumen-monto th:nth-child({idx_monto_resumen}) {{ 
                text-align: right !important; 
                padding-right: 25px !important; 
                font-weight: bold !important;
            }}"""
    except: estilo_resumen = ""

    try:
        idx_monto_detalle = df_maestra_copy.columns.get_loc("Monto") + 1
        estilo_detalle = f"""
            .detalle-monto td:nth-child({idx_monto_detalle}) {{ 
                text-align: right !important; 
                padding-right: 25px !important; 
                font-family: 'Consolas', 'Monaco', monospace;
            }}
            .detalle-monto th:nth-child({idx_monto_detalle}) {{ 
                text-align: right !important; 
                padding-right: 25px !important; 
                font-weight: bold !important;
            }}"""
    except: estilo_detalle = ""

# --- 1. AQUÍ PEGAS LA LÓGICA DE ACTIVACIÓN (ANTES DEL TEMPLATE) ---
    estados = {
        "bal": ver_balance,
        "res": ver_resumen_fin,
        "mor": ver_morosidad,
        "con": ver_control,
        "pdf": ver_pdfs
    }

    primera_activa = None
    for clave, valor in estados.items():
        if valor:
            primera_activa = clave
            break

    c_bal = "active" if primera_activa == "bal" else ""
    p_bal = "show active" if primera_activa == "bal" else ""
    c_res = "active" if primera_activa == "res" else ""
    p_res = "show active" if primera_activa == "res" else ""
    c_mor = "active" if primera_activa == "mor" else ""
    p_mor = "show active" if primera_activa == "mor" else ""
    c_con = "active" if primera_activa == "con" else ""
    p_con = "show active" if primera_activa == "con" else ""
    c_pdf = "active" if primera_activa == "pdf" else ""
    p_pdf = "show active" if primera_activa == "pdf" else ""

    # Variable auxiliar para el JS de control de pagos
    mes_seleccionado = mes_actual

    html_template = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        {script_analitica}
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ background-color: #f4f1ea; font-family: 'Segoe UI', sans-serif; }}
            .navbar {{ background: linear-gradient(90deg, #5d4037 0%, #8d6e63 100%); color: white; border-bottom: 5px solid #a8c69f; padding: 20px 0; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }}
            .navbar-brand-custom {{ font-size: 1.75rem; font-weight: 800; letter-spacing: -0.5px; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); }}
            .badge-periodo {{ 
                font-size: 1rem; padding: 8px 18px; 
                background-color: rgba(255,255,255,0.15); 
                border: 1px solid rgba(255,255,255,0.3); 
                backdrop-filter: blur(5px);
                display: flex; flex-direction: column; align-items: flex-end; line-height: 1.2;
            }}
            .fecha-sync {{ font-size: 0.65rem; font-weight: normal; opacity: 0.8; margin-top: 2px; }}
            .card-metric {{ background: white; border-top: 4px solid #8d6e63; border-radius: 8px; padding: 15px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); min-height: 90px; }}
            .table-container {{ background: white; border-radius: 8px; padding: 15px; overflow-x: auto; margin-bottom: 20px; }}
            .text-center th {{ text-align: center !important; }}
            .text-center td {{ text-align: center !important; vertical-align: middle; }}
            {estilo_resumen}
            {estilo_detalle}
            .resumen-monto th, .detalle-monto th {{ text-align: center; }}
            .resumen-monto td:not(:nth-child({df_resumen.columns.get_loc("Monto")+1 if "Monto" in df_resumen.columns else 0})), 
            .detalle-monto td:not(:nth-child({df_maestra_copy.columns.get_loc("Monto")+1 if "Monto" in df_maestra_copy.columns else 0})) 
            {{ text-align: center; }}
            .section-title {{ color: #8d6e63; font-weight: bold; margin-top: 25px; border-left: 5px solid #a8c69f; padding-left: 10px; margin-bottom: 15px; }}
            .nota-salud {{ padding: 20px; border-radius: 8px; margin-bottom: 20px; color: white; text-align: center; }}
            .alert-info-custom {{ background-color: #d1ecf1; border-left: 6px solid #0c5460; color: #0c5460; padding: 15px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            .progress {{ height: 20px; background-color: #e9ecef; border-radius: 10px; margin-top: 8px; }}
            .progress-bar {{ background-color: #a8c69f; color: #2c3e50; font-weight: bold; }}
            .table-control thead th {{ background-color: #2F5597; color: white; white-space: nowrap; }}
            .table-control td {{ text-align: center; vertical-align: middle; white-space: nowrap; }}
            .sticky-col {{ position: sticky; left: 0; z-index: 2; border-right: 2px solid #2F5597 !important; }}
            .nav-pills .nav-link {{ color: #8d6e63; border: 1px solid #8d6e63; margin-right: 5px; }}
            .nav-pills .nav-link.active {{ background-color: #8d6e63; color: white; border: 1px solid #8d6e63; }}
            footer {{ margin-top: 50px; padding: 25px 0; border-top: 1px solid #ddd; color: #8d6e63; font-size: 0.85rem; text-align: center; }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark mb-4">
            <div class="container d-flex justify-content-between align-items-center">
                <span class="navbar-brand-custom">🏢 Gestión Financiera Edificio Creativa II</span>
                <div class="badge rounded-pill badge-periodo">
                    <span>Período: {mes_actual}</span>
                    <span class="fecha-sync">Actualizado el: {fecha_actualizacion}</span>
                </div>
            </div>
        </nav>

        <div class="container">
            <ul class="nav nav-tabs mb-4" id="portalTabs" role="tablist">
                {f'<li class="nav-item"><button class="nav-link {c_bal}" data-bs-toggle="tab" data-bs-target="#balance">📊 Balance</button></li>' if ver_balance else ""}
                {f'<li class="nav-item"><button class="nav-link {c_res}" data-bs-toggle="tab" data-bs-target="#salud">📈 Resumen Financiero</button></li>' if ver_resumen_fin else ""}
                {f'<li class="nav-item"><button class="nav-link {c_mor}" data-bs-toggle="tab" data-bs-target="#morosidad">🚦 Morosidad</button></li>' if ver_morosidad else ""}
                {f'<li class="nav-item"><button class="nav-link {c_con}" data-bs-toggle="tab" data-bs-target="#control">🗓️ Control de Pagos</button></li>' if ver_control else ""}
                {f'<li class="nav-item"><button class="nav-link {c_pdf}" data-bs-toggle="tab" data-bs-target="#tab-pdf">📄 Informes PDFs</button></li>' if ver_pdfs else ""}
            </ul>

            <div class="tab-content">
                {f"""<div class="tab-pane fade {p_bal}" id="balance">
                    <div class="row g-3 mb-4">
                        <div class="col-md-3"><div class="card-metric"><small class="fw-bold">Saldo Inicial</small><div class="h5">{metricas.get('saldo_i', '$ 0')}</div></div></div>
                        <div class="col-md-3"><div class="card-metric"><small class="fw-bold">Ingresos (+)</small><div class="h5 text-success">{metricas.get('ingresos', '$ 0')}</div></div></div>
                        <div class="col-md-3"><div class="card-metric"><small class="fw-bold">Egresos (-)</small><div class="h5 text-danger">{metricas.get('egresos', '$ 0')}</div></div></div>
                        <div class="col-md-3"><div class="card-metric"><small class="fw-bold">Saldo Final</small><div class="h5">{metricas.get('saldo_f', '$ 0')}</div></div></div>
                    </div>
                    <div class="row mb-4">
                        <div class="col-md-7">
                            <h5 class="section-title">⚖️ Gestión de Caja y Solvencia</h5>
                            <div style="background-color: {metricas.get('bg_gestion', '#eee')}; color: {metricas.get('color_gestion', '#333')}; padding:15px; border-radius:8px; border:1px solid;">
                                <strong>Nota:</strong> {metricas.get('nota_gestion', 'Información no disponible')}
                            </div>
                            <div class="mt-3 p-2">
                                <strong>Disponibilidad Total de Fondos: {metricas.get('supervivencia', '0%')}</strong>
                                <div class="progress"><div class="progress-bar" style="width: {metricas.get('supervivencia_num', 0)}%"></div></div>
                                <p class="text-muted mb-1" style="font-size: 0.85em; font-style: italic;">
                                    * Porcentaje de dinero remanente sobre el total de recursos que circularon en el mes.
                                </p>
                            </div>
                        </div>
                        <div class="col-md-5">{div_gauge}</div>
                    </div>
                    <h5 class="section-title">🏷️ Resumen por Categoría</h5>
                    <div class="table-container resumen-monto">{t_resumen}</div>
                    <p style="color: #6c757d; font-size: 0.88rem; margin-bottom: 12px;">
                       Vista consolidada de los movimientos financieros clasificados por flujo de caja y categoría.
                    </p>
                    <hr style="border: none; border-top: 1px solid #dee2e6; margin: 2rem 0;">
                    <h5 class="section-title">🔎 Detalle de Movimientos </h5>
                    <div class="table-container detalle-monto" style="font-size: 0.85rem;">{t_detalle}</div>
                    <p style="color: #6c757d; font-size: 0.88rem; margin-bottom: 12px;">
                       Registro completo de ingresos y egresos correspondientes al período seleccionado.
                    </p>
                </div>""" if ver_balance else ""}

                {f"""<div class="tab-pane fade {p_res}" id="salud">
                    <div class="nota-salud" style="background-color: {salud_data['color_salud']};">
                        <h3 class="mb-0">🛡️ Diagnóstico del Año: {salud_data['avg_sostenibilidad']} de Sostenibilidad</h3>
                        <p class="mb-0">{salud_data['status_msg']}</p>
                    </div>
                    <div class="mt-4 mb-2">
                        <h5 class="section-title">🌊 Análisis de Flujo y Sostenibilidad</h5>
                        <p class="text-secondary mb-0" style="font-size: 0.95em;">
                            Comparativa de los últimos 12 meses: <strong>Recaudación Comunidad vs. Gastos Operacionales</strong>.
                        </p>
                    </div>
                    <div class="mt-3 mb-1">
                        <h6 style="color: #475569; font-size: 0.95em; font-weight: 600; ">
                            🗓️ Evolución de Ingresos y Egresos Operacionales
                        </h6>
                    </div>
                    <div class="table-container">{div_tendencia}</div>
                    <div class="row">
                        <div class="col-md-7">
                            <h5 class="section-title">🍰 ¿Cómo se distribuye el gasto?</h5>
                            <div class="table-container">{div_torta}</div>
                        </div>
                        <div class="col-md-5">
                            <h5 class="section-title">💰 Por cada $1.000 que pagas:</h5>
                            <ul class="list-group shadow-sm">{html_mil}</ul>
                        </div>
                    </div>
                    <h5 class="section-title">📋 Desglose Mensual: Recaudación vs. Gastos</h5>
                    <div class="table-container">{t_salud_mensual}</div>
                    <p style="color: #6B7280; font-size: 0.85rem; margin-top: 12px; line-height: 1.4; font-style: italic;">
                        <strong>Nota:</strong> Este cuadro compara el total de ingresos recibidos frente a los egresos ejecutados en cada período. 
                        La diferencia positiva o negativa representa el excedente o déficit operativo mensual de la comunidad. 
                        (No considera saldos de arrastre de meses anteriores).
                    </p>

                </div>""" if ver_resumen_fin else ""}

                {f"""<div class="tab-pane fade {p_mor}" id="morosidad">
                    <div class="row g-3 mb-4">
                        <div class="col-md-6"><div class="card-metric" style="border-top-color: #0d6efd;"><small class="fw-bold">Total Pendiente del Mes ({mes_actual})</small><div class="h4 text-primary">{valor_mora_mes}</div></div></div>
                        <div class="col-md-6"><div class="card-metric" style="border-top-color: #dc3545;"><small class="fw-bold">Total Deuda Histórica Acumulada</small><div class="h4 text-danger">{valor_deuda_h}</div></div></div>
                    </div>
                    {html_exclusion}
                    <h5 class="section-title">📅 Pendientes del Mes Actual</h5>
                    <div class="table-container text-center">{t_mora_mes}</div>

                    <hr class="my-5" style="opacity: 0.1;"> 
                    <div class="alert alert-light border-0 text-muted sm-small mb-4" style="font-size: 0.85rem; background-color: #f8f9fa;">
                        🔍 <b>Nota sobre Deuda Histórica:</b> Los montos a continuación corresponden a saldos que no fueron cubiertos en meses anteriores y se mantienen pendientes en la contabilidad de la comunidad.
                    </div>

                    <h5 class="section-title">🕵️ Análisis de Deuda Histórica Acumulada</h5>
                    <div class="table-container text-center" style="font-size: 0.85rem;">{t_mora_hist_html}</div>
                </div>""" if ver_morosidad else ""}

                {f"""<div class="tab-pane fade {p_con}" id="control">
                    <h5 class="section-title">🗓️ Matriz de Pagos por Departamento (Vista 16 meses)</h5>
                    <div class="alert alert-warning py-2 shadow-sm" style="font-size: 0.95rem; border-left: 4px solid #ffc107;">
                        <i class="fas fa-exclamation-circle mr-1"></i>
                        <b>Nota:</b> Este historial refleja todos los abonos registrados a la fecha de emisión, 
                        incluyendo pagos anticipados posteriores al periodo de <b>{mes_seleccionado}</b>.
                    </div>
                    <div style="font-size: 0.95rem;">
                    {html_exclusion}
                    </div>                    
                    <div class="table-container">{html_control}</div>
                    <div class="alert alert-light border shadow-sm"><small><b>Leyenda:</b> ✅ Pago registrado | ❓ Dato requiere revisión | Celda vacía: Pendiente.</small></div>
                </div>""" if ver_control else ""}

                {f"""<div class="tab-pane fade {p_pdf}" id="tab-pdf">
                    <div class="alert alert-light border shadow-sm d-flex align-items-center py-2 px-3 mb-4">
                        <span class="me-2">📅</span>
                        <span style="color: #5d4037;">Documentación Oficial generada para el periodo <b>{mes_actual}</b></span>
                    </div>
                    <h5 class="section-title">📄 Documentos del Periodo</h5>
                    <ul class="nav nav-pills mb-3" id="pills-tab" role="tablist">
                        <li class="nav-item" role="presentation">
                            <button class="nav-link active" data-bs-toggle="pill" data-bs-target="#pdf1" type="button">1. Estado de Cuenta</button>
                        </li>
                        <li class="nav-item" role="presentation">
                            <button class="nav-link" data-bs-toggle="pill" data-bs-target="#pdf2" type="button">2. Informe de Gestión</button>
                        </li>
                    </ul>
                    <div class="tab-content" id="pills-tabContent">
                        <div class="tab-pane fade show active" id="pdf1">{html_pdf_estado}</div>
                        <div class="tab-pane fade" id="pdf2">{html_pdf_gestion}</div>
                    </div>
                </div>""" if ver_pdfs else ""}
            </div>
        </div>

        <footer>
            <div class="container">
                <p class="mb-1 fw-bold">🏢 Edificio Creativa II - Gestión Financiera</p>
                <p class="mb-0"> Información actualizada el {fecha_actualizacion} </p>
            </div>
        </footer>

        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script>
            // Forzar resize de gráficos al cambiar pestañas
            document.querySelectorAll('button[data-bs-toggle="tab"]').forEach(tab => {{
                tab.addEventListener('shown.bs.tab', () => {{
                    window.dispatchEvent(new Event('resize'));
                }});
            }});
        </script>
    </body>
    </html>
    """

# --- FIN DE LA GENERACIÓN DEL REPORTE EXTERNO HTML ---
    
    # Usamos tu GPS para guardar el archivo donde está el ejecutable
    ruta_final = obtener_ruta_recurso("index.html")
    with open(ruta_final, "w", encoding="utf-8") as f:
        f.write(html_template)

#--- FIN DE LA FUNCIÓN PARA HTML ---


# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="Gestión Creativa II", layout="wide")

# --- BLOQUE DE TÍTULO ---
st.markdown("""
    <style>
    .banner-contenedor {
        background-image: 
            linear-gradient(rgba(0, 0, 0, 0.75), rgba(0, 0, 0, 0.75)),
            url('https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=2000&auto=format&fit=crop');
        background-size: cover; background-position: center; padding: 40px 0px;
        border-radius: 15px; margin-bottom: 25px; text-align: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    }
    .banner-titulo {
        color: #FFFFFF !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 2.5rem; font-weight: bold; margin: 0;
        text-shadow: 2px 2px 10px rgba(0,0,0,0.9);
    }
    .banner-subtitulo {
        color: #f0f0f0 !important; font-size: 1.2rem; margin: 5px 0 0 0;
        text-shadow: 1px 1px 5px rgba(0,0,0,0.8);
    }
    /* Estilo para encabezados (Headers) */
    [data-testid="stHeaderRowCell"] {
        text-align: center !important;
        font-weight: bold !important;
    }
    </style>
    <div class="banner-contenedor">
        <h1 class="banner-titulo">🏢 Sistema de Gestión Financiera</h1>
        <p class="banner-subtitulo">Edificio Creativa II - Panel de Administración</p>
    </div>
    """, unsafe_allow_html=True)

@st.cache_data
def load_data():
    try:
        # --- CAMBIO APLICADO PARA PASO 2 ---
        file_path = obtener_ruta_recurso("Gestión ECreativa II v1.0.xlsx")
        # -----------------------------------
        
        xls = pd.ExcelFile(file_path)
        df_ing = pd.read_excel(xls, sheet_name=0)
        df_egr = pd.read_excel(xls, sheet_name=1)
        
        df_control = None
        if 'ControlPagoxDeptos' in xls.sheet_names:
            df_control = pd.read_excel(xls, sheet_name='ControlPagoxDeptos')
            nuevas_cols = []
            for col in df_control.columns:
                if isinstance(col, (pd.Timestamp, datetime)):
                    nuevas_cols.append(col.strftime('%m/%Y'))
                else:
                    nuevas_cols.append(str(col))
            df_control.columns = nuevas_cols
        
        for col in COLS_FINANCIERAS:
            if col in df_ing.columns:
                df_ing[col] = pd.to_numeric(df_ing[col], errors='coerce').fillna(0)
            
        df_egr['Gasto Mensual'] = pd.to_numeric(df_egr['Gasto Mensual'], errors='coerce').fillna(0)
        df_ing['Unidad'] = df_ing['Unidad'].astype(str).str.strip()
        df_ing['Tipo de Unidad'] = df_ing['Tipo de Unidad'].astype(str).str.strip()
        
        return df_ing, df_egr, df_control
    except Exception as e:
        st.error(f"Error al cargar el archivo Excel: {e}")
        return None, None, None

def formato_clp(x):
    try:
        # 1. Si es un valor nulo (NaN, None), devolvemos "0"
        if pd.isna(x) or x is None:
            return "0"
        
        # 2. Convertimos a float primero (por si viene como "1500.0") 
        # y luego a int para quitar decimales.
        valor_numerico = int(float(x))
        
        # 3. Formateamos con puntos
        return "{:,}".format(valor_numerico).replace(",", ".")
    except (ValueError, TypeError):
        # 4. Si llega cualquier otra cosa que no sea número, devolvemos "0"
        return "0"

df_ing, df_egr, df_control = load_data()

if df_ing is not None:

    # --- PANEL DE CONTROL ADMINISTRATIVO EN SIDEBAR ---
    # Este bloque reemplaza totalmente al anterior para evitar el error de variable no definida.
    with st.sidebar:
        import time
        from datetime import datetime
        import base64
        import requests

        st.title("⚙️ Administración")

        # Lógica de limpieza previa
        if st.session_state.get('reset_subida', False):
            st.session_state['autorizar_subida'] = False
            st.session_state['reset_subida'] = False

        # --- USO DE TABS PARA AHORRAR ESPACIO VERTICAL ---
        # Añadida la tercera pestaña "🚨 Estados" para manejo de emergencias/espera
        tab_datos, tab_server, tab_emergencia = st.tabs(["📊 Datos", "🌐 Servidor", "🚨 Estados"])

        with tab_datos:
            # 1. Sincronización de Datos (Excel -> App)
            if st.button("🔄 Sincronizar con Excel", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
            
            st.divider()
            
            portal_listo = st.session_state.get('sincro_exitosa', False)
            
            if portal_listo:
                # Reemplazamos espacio_servidor por componentes nativos directos
                st.success("✅ Portal Sincronizado")
                fecha_fija = st.session_state.get('fecha_temporal_sincro', '')
                if fecha_fija:
                    st.caption(f"🕒 Actualizado: {fecha_fija}")
                
                if st.button("🏗️ Actualizar Portal Local", use_container_width=True, help="Fuerza la creación del HTML con los filtros actuales"):
                    st.session_state['sincro_exitosa'] = False
                    st.session_state.pop('fecha_temporal_sincro', None)
                    st.rerun()
            else:
                # --- MEJORA: Solo ejecutamos la animación y cambiamos el estado ---
                st.write("⏳ Generando base del portal...")
                progreso = st.progress(0)
                for i in range(100):
                    import time
                    time.sleep(0.01) 
                    progreso.progress(i + 1)
                
                # ¡PASO CLAVE!: Marcamos el éxito en el state y guardamos la hora
                from datetime import datetime
                st.session_state['sincro_exitosa'] = True
                st.session_state['fecha_temporal_sincro'] = datetime.now().strftime("%H:%M:%S")

                # Forzamos el reinicio inmediato para que aparezca el mensaje de éxito (bloque IF superior)
                st.rerun()

        with tab_server:
            # --- SECCIÓN DE PUBLICACIÓN (PASO FINAL) ---
            st.markdown("### Publicar en Web")
            autorizar = st.checkbox("🔓 Autorizar subida manual", value=False, key="autorizar_subida")
            
            boton_deshabilitado = not (portal_listo and autorizar)

            # 1. Botón principal (Abre el diálogo)
            if st.button("📤 Subir Portal al Servidor", 
                         type="primary", 
                         use_container_width=True, 
                         disabled=boton_deshabilitado,
                         key="btn_subida_sidebar"):
                st.session_state['esperando_confirmacion'] = True

        with tab_emergencia:
            # --- SECCIÓN DE ESTADOS DE ESPERA / EMERGENCIA ---
            st.markdown("### Interruptores de Estado")
            st.info("Activa pantallas temporales que reemplazan el portal financiero.")

            # --- BOTÓN ESPECIAL: MODO CONSTRUCCIÓN ---
            if st.button("🚧 Poner en Construcción", help="Sube una página de próximamente al servidor", use_container_width=True):
                try:
                    with st.status("🛠️ Subiendo modo construcción...", expanded=True) as status_c:
                        html_mantenimiento = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Portal En Construcción</title><style>body { margin: 0; padding: 0; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: white; overflow: hidden; font-family: sans-serif; }.container { text-align: center; }img { max-width: 90%; height: auto; display: block; margin: 0 auto; }p { color: #555; margin-top: 20px; font-size: 1.2em; }</style></head><body><div class="container"><img src="construccion.png" alt="Portal en construcción"><p>🏗️ Próximamente disponible para todos los copropietarios.</p></div></body></html>"""
                        
                        TOKEN_GITHUB = "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"
                        REPO_GITHUB = "Edificio-Creativa-II/portal"
                        URL_API = f"https://api.github.com/repos/{REPO_GITHUB}/contents/index.html"
                        headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github.v3+json"}
                        
                        res = requests.get(URL_API, headers=headers, timeout=10)
                        sha = res.json().get('sha') if res.status_code == 200 else None
                        
                        encoded = base64.b64encode(html_mantenimiento.encode('utf-8')).decode('utf-8')
                        payload = {"message": "🚧 Modo Construcción", "content": encoded, "branch": "main"}
                        if sha: payload["sha"] = sha
                        
                        subida = requests.put(URL_API, headers=headers, json=payload, timeout=15)
                        if subida.status_code in [200, 201]:
                            status_c.update(label="✅ Modo Construcción Activo", state="complete")
                            time.sleep(1.5)
                            st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

            st.divider()

            # --- BOTÓN ESPECIAL: MODO MANTENCIÓN ---
            if st.button("🛠️ Poner en Mantención", help="Sube una página de mantención financiera al servidor", use_container_width=True):
                try:
                    with st.status("🛠️ Subiendo modo mantención...", expanded=True) as status_m:
                        html_mante_final = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Portal en Mantenimiento</title><style>body { margin: 0; padding: 0; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #f7f7f7; overflow: hidden; font-family: sans-serif; }.container { text-align: center; }img { max-width: 90%; height: auto; display: block; margin: 0 auto; box-shadow: 0 10px 30px rgba(0,0,0,0.08); border-radius: 12px; }p { color: #5d4037; margin-top: 30px; font-size: 1.1em; font-weight: 500; }</style></head><body><div class="container"><img src="mantenimiento.png" alt="Portal en Mantenimiento"><p>🛠️ Próximamente volverá a estar disponible para todos los copropietarios.</p></div></body></html>"""
                        
                        TOKEN_GITHUB = "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"
                        REPO_GITHUB = "Edificio-Creativa-II/portal"
                        URL_API = f"https://api.github.com/repos/{REPO_GITHUB}/contents/index.html"
                        headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github.v3+json"}
                        
                        res = requests.get(URL_API, headers=headers, timeout=10)
                        sha = res.json().get('sha') if res.status_code == 200 else None
                        
                        encoded = base64.b64encode(html_mante_final.encode('utf-8')).decode('utf-8')
                        payload = {"message": "🛠️ Modo Mantención", "content": encoded, "branch": "main"}
                        if sha: payload["sha"] = sha
                        
                        subida = requests.put(URL_API, headers=headers, json=payload, timeout=15)
                        if subida.status_code in [200, 201]:
                            status_m.update(label="✅ Modo Mantención Activo", state="complete")
                            time.sleep(1.5)
                            st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

        # --- DIÁLOGOS DE CONFIRMACIÓN (Fuera de las pestañas para asegurar visibilidad) ---
        if st.session_state.get('esperando_confirmacion', False):
            with st.sidebar.container(border=True):
                st.warning("⚠️ **¿Publicar cambios?**")
                col_si, col_no = st.columns(2)
                
                with col_no:
                    if st.button("❌ No", use_container_width=True):
                        st.session_state['esperando_confirmacion'] = False
                        st.rerun()
                
                with col_si:
                    if st.button("✅ Sí", type="primary", use_container_width=True):
                        st.session_state['ejecutar_subida'] = True
                        st.session_state['esperando_confirmacion'] = False
                        st.rerun()

        # 3. EJECUCIÓN DE LA SUBIDA
        if st.session_state.get('ejecutar_subida', False):
            try:
                with st.status("🚀 Publicando en servidor...", expanded=True) as status:
                    TOKEN_GITHUB = "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" 
                    REPO_GITHUB = "Edificio-Creativa-II/portal"
                    URL_API = f"https://api.github.com/repos/{REPO_GITHUB}/contents/index.html"
                    headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github.v3+json"}
                    
                    with open("index.html", "r", encoding="utf-8") as f:
                        html_content = f.read()

                    # --- MEJORA: INYECCIÓN DE TÍTULO SI NO EXISTE ---
                    titulo_web = "<title>Portal Financiero - Edificio Creativa II</title>"
                    if "<title>" not in html_content:
                        html_content = html_content.replace("<head>", f"<head>{titulo_web}")
                    # -----------------------------------------------

                    res = requests.get(URL_API, headers=headers, timeout=10)
                    sha = res.json().get('sha') if res.status_code == 200 else None

                    encoded = base64.b64encode(html_content.encode('utf-8')).decode('utf-8')
                    payload = {"message": f"Web Update {datetime.now().strftime('%H:%M')}", "content": encoded, "branch": "main"}
                    if sha: payload["sha"] = sha

                    subida = requests.put(URL_API, headers=headers, json=payload, timeout=15)
                    
                    if subida.status_code in [200, 201]:
                        status.update(label="✅ ¡Publicado con éxito!", state="complete")
                        st.session_state['reset_subida'] = True
                        st.session_state['ejecutar_subida'] = False
                        st.balloons()
                        time.sleep(1.5)
                        st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.session_state['ejecutar_subida'] = False

        st.divider()
    # --- FIN PANEL DE CONTROL ADMINISTRATIVO EN SIDEBAR ---

        
        # --- SECCIÓN DE CONSULTAS (FILTROS) ---
        st.header("🔍 Consultas")
        busqueda_unidad = st.sidebar.text_input("Buscador Historial (Ej: 201)", "").strip()
        
        # 1. Obtenemos los meses únicos
        meses_sucios = df_ing['Período'].unique()
        
        # 2. Ordenamos de más reciente a más antiguo
        meses_disponibles = sorted(
            meses_sucios, 
            key=lambda x: pd.to_datetime(x, format='%m/%Y'), 
            reverse=True
        )
        
        # 3. Selector de Período
        mes_seleccionado = st.sidebar.selectbox("Seleccione el Período", meses_disponibles)

        # 4. Aplicar la regla de negocio inmediatamente
        UNIDADES_EXCLUIDAS = calcular_excluidos_del_mes(mes_seleccionado, DICCIONARIO_EXCLUSIONES)


    # --- MÓDULO: HISTORIAL DE PAGOS POR UNIDAD ---
    if busqueda_unidad:
        # 1. Filtramos los datos de la unidad seleccionada
        df_historial = df_ing[df_ing['Unidad'].astype(str).str.contains(busqueda_unidad, case=False)].copy()
        
        if not df_historial.empty:
            st.markdown(f"### 📈 Historial de Pagos - Unidad {busqueda_unidad}")
            
            col_h1, col_h2 = st.columns([1, 2])
            
            with col_h1:
                # Métrica de total histórico
                total_h = df_historial['Pago Total del Mes'].sum()
                st.metric("Total Pagado Histórico", f"$ {formato_clp(total_h)}")
                
                # PREPARACIÓN DE TABLA PARA VISUALIZACIÓN
                df_h_display = df_historial[['Período', 'Tipo de Ingreso', 'Pago Total del Mes']].copy()
                
                # A) Formateo de Monto (Texto para alinear a la izquierda)
                df_h_display['Pago Total del Mes'] = df_h_display['Pago Total del Mes'].apply(lambda x: f"$ {formato_clp(x)}")
                
                # B) Formateo de Período (Texto para alinear a la izquierda)
                df_h_display['Período'] = df_h_display['Período'].astype(str)
                
                # C) Renderizado con anchos controlados
                st.dataframe(
                    df_h_display,
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "Período": st.column_config.Column(width="small"),
                        "Tipo de Ingreso": st.column_config.Column(width="medium"),
                        "Pago Total del Mes": st.column_config.Column("Monto Pagado", width="medium")
                    }
                )

            # --- COLUMNA 2: GRÁFICO DE TENDENCIA (LÍNEA FLOTANTE) ---
            with col_h2:
                with st.container(border=True):
                    # Título manual para evitar errores de Plotly
                    st.markdown("#### 📈 Tendencia de Pagos Mensuales")
                    
                    # CORRECCIÓN DE ORDEN: Convertimos temporalmente a datetime para ordenar cronológicamente
                    df_plot = df_historial.copy()
                    df_plot['Fecha_Temp'] = pd.to_datetime(df_plot['Período'], format='%m/%Y')
                    df_plot = df_plot.sort_values('Fecha_Temp').copy()
                    
                    # ASEGURAR DATOS LIMPIOS: Convertimos Período a string y Monto a float
                    df_plot['Período'] = df_plot['Período'].astype(str)
                    df_plot['Monto_Num'] = df_plot['Pago Total del Mes'].astype(float)
                    
                    # Creamos el texto formateado en una columna aparte
                    df_plot['Texto_Etiqueta'] = df_plot['Monto_Num'].apply(lambda x: f"$ {formato_clp(x)}")
                    
                    max_pago = df_plot['Monto_Num'].max() if not df_plot.empty else 100
                    limite_superior = max_pago * 1.3  # Un poco más de aire para las etiquetas
                    
                    # Creamos el gráfico indicando explícitamente la columna de texto
                    fig_h = px.line(
                        df_plot, 
                        x='Período', 
                        y='Monto_Num',
                        markers=True,
                        text='Texto_Etiqueta' # Referencia directa a la columna
                    )
                    
                    # Estilo de la línea y puntos
                    fig_h.update_traces(
                        line_color="#1E3A8A", 
                        line_width=3,
                        marker=dict(size=10, color="#1E3A8A", line=dict(width=2, color="white")),
                        textposition="top center",
                        cliponaxis=False # Evita que el texto se corte en los bordes
                    )
                    
                    fig_h.update_layout(
                        xaxis_title=None,
                        yaxis_title=None,
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        margin=dict(l=10, r=10, t=30, b=10), # Margen superior ajustado
                        yaxis=dict(
                            range=[0, limite_superior],
                            tickprefix="$ ",
                            showgrid=True,
                            gridcolor='#F3F4F6'
                        ),
                        xaxis=dict(type='category', showgrid=False) # Forzamos tipo categoría
                    )
                    
                    st.plotly_chart(fig_h, use_container_width=True, config={'displayModeBar': False})

            st.markdown("---")

        else:
            st.info(f"💡 No hay registros para la unidad '{busqueda_unidad}'.")

# 1. CÁLCULOS DE SALDOS
    df_ing_mes = df_ing[df_ing['Período'] == mes_seleccionado].copy()
    df_egr_mes = df_egr[df_egr['Período'] == mes_seleccionado].copy()

    saldo_inicial = df_ing_mes[df_ing_mes['Tipo de Ingreso'] == 'Saldo Mes Anterior']['Pago Total del Mes'].sum()
    ingresos_puros = df_ing_mes[df_ing_mes['Tipo de Ingreso'] != 'Saldo Mes Anterior']['Pago Total del Mes'].sum()
    gastos_puros = df_egr_mes['Gasto Mensual'].sum()
    saldo_final = (saldo_inicial + ingresos_puros) - gastos_puros

    # 2. PREPARACIÓN DE DATOS PARA TABLAS
    ing_to_concat = df_ing_mes.copy()
    ing_to_concat['Flujo'], ing_to_concat['Orden'] = "(+) INGRESO", 1
    ing_to_concat['Categoría'], ing_to_concat['Sub-Categoría'] = ing_to_concat['Tipo de Ingreso'], ing_to_concat['Tipo de Unidad']
    
    # --- CAMBIO PARA PRIVACIDAD NOMINAL EN DETALLE (PDF Y EXCEL) ---
    def formatear_detalle_ingreso(row):
        # Si la categoría es Gastos Comunes o el flujo es ingreso de departamento
        if str(row['Categoría']) == "Gastos Comunes" or str(row['Sub-Categoría']) == "Departamento":
            return f"Departamento - {row['Unidad']}"
        return f"{row['Unidad']} - {row['Residente']}"

    ing_to_concat['Detalle'] = ing_to_concat.apply(formatear_detalle_ingreso, axis=1)
    ing_to_concat['Monto'] = ing_to_concat['Pago Total del Mes']
    
    egr_to_concat = df_egr_mes.copy()
    egr_to_concat['Flujo'], egr_to_concat['Orden'] = "(-) EGRESO", 2
    egr_to_concat['Categoría'], egr_to_concat['Sub-Categoría'] = egr_to_concat['Tipo de Egreso'], egr_to_concat['Clasificación']
    egr_to_concat['Detalle'], egr_to_concat['Monto'] = egr_to_concat['Descripción'].astype(str), egr_to_concat['Gasto Mensual']

    cols_m = ['Orden', 'Flujo', 'Categoría', 'Sub-Categoría', 'Detalle', 'Monto']
    tabla_maestra = pd.concat([ing_to_concat[cols_m], egr_to_concat[cols_m]], ignore_index=True).sort_values('Orden')
    tabla_resumen = tabla_maestra.groupby(['Orden', 'Flujo', 'Categoría', 'Sub-Categoría'])['Monto'].sum().reset_index().sort_values('Orden')

    # 3. DEFINICIÓN DE FUNCIONES (Primero Excel, luego PDF)
    def generate_excel():
        output = io.BytesIO()
        sheet_clean = str(mes_seleccionado).replace("/", "-")
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_clean)
            worksheet.hide_gridlines(2)
            worksheet.set_paper(1)
            worksheet.set_portrait()
            worksheet.fit_to_pages(1, 1)
            worksheet.center_horizontally()
            worksheet.set_margins(0.5, 0.5, 0.5, 0.5)
            title_fmt = workbook.add_format({'bold': True, 'size': 16, 'font_name': 'Arial', 'align': 'center', 'font_color': '#2F5597'})
            summary_label_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'size': 11, 'bg_color': '#F2F2F2', 'border': 2, 'align': 'left'})
            money_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': 'Arial', 'size': 11, 'align': 'right', 'border': 2})
            saldo_final_lbl_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'size': 11, 'bg_color': '#D9EAD3', 'border': 2, 'align': 'left'})
            saldo_final_val_fmt = workbook.add_format({'bold': True, 'num_format': '#,##0', 'font_name': 'Arial', 'size': 11, 'align': 'right', 'bg_color': '#D9EAD3', 'border': 2})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            ingreso_text_fmt = workbook.add_format({'border': 1, 'font_name': 'Arial', 'font_color': '#002060'})
            ingreso_money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'font_name': 'Arial', 'align': 'right', 'font_color': '#002060'})
            egreso_text_fmt = workbook.add_format({'border': 1, 'font_name': 'Arial', 'font_color': '#C00000'})
            egreso_money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'font_name': 'Arial', 'align': 'right', 'font_color': '#C00000'})
            worksheet.merge_range('A1:E1', f'ESTADO DE CUENTA - PERÍODO {mes_seleccionado}', title_fmt)
            resumen_data = [('Saldo Inicial (Anterior)', saldo_inicial), ('Total Ingresos Mes', ingresos_puros), ('Total Egresos Mes', gastos_puros)]
            for i, (label, value) in enumerate(resumen_data):
                worksheet.merge_range(i+2, 0, i+2, 1, label, summary_label_fmt)
                worksheet.write(i+2, 2, value, money_fmt)
            worksheet.merge_range(5, 0, 5, 1, 'SALDO FINAL DISPONIBLE', saldo_final_lbl_fmt)
            worksheet.write(5, 2, saldo_final, saldo_final_val_fmt)
            row_idx = 8
            worksheet.write(row_idx, 0, '1. RESUMEN POR CATEGORÍA', workbook.add_format({'bold': True, 'font_color': '#2F5597'}))
            row_idx += 1
            df_res_exp = tabla_resumen.drop(columns=['Orden'])
            for c, col_name in enumerate(df_res_exp.columns): worksheet.write(row_idx, c, col_name, header_fmt)
            row_idx += 1
            for row_data in df_res_exp.values:
                is_ingreso = "(+) INGRESO" in str(row_data[0])
                fmt_t, fmt_m = (ingreso_text_fmt, ingreso_money_fmt) if is_ingreso else (egreso_text_fmt, egreso_money_fmt)
                worksheet.write(row_idx, 0, row_data[0], fmt_t); worksheet.write(row_idx, 1, row_data[1], fmt_t)
                worksheet.write(row_idx, 2, row_data[2], fmt_t); worksheet.write(row_idx, 3, row_data[3], fmt_m); row_idx += 1
            row_idx += 2
            worksheet.write(row_idx, 0, '2. DETALLE DE MOVIMIENTOS', workbook.add_format({'bold': True, 'font_color': '#2F5597'}))
            row_idx += 1

            # --- CAMBIO AQUÍ ---
            # --- SOLUCIÓN UNIFICADA PARA EXCEL ---
            # Creamos formatos que hereden el color pero fuercen la alineación a la izquierda
            unidad_ingreso_fmt = workbook.add_format({
                'align': 'left', 
                'border': 1, 
                'font_name': 'Arial', 
                'font_color': '#002060'
            })
            unidad_egreso_fmt = workbook.add_format({
                'align': 'left', 
                'border': 1, 
                'font_name': 'Arial', 
                'font_color': '#C00000'
            })

            # --- 2. PREPARACIÓN DE DATOS Y LIMPIEZA ---
            df_det_exp = tabla_maestra.drop(columns=['Orden'], errors='ignore').copy()
            
            def limpiar_detalle_excel_numerico(row):
                det = str(row['Detalle']).strip()
                if "(+) INGRESO" in str(row['Flujo']):
                    if 'Edificio' in det: return 'Edificio'
                    import re
                    numeros = re.findall(r'\d+', det)
                    if numeros:
                        # Retornamos INT para evitar la advertencia de "Número como Texto"
                        return int(numeros[0])
                    return det
                return det

            df_det_exp['Detalle'] = df_det_exp.apply(limpiar_detalle_excel_numerico, axis=1)

            # 3. Lógica de Ordenamiento
            def orden_export(row):
                det = row['Detalle']
                if "(+) INGRESO" in str(row['Flujo']):
                    if det == 'Edificio': return 0
                    try: return int(det)
                    except: return 9998
                return 9999

            df_det_exp['temp_sort'] = df_det_exp.apply(orden_export, axis=1)
            df_det_exp = df_det_exp.sort_values('temp_sort', ascending=True).reset_index(drop=True)
            
            # Eliminamos columnas auxiliares
            df_det_exp = df_det_exp.drop(columns=['temp_sort', 'Unidad'], errors='ignore')

            # --- 4. ESCRITURA EN EL EXCEL ---
            # Escribir Encabezados
            for c, col_name in enumerate(df_det_exp.columns): 
                worksheet.write(row_idx, c, col_name, header_fmt)
            
            row_idx += 1
            # Escribir Datos
            for row_data in df_det_exp.values:
                is_ingreso = "(+) INGRESO" in str(row_data[0])
                
                # Asignamos formatos según el flujo (Ingreso/Egreso)
                if is_ingreso:
                    fmt_t, fmt_m, fmt_u = ingreso_text_fmt, ingreso_money_fmt, unidad_ingreso_fmt
                else:
                    fmt_t, fmt_m, fmt_u = egreso_text_fmt, egreso_money_fmt, unidad_egreso_fmt
                
                for c, val in enumerate(row_data):
                    if c == 4: # Columna Monto (Índice 4)
                        worksheet.write(row_idx, c, val, fmt_m)
                    elif c == 3: # Columna Detalle (Índice 3 - Donde están las unidades)
                        # Usamos el formato que alinea a la izquierda
                        worksheet.write(row_idx, c, val, fmt_u)
                    else:
                        # Resto de columnas (Flujo, Categoría, Sub-Categoría)
                        worksheet.write(row_idx, c, val, fmt_t)
                row_idx += 1

            for i, col in enumerate(df_det_exp.columns):
                max_len = max(df_det_exp[col].astype(str).map(len).max(), len(str(col))) + 2
                worksheet.set_column(i, i, min(max_len, 40))
        return output.getvalue()

    def generate_pdf_estado(mes, s_ini, ing_p, gas_p, s_fin, df_res, df_det, df_control=None):
        from fpdf import FPDF
        class PDF_Report(FPDF):
            def header(self):
                # --- FORZAMOS COLOR AZUL PARA EL TÍTULO ---
                self.set_text_color(47, 85, 151)
                self.set_font("Arial", "B", 14)

                # --- CAMBIO DINÁMICO DE TÍTULO SEGÚN ORIENTACIÓN ---
                if self.cur_orientation == 'L':
                    self.cell(0, 10, f"MATRIZ DE CONTROL DE PAGOS - PERIODO {mes}", ln=True, align='C')
                else:
                    self.cell(0, 10, f"ESTADO DE CUENTA MENSUAL - PERIODO {mes}", ln=True, align='C')
                
                self.set_font("Arial", "I", 9); self.set_text_color(100, 100, 100)
                self.cell(0, 5, "Edificio Creativa II", 0, 0, 'L')
                self.cell(0, 5, f"Pag. {self.page_no()}", 0, 1, 'R')
                
                # Ajuste de línea horizontal según ancho de página
                line_width = 260 if self.cur_orientation == 'L' else 200
                self.line(10, 26, line_width, 26); self.ln(5)
                # Un poco más de espacio tras la línea
                self.ln(6)
                
                # Solo mostrar encabezado de detalle en hojas verticales (P)
                if self.cur_orientation == 'P' and self.page_no() > 1:
                    self.set_font("Arial", "B", 10); self.set_text_color(47, 85, 151)
                    self.cell(0, 8, "2. DETALLE DE MOVIMIENTOS (Continuación)", ln=True)
                    self.set_fill_color(47, 85, 151); self.set_text_color(255, 255, 255)
                    self.cell(30, 8, "Flujo", 1, 0, 'C', True); self.cell(45, 8, "Categoria", 1, 0, 'C', True)
                    self.cell(85, 8, "Detalle", 1, 0, 'C', True); self.cell(30, 8, "Monto", 1, 1, 'C', True)

        pdf = PDF_Report(orientation='P', unit='mm', format='Letter')
        pdf.set_auto_page_break(auto=True, margin=20); pdf.add_page()
        pdf.set_font("Arial", "B", 11); pdf.set_text_color(0, 0, 0)
        res_data = [("Saldo Inicial (Anterior)", s_ini), ("Total Ingresos Mes", ing_p), ("Total Egresos Mes", gas_p)]
        for lab, val in res_data:
            pdf.set_fill_color(240, 240, 240); pdf.cell(110, 9, lab, 1, 0, 'L', 1)
            pdf.cell(0, 9, f"$ {formato_clp(val)}", 1, 1, 'R')
        pdf.set_fill_color(217, 234, 211); pdf.cell(110, 9, "SALDO FINAL DISPONIBLE", 1, 0, 'L', 1)
        pdf.cell(0, 9, f"$ {formato_clp(s_fin)}", 1, 1, 'R'); pdf.ln(8)
        pdf.set_font("Arial", "B", 10); pdf.set_text_color(47, 85, 151)
        pdf.cell(0, 8, "1. CONSOLIDADO POR CATEGORIA Y SUB-CATEGORIA", ln=True)
        pdf.set_fill_color(47, 85, 151); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", "B", 8)
        pdf.cell(30, 8, "Flujo", 1, 0, 'C', 1); pdf.cell(55, 8, "Categoria", 1, 0, 'C', 1)
        pdf.cell(55, 8, "Sub-Categoria", 1, 0, 'C', 1); pdf.cell(50, 8, "Monto", 1, 1, 'C', 1)
        pdf.set_font("Arial", "", 8)
        for _, f in df_res.iterrows():
            c = (0, 32, 96) if "(+) INGRESO" in str(f['Flujo']) else (192, 0, 0)
            pdf.set_text_color(*c); pdf.cell(30, 7, str(f['Flujo']), 1)
            pdf.cell(55, 7, str(f['Categoría']), 1); pdf.cell(55, 7, str(f.get('Sub-Categoría', '')), 1)
            pdf.cell(50, 7, f"$ {formato_clp(f['Monto'])}", 1, 1, 'R')
        pdf.ln(8); pdf.set_font("Arial", "B", 10); pdf.set_text_color(47, 85, 151)
        pdf.cell(0, 8, "2. DETALLE DE MOVIMIENTOS", ln=True)

        # --- CAMBIO AQUÍ ---
        # --- SOLUCIÓN UNIFICADA PARA PDF ---
        df_det_pdf = df_det.copy()

   
#-----     
        # 1. Limpieza de Detalle (UNIFICADA CON LA APP WEB)
        def limpiar_detalle_pdf(row):
            det = str(row['Detalle']).strip()
            sub_cat = str(row.get('Sub-Categoría', '')).strip()
            flujo = str(row['Flujo'])
            
            # Si es un ingreso de un Departamento (igual que en la web)
            if "(+) INGRESO" in flujo and sub_cat == "Departamento":
                import re
                # Buscamos el número de unidad
                unidad_match = re.search(r'(\d+)', det)
                if unidad_match:
                    return f"Departamento - {unidad_match.group(1)}"
                return "Departamento"
            
            # Caso especial para ingresos del Edificio
            if "(+) INGRESO" in flujo and 'Edificio' in det:
                return 'Edificio'
                
            return det

        df_det_pdf['Detalle'] = df_det_pdf.apply(limpiar_detalle_pdf, axis=1)


        # 2. Lógica de Orden (Ajustada para detectar el número dentro del nuevo formato)
        def orden_pdf_logic(row):
            det = str(row['Detalle'])
            if "(+) INGRESO" in str(row['Flujo']):
                if 'Edificio' in det: return 0
                import re
                # Extraemos el número incluso si dice "Departamento - 201"
                num = re.search(r'(\d+)', det)
                if num:
                    return int(num.group(1))
                return 9998
            return 9999
#-----


        df_det_pdf['temp_sort'] = df_det_pdf.apply(orden_pdf_logic, axis=1)
        df_det_pdf = df_det_pdf.sort_values('temp_sort', ascending=True).reset_index(drop=True)
        # -----------------------------------------------------------------

        pdf.set_fill_color(47, 85, 151); pdf.set_text_color(255, 255, 255)
        pdf.cell(30, 8, "Flujo", 1, 0, 'C', 1); pdf.cell(45, 8, "Categoria", 1, 0, 'C', 1)
        pdf.cell(85, 8, "Detalle", 1, 0, 'C', 1); pdf.cell(30, 8, "Monto", 1, 1, 'C', True)
        pdf.set_font("Arial", "", 7)

        # IMPORTANTE: Ahora iteramos sobre el nuevo df_det_pdf ordenado
        for _, f in df_det_pdf.iterrows(): 
            c = (0, 32, 96) if "(+) INGRESO" in str(f['Flujo']) else (192, 0, 0)
            pdf.set_text_color(*c); pdf.cell(30, 7, str(f['Flujo']), 1)
            pdf.cell(45, 7, str(f['Categoría'])[:25], 1); pdf.cell(85, 7, str(f['Detalle'])[:48], 1)
            pdf.cell(30, 7, f"$ {formato_clp(f['Monto'])}", 1, 1, 'R')

 
        # --- NUEVA HOJA: MATRIZ DE CONTROL DE PAGOS (HOJA 3 - ORIENTACIÓN HORIZONTAL) ---
        if df_control is not None:
            pdf.add_page(orientation='L')
            pdf.set_font("Arial", "B", 10); pdf.set_text_color(47, 85, 151)
            pdf.cell(0, 8, "3. HISTORIAL DE PAGOS REGISTRADOS (Vista 16 meses)", ln=True)
            pdf.ln(2)

            # --- INSERCIÓN DE NOTA ACLARATORIA ---
            pdf.set_font("Arial", "I", 8); pdf.set_text_color(100, 100, 100)
            nota_historial = (f"Nota: Este historial refleja todos los abonos registrados a la fecha de emisión, "
                              f"incluyendo pagos anticipados posteriores al periodo de {mes_seleccionado}.")
            pdf.multi_cell(0, 5, nota_historial)
            pdf.ln(2)
            # -------------------------------------

            # Selección de columnas: Unidad + últimos 16 meses
            cols_m = [df_control.columns[0]] + list(df_control.columns[1:][-16:])
            
            # --- LÓGICA DE CENTRADO ---
            w_u = 25  # Ancho columna Unidad
            # Definimos un ancho total para la tabla (ejemplo 250mm para dejar margen en Landscape)
            ancho_total_tabla = 255 
            w_mes = (ancho_total_tabla - w_u) / (len(cols_m) - 1)
            
            # Calculamos la posición X inicial para que quede centrada
            # (Ancho de página - Ancho de tabla) / 2
            pos_x_centrada = (pdf.w - ancho_total_tabla) / 2
            # ---------------------------

            # Encabezado Matriz
            pdf.set_fill_color(47, 85, 151); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", "B", 7)
            
            # Posicionamos en el X calculado
            pdf.set_x(pos_x_centrada)
            
            for col in cols_m:
                pdf.cell(w_u if col == cols_m[0] else w_mes, 8, str(col), 1, 0, 'C', True)
            pdf.ln()

            # Cuerpo de la Matriz
            pdf.set_text_color(0, 0, 0)
            for _, row in df_control.iterrows():
                # Importante: En cada línea nueva, debemos volver a centrar la X
                pdf.set_x(pos_x_centrada)
                
                pdf.set_font("Arial", "B", 7); pdf.set_fill_color(245, 245, 245)
                
                # --- CORRECCIÓN DE NÚMERO DE DEPARTAMENTO ---
                val_unidad = str(row[cols_m[0]])
                try:
                    val_unidad = str(int(float(val_unidad)))
                except:
                    val_unidad = val_unidad.replace(".0", "")
                
                pdf.cell(w_u, 6, val_unidad, 1, 0, 'C', True)
                
                pdf.set_font("Arial", "", 7)
                for m in cols_m[1:]:
                    val = str(row[m]).strip().lower()
                    if val in ["1", "1.0"]:
                        pdf.set_fill_color(212, 237, 218) # Verde
                        pdf.cell(w_mes, 6, "OK", 1, 0, 'C', True)
                    elif val in ["0", "0.0", "nan", "none", ""]:
                        pdf.cell(w_mes, 6, "", 1, 0, 'C') # Vacío
                    else:
                        pdf.set_fill_color(255, 243, 205) # Amarillo
                        pdf.cell(w_mes, 6, "?", 1, 0, 'C', True)
                pdf.ln()

        return pdf.output(dest='S')

    # 4. EJECUCIÓN OBLIGATORIA (Aquí es donde se crea la variable para la web)
    pdf_raw = generate_pdf_estado(
        mes_seleccionado, 
        saldo_inicial, 
        ingresos_puros, 
        gastos_puros, 
        saldo_final, 
        tabla_resumen, 
        tabla_maestra,
        df_control     # Muestra la Tercera hoja en el EDC en el portal
    )

#----Fin del proceso de exportación a PDF


#--
# --- LÓGICA DE EXPORTACIÓN A PDF (Pestaña Control de Pagos) ---
    def generate_pdf_morosidad(df_matriz):
        from fpdf import FPDF
        from datetime import datetime
        
        class PDF_M(FPDF):
            def header(self):
                self.set_font("Arial", "B", 14)
                self.set_text_color(47, 85, 151)
                # Título centrado para hoja Carta Horizontal (279.4mm)
                _ = self.cell(0, 10, f"MATRIZ DE CONTROL DE PAGOS - {datetime.now().strftime('%d/%m/%Y')}", ln=True, align='C')
                self.ln(5)

        # 'L' = Landscape (Horizontal), 'Letter' = Carta
        pdf = PDF_M(orientation='L', unit='mm', format='Letter')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # --- CONFIGURACIÓN DE COLUMNAS (Para que quepan 16 meses + Depto) ---
        # Anchos: Depto (35mm) + 16 meses (14mm c/u) = 259mm totales (perfecto para Carta Horizontal)
        col_unid_width = 35
        mes_width = 14
        
        # --- ENCABEZADO DE TABLA ---
        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(47, 85, 151)
        pdf.set_text_color(255, 255, 255)
        
        # Columna Departamento
        columnas = df_matriz.columns.tolist()
        col_id = columnas[0]
        meses = columnas[1:]
        
        _ = pdf.cell(col_unid_width, 8, "Depto", 1, 0, 'C', True)
        
        for mes in meses:
            # Limpiamos el nombre del mes
            nombre_mes = str(mes).replace("<br>", " ")
            _ = pdf.cell(mes_width, 8, nombre_mes, 1, 0, 'C', True)
        pdf.ln()
        
        # --- DATOS ---
        pdf.set_font("Arial", "", 8)
        pdf.set_text_color(0, 0, 0)
        
        for _, fila in df_matriz.iterrows():
            # Celda Departamento (Gris suave para distinguir filas)
            pdf.set_fill_color(245, 245, 245)
            pdf.set_font("Arial", "B", 8)
            _ = pdf.cell(col_unid_width, 7, str(fila[col_id]), 1, 0, 'C', True)
            
            # Celdas de meses
            pdf.set_font("Arial", "", 8)
            for mes in meses:
                valor = str(fila[mes]).strip().lower()
                
                # Si el valor es 1 (Pagado)
                if valor in ["1", "1.0"]:
                    pdf.set_fill_color(212, 237, 218) # Verde claro
                    pdf.set_text_color(21, 87, 36)    # Texto verde oscuro
                    _ = pdf.cell(mes_width, 7, "OK", 1, 0, 'C', True)
                else:
                    # Fondo blanco para deuda (0 o vacío)
                    pdf.set_fill_color(255, 255, 255)
                    _ = pdf.cell(mes_width, 7, "", 1, 0, 'C', True)
                
                pdf.set_text_color(0, 0, 0)
            pdf.ln()

        # Retornamos los bytes finales
        return pdf.output()

#--
# --- LÓGICA DE EXPORTACIÓN A PDF (Pestaña Control Morosidad) ---
    def generate_pdf_morosidad_historica(df, mes_corte):
        from fpdf import FPDF
        from datetime import datetime
        
        class PDF_MH(FPDF):
            def header(self):
                self.set_font("Arial", "B", 14)
                self.set_text_color(31, 78, 120)
                _ = self.cell(0, 10, f"REPORTE DE MOROSIDAD HISTORICA - CORTE {mes_corte}", ln=True, align='C')
                self.set_font("Arial", "I", 9)
                self.set_text_color(100, 100, 100)
                _ = self.cell(0, 5, f"Fecha de extraccion: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='R')
                self.ln(5)

        pdf = PDF_MH(orientation='L', unit='mm', format='Letter')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Anchos: Unidad(20), Residente(55), Meses(15), Deuda(35), Gravedad(25), Detalle(100)
        widths = [20, 55, 15, 35, 25, 100]
        headers = ["Unidad", "Residente", "Meses", "Deuda Total", "Gravedad", "Detalle Meses"]
        
        pdf.set_font("Arial", "B", 9)
        pdf.set_fill_color(31, 78, 120)
        pdf.set_text_color(255, 255, 255)
        for i, h in enumerate(headers):
            _ = pdf.cell(widths[i], 8, h, 1, 0, 'C', True)
        pdf.ln()
        
        pdf.set_font("Arial", "", 8)
        pdf.set_text_color(0, 0, 0)
        
        for _, fila in df.iterrows():
            # LIMPIEZA DE EMOJIS: Quitamos los círculos para evitar el error de fuente
            gravedad_texto = str(fila['Gravedad']).replace("🔴 ", "").replace("🟡 ", "")
            
            # Definir color de texto según gravedad (Rojo para Crítico, Naranja para Riesgo)
            if "Crítico" in gravedad_texto:
                txt_color = (192, 0, 0) 
            else:
                txt_color = (255, 140, 0)
            
            _ = pdf.cell(widths[0], 7, str(fila['Unidad']), 1, 0, 'C')
            _ = pdf.cell(widths[1], 7, str(fila['Residente'])[:30], 1, 0, 'L')
            _ = pdf.cell(widths[2], 7, str(fila['Meses Deuda']), 1, 0, 'C')
            _ = pdf.cell(widths[3], 7, f"$ {formato_clp(fila['Deuda Total Est.'])}", 1, 0, 'R')
            
            # Celda Gravedad: Usamos el texto limpio y le aplicamos el color
            pdf.set_text_color(*txt_color)
            pdf.set_font("Arial", "B", 8)
            _ = pdf.cell(widths[4], 7, gravedad_texto, 1, 0, 'C')
            
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", "", 7)
            
            # Limpieza básica de acentos o caracteres especiales en el detalle si fuera necesario
            detalle_limpio = str(fila['Detalle Meses']).replace("<br>", " ")
            _ = pdf.cell(widths[5], 7, detalle_limpio[:70], 1, 1, 'L')
            
        return pdf.output()

# --- FIN LÓGICA DE EXPORTACIÓN A PDF (Pestaña Control Morosidad) ---


# --- BLOQUE DE EXPORTACION EN SIDEBAR ---
    st.sidebar.markdown("### 📥 Descargar Estado de Cuenta")
    col_ex, col_pdf = st.sidebar.columns(2)

    with col_ex:
        st.download_button(
            label="❎ Excel", 
            data=generate_excel(), 
            file_name=f"Estado_de_Cuenta_{mes_seleccionado.replace('/','-')}.xlsx",
            use_container_width=True
        )

    with col_pdf:
        # 1. Llamamos a la función y guardamos el resultado
        # El uso de 'bytes()' asegura que Streamlit reciba datos puros
        try:
            pdf_final = bytes(generate_pdf_estado(
                mes_seleccionado, 
                saldo_inicial, 
                ingresos_puros, 
                gastos_puros, 
                saldo_final, 
                tabla_resumen,
                tabla_maestra
            ))
            
            # 2. El botón usa los datos ya guardados
            st.download_button(
                label="📕 PDF", 
                data=pdf_final, 
                file_name=f"Estado_de_Cuenta_{mes_seleccionado.replace('/','-')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="btn_pdf_v99"
            )
        except:
            pass # Si hay algún None residual, este bloque lo atrapa
#----

# --- SECCIÓN DE CRÉDITOS (ESTILO FOOTER DISCRETO) ---

    # 1. Definimos la función de información (puedes mantener el @st.dialog anterior si prefieres)
    @st.dialog("Información del Sistema")
    def mostrar_info():
        st.markdown(f"""
            ### 🏢 Sistema de Gestión Financiera - Edificio Creativa II
            **Versión 1.0** (2026)
            
            Esta aplicación es una herramienta propietaria diseñada y desarrollada 
            por **Daniel Guerrero Borbonet**. 
            
            Su propósito es centralizar la información financiera y proporcionar 
            capacidades de simulación avanzada para la toma de decisiones 
            estratégicas en la administración, fortaleciendo la continuidad 
            operativa y la transparencia hacia la comunidad mediante una gestión
            activa del servicio en tiempo real.        

            ---
            *Diseño basado en **ciencia de datos**, que combina estadística, 
             matemáticas, programación e inteligencia artificial, analítica y
             lógica, para analizar datos masivos. En este sistema, se aplica para
             transformar registros contables complejos en claridad estratégica y
             transparencia financiera.*

            *Desarrollado con Python y Streamlit.*
        """)

    st.sidebar.markdown("<br>" * 3, unsafe_allow_html=True) # Espacio de cortesía
    st.sidebar.divider()

    # 2. Centramos el Botón y el Copyright usando HTML/CSS
    st.sidebar.markdown("""
        <style>
            /* Estilo para que el botón se vea centrado y elegante */
            .footer-container {
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                width: 100%;
                gap: 10px;
            }
        </style>
        <div class="footer-container">
    """, unsafe_allow_html=True)

    # 3. Botón de "Acerca de"
    if st.sidebar.button("❔ Acerca de esta App", key="btn_centrado", use_container_width=True):
        mostrar_info()

    # Al final de tu sidebar
    st.sidebar.markdown(
        "<div style='text-align: center; opacity: 0.7; font-size: 0.8rem;'>"
        "© 2026 Edificio Creativa II - Todos los Derechos Reservados"
        "</div>", 
        unsafe_allow_html=True
    )
            
    tab_bal, tab_ing, tab_egr, tab_salud, tab_mor, tab_control, tab_avisos, tab_utilidad = st.tabs(["📊 BALANCE GENERAL", "💰 INGRESOS", "📉 EGRESOS", "📈 RESUMEN FINANCIERO", "🚦 CONTROL MOROSIDAD", "🗓️ CONTROL DE PAGOS", "✉️ GENERADOR DE AVISOS", "🧮 SIMULADOR DE REAJUSTES"])

# --- PESTAÑA BALANCE GENERAL ---   
with tab_bal:
    st.write(f"### 📜 Estado de Cuenta - {mes_seleccionado}")

    # --- SUGERENCIA: Nota informativa de exclusiones ---
    excluidos_hoy = calcular_excluidos_del_mes(mes_seleccionado, DICCIONARIO_EXCLUSIONES)
    if excluidos_hoy:
        st.info(f"ℹ️ **Nota de Gestión:** En este periodo existen {len(excluidos_hoy)} unidades exentas de cobro según regla de negocio.")

    # --- 1. CÁLCULOS CENTRALIZADOS (Para evitar NameError y repeticiones) ---
    ratio_operacion = (gastos_puros / ingresos_puros * 100) if ingresos_puros > 0 else 0
    variacion_pct = ratio_operacion - 100
    diferencia_dinero = gastos_puros - ingresos_puros
    
    total_recursos_disponibles = saldo_inicial + ingresos_puros
    supervivencia_caja = (saldo_final / total_recursos_disponibles * 100) if total_recursos_disponibles > 0 else 0

    # --- 2. BLOQUE DE MÉTRICAS PRINCIPALES ---
    with st.container(border=True):
        st.write("### 📊 Análisis Financiero del Período")
        c1, c2, c3, c4 = st.columns(4)
        
        c1.metric("Saldo Inicial", f"$ {formato_clp(saldo_inicial)}")
        c2.metric("Ingresos (+)", f"$ {formato_clp(ingresos_puros)}")
        
        # COLUMNA 3: EGRESOS (Con signo forzado para activar color)
        c3.metric(
            label="Egresos (-)", 
            value=f"$ {formato_clp(gastos_puros)}", 
            delta=f"{variacion_pct:+.1f}%", 
            delta_color="inverse" 
        )
        
        c4.metric("Saldo Final", f"$ {formato_clp(saldo_final)}")

    # --- 3. MENSAJE DE RESUMEN (Fuera del contenedor para no afectar altura) ---
    if variacion_pct > 0:
        # Usamos warning (amarillo) porque hubo un déficit
        st.warning(f"💡 **Déficit Mensual:** Los gastos superaron la recaudación en un **{variacion_pct:.1f}%**.")
    else:
        # Usamos success (verde) porque hubo excedente
        st.success(f"💡 **Resultado Positivo:** La recaudación superó los gastos del mes, generando un **{abs(variacion_pct):.1f}%** de excedente.")

    # --- 4. GESTIÓN DE CAJA Y SOLVENCIA ---
    st.write("### ⚖️ Gestión de Caja y Solvencia")

    col_i1, col_i2 = st.columns([2, 1])

    with col_i1:
        if gastos_puros > ingresos_puros:
            # Ya usamos diferencia_dinero calculada arriba
            st.warning(f"""
                **Nota:** Los egresos superaron la recaudación en un {variacion_pct:.1f}%. 
                La diferencia **$ {formato_clp(diferencia_dinero)}** se cubrió con el saldo del mes anterior (Saldo Inicial).
            """)
        else:
            excedente = abs(diferencia_dinero)
            st.success(f"""
                **Nota:** La recaudación cubrió el 100% de los egresos. 
                Se generó un excedente de **$ {formato_clp(excedente)}** que se suma al saldo del próximo mes.
            """)
        #     st.success(f"**Nota:** La recaudación cubrió los gastos y generó un excedente de **$ {formato_clp(excedente)}**.")
        
        # Barra de progreso
        st.write(f"**Disponibilidad Total de Fondos: {supervivencia_caja:.1f}%**")
        st.progress(min(supervivencia_caja / 100, 1.0))
        st.caption("Porcentaje de dinero remanente sobre el total de recursos que circularon en el mes.")

    with col_i2:
        variacion_patrimonio = saldo_final - saldo_inicial
        st.metric(
            label="Variación Patrimonial", 
            value=f"$ {formato_clp(saldo_final)}",
            delta=f"$ {formato_clp(variacion_patrimonio)} este mes",
            delta_color="normal" if variacion_patrimonio >= 0 else "inverse"
        )

# ---
    # --- GRÁFICO DE INDICADOR (GAUGE) ---
    import plotly.graph_objects as go

    # --- CONFIGURACIÓN DE NARRATIVA ---
    if ratio_operacion > 100:
        detalle_ratio = f"⚠️ Gastamos {ratio_operacion-100:.1f}% más de lo que recaudamos"
        color_ratio = "#dc3545" # Rojo
    else:
        detalle_ratio = f"✅ Gastamos solo el {ratio_operacion:.1f}% de lo recaudado"
        color_ratio = "#28a745" # Verde

    # --- GRÁFICO DE INDICADOR (GAUGE) CON LEYENDA ---
    fig_gauge = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = ratio_operacion,
        number = {
            'suffix': "%", 
            'font': {'color': color_ratio, 'size': 50},
            'valueformat': ".1f"
        },
        domain = {'x': [0, 1], 'y': [0.15, 1]}, # Subimos el gráfico para dejar espacio abajo
        title = {
            'text': "<b>Índice de Gasto Operativo</b><br><span style='font-size:0.8em;color:gray'>(Egresos / Ingresos del Mes)</span>", 
            'font': {'size': 20}
        },
        gauge = {
            'axis': {'range': [0, max(150, ratio_operacion + 20)], 'ticksuffix': "%"},
            'bar': {'color': "#31333F"},
            'steps': [
                {'range': [0, 100], 'color': '#e1f5fe'},   # Azul claro: Rango de Ingresos
                {'range': [100, 200], 'color': '#f8d7da'}  # Rojo claro: Uso de Ahorros
            ],
            'threshold': {
                'line': {'color': "black", 'width': 4},
                'thickness': 0.75,
                'value': 100
            }
        }
    ))

    # Añadimos la explicación de texto "pegada" al gráfico
    fig_gauge.add_annotation(
        text=f"<b>{detalle_ratio}</b>",
        x=0.5, y=0.1,  # Posición centrada abajo
        showarrow=False,
        font=dict(size=16, color=color_ratio),
        align="center"
    )

    fig_gauge.update_layout(height=400, margin=dict(l=30, r=30, t=80, b=20))
    st.plotly_chart(fig_gauge, use_container_width=True)
    st.write("---")
 # --- 

    # --- FUNCIÓN DE FORMATEO ÚNICA Y ULTRA-RESISTENTE ---
    def formatear_monto_flujo(fila):
        monto_f = f"$ {formato_clp(fila['Monto'])}"
        # Extraemos el valor de flujo, lo pasamos a minúsculas y quitamos espacios
        valor_flujo = str(fila.get('Flujo', '')).lower().strip()
        
        # Si empieza con 'e' (egreso, egresos, e) o contiene 'egr'
        if valor_flujo.startswith('e') or 'egr' in valor_flujo:
            return f"🔴 {monto_f}"
        else:
            return f"🟢 {monto_f}"

    # 2. Resumen por Categoría
    st.write("### 🏷️ Resumen por Categoría")
    st.caption("Vista consolidada de los movimientos financieros clasificados por flujo de caja y categoría.")

    # --- OPCIÓN DE VISTA (SWITCH) ---
    modo_vista = st.radio(
        "Seleccionar formato de visualización:",
        ["📋 Tabla de Resumen", "📊 Gráfico de Distribución"],
        horizontal=True,
        key="switch_vista_resumen"
    )

    # Preparamos los datos base para la tabla
    df_resumen_display = tabla_resumen.drop(columns=['Orden'], errors='ignore').copy()
    
    if modo_vista == "📋 Tabla de Resumen":
        if 'Monto' in df_resumen_display.columns:
            df_resumen_display['Monto'] = df_resumen_display.apply(formatear_monto_flujo, axis=1)
        st.dataframe(df_resumen_display, use_container_width=True, hide_index=True)
    
    else:
        import plotly.express as px
        import pandas as pd
        
        # --- SOLUCIÓN MIXTA DEFINITIVA ---
        df_base_plot = tabla_resumen.copy() # Usamos una copia con nombre único
        df_base_plot['Tipo_Color'] = df_base_plot['Flujo'].astype(str).str.upper().str.strip().apply(
            lambda x: 'INGRESO' if 'INGR' in x else 'EGRESO'
        )
        
        # 1. Procesamos Ingresos y Egresos por separado
        # IMPORTANTE: Cambiamos df_ing por df_ing_plot y df_egr por df_egr_plot 
        # para NO sobrescribir las variables globales de la app.
        df_ing_plot = df_base_plot[df_base_plot['Tipo_Color'] == 'INGRESO'].copy()
        df_egr_plot = df_base_plot[df_base_plot['Tipo_Color'] == 'EGRESO'].groupby(
            ['Categoría', 'Flujo', 'Tipo_Color'], sort=False
        )['Monto'].sum().reset_index()
        
        # 2. UNIÓN: Para que los verdes salgan arriba en Plotly sin 'reversed'
        df_plot = pd.concat([df_ing_plot, df_egr_plot], ignore_index=True)
        
        # 3. ETIQUETA MANUAL
        df_plot['Monto_Texto'] = df_plot['Monto'].apply(lambda x: f"${formato_clp(int(x))}")
        
        color_map = {"INGRESO": "#28a745", "EGRESO": "#dc3545"}

        fig_resumen = px.bar(
            df_plot,
            x="Monto",
            y="Categoría",
            color="Tipo_Color",
            orientation='h',
            text="Monto_Texto", 
            color_discrete_map=color_map,
            labels={"Monto": "Monto Total ($)", "Categoría": "", "Tipo_Color": "Flujo"},
            category_orders={"Categoría": df_plot["Categoría"].tolist()}
        )

        # --- CONFIGURACIÓN DEL LAYOUT ---
        fig_resumen.update_layout(
            font=dict(size=16),
            separators=",.", 
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.35,
                xanchor="center",
                x=0.5,
                title_text=""
            ),
            margin=dict(l=20, r=120, t=20, b=100), 
            height=450,
            xaxis=dict(
                tickfont=dict(size=14),
                tickformat=",.0f", 
                showgrid=True,
                gridcolor='rgba(200,200,200,0.2)'
            ),
            yaxis=dict(
                tickfont=dict(size=15),
                autorange=True 
            )
        )
        
        fig_resumen.update_traces(
            textfont_size=14,
            textposition='outside',
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Monto: %{text}<extra></extra>"
        )

        st.plotly_chart(fig_resumen, use_container_width=True)

    # --- SEPARADOR ELEGANTE ---
    st.markdown("---")

    # 3. Detalle Total de Movimientos
    st.write("### 🔎 Detalle Total de Movimientos")
    
    # 1. Copia y limpieza de nombres de columnas
    df_maestra_display = tabla_maestra.drop(columns=['Orden'], errors='ignore').copy()
    df_maestra_display.columns = [str(c).strip() for c in df_maestra_display.columns]

    # 2. PROCESAMIENTO INTELIGENTE DE COLUMNA 'DETALLE'
    #    (UNIFICADO CON PORTAL Y PDF)
    if 'Detalle' in df_maestra_display.columns and 'Flujo' in df_maestra_display.columns:
        def procesar_detalle_equilibrado(fila):
            detalle = str(fila['Detalle']).strip()
            flujo = str(fila['Flujo']).upper()
            sub_cat = str(fila.get('Sub-Categoría', '')).strip().upper()
            
            # Si es un ingreso de un Departamento
            if 'INGRESO' in flujo and (sub_cat == "DEPARTAMENTO" or 'DEPARTAMENTO' in detalle.upper()):
                if 'Edificio' in detalle: return 'Edificio'
                
                import re
                # Buscamos el número (ej: de "201 - Daniel" sacamos "201")
                unidad_match = re.search(r'(\d+)', detalle)
                if unidad_match:
                    return f"Departamento - {unidad_match.group(1)}"
                return "Departamento"
            
            return detalle

        df_maestra_display['Detalle'] = df_maestra_display.apply(procesar_detalle_equilibrado, axis=1)

        # 3. ORDENAMIENTO NATURAL (TIPO WINDOWS 10)
        def logica_orden_final(fila):
            detalle = str(fila['Detalle'])
            flujo = str(fila['Flujo']).upper()
            
            if 'INGRESO' in flujo:
                if detalle == 'Edificio': return 0
                import re
                # Extraemos el número para ordenar numéricamente (1, 2, 10...)
                num_match = re.search(r'(\d+)', detalle)
                if num_match:
                    return int(num_match.group(1))
                return 9998
            else:
                return 9999 # Egresos al final

        df_maestra_display['temp_sort'] = df_maestra_display.apply(logica_orden_final, axis=1)
        df_maestra_display = df_maestra_display.sort_values('temp_sort', ascending=True).reset_index(drop=True)
        df_maestra_display = df_maestra_display.drop(columns=['temp_sort'])

        st.caption("Registro completo de ingresos y egresos correspondientes al período seleccionado.")

    # 4. FORMATEO DE MONTOS (🟢/🔴)
    if 'Monto' in df_maestra_display.columns:
        df_maestra_display['Monto'] = df_maestra_display.apply(formatear_monto_flujo, axis=1)

    # 5. CONVERSIÓN A TEXTO PARA PANTALLA
    for col in df_maestra_display.columns:
        df_maestra_display[col] = df_maestra_display[col].astype(str).replace(['nan', '0.0'], '-')

    # 6. MOSTRAR TABLA
    st.dataframe(
        df_maestra_display, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "Detalle": st.column_config.Column("Unidad / Detalle", width="medium"),
            "Monto": st.column_config.Column("Monto (Flujo)", width="medium")
        }
    )


# --- GENERACIÓN DEL REPORTE EXTERNO HTML ---
    # 1. Preparamos el diccionario de métricas con narrativa incluida
    diferencia_dinero = gastos_puros - ingresos_puros
    
    # Definimos la nota de gestión según el resultado
    if gastos_puros > ingresos_puros:
        nota_gestion = f"Los egresos superaron la recaudación en un {variacion_pct:.1f}%. La diferencia $ {formato_clp(diferencia_dinero)} se cubrió con el saldo del mes anterior (Saldo Inicial)."
        color_gestion = "#856404" # Color café/amarillo (warning)
        bg_gestion = "#fff3cd"
    else:
        nota_gestion = f"La recaudación cubrió el 100% de los egresos. Se generó un excedente de $ {formato_clp(abs(diferencia_dinero))} para el próximo mes."
        color_gestion = "#155724" # Color verde (success)
        bg_gestion = "#d4edda"

    mis_metricas_web = {
        'saldo_i': f"$ {formato_clp(saldo_inicial)}",
        'ingresos': f"$ {formato_clp(ingresos_puros)}",
        'egresos': f"$ {formato_clp(gastos_puros)}",
        'saldo_f': f"$ {formato_clp(saldo_final)}",
        'supervivencia': f"{supervivencia_caja:.1f}%",
        'supervivencia_num': supervivencia_caja, # Para la barra de progreso
        'variacion_pct': f"{variacion_pct:+.1f}%",
        'nota_gestion': nota_gestion,
        'color_gestion': color_gestion,
        'bg_gestion': bg_gestion,
        'detalle_ratio': detalle_ratio,
        'color_ratio': color_ratio
    }

# --- FIN DE LA GENERACIÓN DEL REPORTE EXTERNO HTML ---


# --- PESTAÑA INGRESOS ---      
with tab_ing:
    st.write(f"### 💰 Análisis de Ingresos - {mes_seleccionado}")

    # --- NUEVA SECCIÓN: ANÁLISIS DE RECAUDACIÓN (ANTES DE GRÁFICOS) ---
    with st.container(border=True):
        # 1. Filtro estandarizado
        df_ggcc = df_ing_mes[
            df_ing_mes["Tipo de Ingreso"].astype(str).str.strip().str.upper() == "GASTOS COMUNES"
        ].copy()

        # --- MODIFICACIÓN: APLICAR EXCLUSIÓN EN TIEMPO REAL ---
        excluidos_actuales = calcular_excluidos_del_mes(mes_seleccionado, DICCIONARIO_EXCLUSIONES)
        # Filtramos el dataframe de cálculo para que el "Esperado" sea real
        df_ggcc_validos = df_ggcc[~df_ggcc['Unidad'].astype(str).str.replace(".0","").isin([str(e).replace(".0","") for e in excluidos_actuales])]
        
        # 2. Cálculos base (Usando el dataframe filtrado por exclusiones)
        monto_esperado = df_ggcc_validos["Monto Actual"].sum()
        monto_recaudado = df_ggcc_validos["Pago Total del Mes"].sum()
        monto_pendiente = monto_esperado - monto_recaudado
        # -----------------------------------------------------

        # 3. Cálculo de porcentajes
        pct_recaudacion = (monto_recaudado / monto_esperado * 100) if monto_esperado > 0 else 0
        pct_pendiente = 100 - pct_recaudacion
            
        st.write("#### 📊 Eficiencia de Recaudación (Gastos comunes)")
        
        # Nota informativa si hay excluidos
        if excluidos_actuales:
            st.caption(f"ℹ️ El monto esperado excluye a {len(excluidos_actuales)} unidades exentas este mes.")

        c1, c2, c3, c4 = st.columns(4)
        
        # Columna 1: Lo que el edificio debería recibir    
        c1.metric("Esperado", f"$ {formato_clp(monto_esperado)}")
            
        # Columna 2: Lo que entró (Solo verde si es >= 95%)
        c2.metric(
            "Recaudado", 
            f"$ {formato_clp(monto_recaudado)}", 
            delta=f"{pct_recaudacion:.1f}% del total",
            delta_color="normal" if pct_recaudacion >= 95 else "off"
        )
            
        # --- NUEVA LÓGICA UNIFICADA: DIFERENCIA Y ESTADO ---
        diferencia_valor = abs(monto_recaudado - monto_esperado)
        
        if monto_recaudado > monto_esperado:
            etiqueta_dif = "(A favor)"
            color_delta = "normal"  # Verde
        elif monto_recaudado < monto_esperado:
            etiqueta_dif = "(En contra)"
            color_delta = "inverse" # Rojo
        else:
            etiqueta_dif = ""
            color_delta = "off"     # Gris

        # Columna 3: Diferencia (Reemplaza a "Pendiente")
        c3.metric(
            "Diferencia", 
            f"$ {formato_clp(diferencia_valor)}", 
            delta=etiqueta_dif, 
            delta_color=color_delta
        )
            
        # 4. LÓGICA DE ESTADO DINÁMICO
        if pct_recaudacion >= 100:
            estado_salud = "✅ EXCELENTE"
        elif pct_recaudacion >= 95:
            estado_salud = "🟢 SATISFACTORIO"
        elif pct_recaudacion >= 85:
            estado_salud = "⚠️ ATENCIÓN"
        else:
            estado_salud = "🔴 CRÍTICO"
            
        c4.metric("Estado", estado_salud)
        # --- FIN NUEVA LÓGICA UNIFICADA ---
        
    st.write("") # Espaciado estético

    # --- 1. NUEVA VISUALIZACIÓN UNIFICADA: RECAUDACIÓN POR UNIDAD (COMPOSICIONAL) ---
    st.write("### 🏠 Detalle de Recaudación por Departamentos")

    if not df_ing_mes.empty:
        import plotly.express as px

        # 🟢 FILTRO: Solo departamentos (Gastos Comunes)
        # Ajusta 'Gastos comunes' al nombre exacto que tengas en tu columna 'Tipo de Ingreso'
        df_solo_deptos = df_ing_mes[df_ing_mes['Tipo de Ingreso'] == 'Gastos comunes'].copy()

        # 1. CÁLCULO DEL TOTAL DINÁMICO (Basado solo en el filtro)
        total_mes = df_solo_deptos['Pago Total del Mes'].sum()
        total_formateado = f"{total_mes:,.0f}".replace(",", ".")

        # 🔵 CONFIGURACIÓN DE PALETA (Al ser una sola categoría, usaremos un azul sólido)
        paleta_azul = ["#0D47A1"] 

        # A. Preparamos los datos para el gráfico
        def extraer_numero(u):
            import re
            nums = re.findall(r'\d+', str(u))
            return int(nums[0]) if nums else 9999
        
        df_solo_deptos['Unidad_Num'] = df_solo_deptos['Unidad'].apply(extraer_numero)
        df_solo_deptos = df_solo_deptos.sort_values('Unidad_Num')

        # C. Creamos el gráfico (Ya no es apilado porque es una sola categoría)
        fig_recaudacion = px.bar(
            df_solo_deptos,
            x='Unidad', 
            y='Pago Total del Mes',
            labels={'Pago Total del Mes': 'Monto Recaudado ($)', 'Unidad': 'Unidad'},
            color_discrete_sequence=paleta_azul
        )

        # D. Ajustes de diseño
        fig_recaudacion.update_layout(
            separators=",.", 
            height=450, # Un poco más bajo para que sea compacto
            font=dict(size=14, family="Source Sans Pro, sans-serif"),
            showlegend=False, # Quitamos leyenda ya que solo hay un color
            xaxis=dict(
                tickfont=dict(size=14),
                tickangle=0, 
                title=None,
                type='category',
                showline=True,
                linecolor='#E0E0E0'
            ),
            yaxis=dict(
                tickfont=dict(size=14, color='black'),
                tickprefix="$ ",
                tickformat=",.0f",
                title=None,
                gridcolor='rgba(200, 200, 200, 0.2)',
                zeroline=False
            ),
            margin=dict(l=20, r=20, t=30, b=50),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )

        fig_recaudacion.update_traces(
            marker_line_color='white',
            marker_line_width=1,
            hovertemplate="<b>Unidad %{x}</b><br>Recaudado: $ %{y:,.0f}<extra></extra>"
        )

        # --- APLICACIÓN DEL CONTENEDOR ---
        with st.container(border=True):
            st.markdown(
                f"""
                <div style="margin-bottom: 5px;">
                    <p style="color: #6B7280; font-size: 0.99rem; font-weight: bold; margin-bottom: 8px;">
                        Recaudación Gastos Comunes: <span style="color: #0D47A1;">$ {total_formateado}</span>
                    </p>
                    <hr style="margin: 0; border: none; border-top: 1px solid #F3F4F6;">
                </div>
                """, 
                unsafe_allow_html=True
            )

            # El gráfico debe ir pegado inmediatamente después del markdown de apertura
            st.plotly_chart(fig_recaudacion, use_container_width=True, config={'displayModeBar': False})

            # NOTA INFORMATIVA CON ESTILO AZUL CLARO
            st.info("ℹ️ **Nota:** Esta sección detalla los pagos recibidos de cada departamento. Los montos reflejan la recaudación efectiva de gastos comunes y otros conceptos específicos del período.")

            st.markdown("</div>", unsafe_allow_html=True)
        # ---------------------------------------------------

    else:
        st.info("No hay datos de recaudación para mostrar en este periodo.")

    st.markdown("---")

    # 2. TABLA DE DATOS MENSUALES
    st.write("### 🟩 **Detalle de Recaudación Total del Mes:**")
    
    df_ing_display = df_ing_mes.copy()

    for col in COLS_FINANCIERAS:
        if col in df_ing_display.columns:
            if col == "Pago Total del Mes":
                df_ing_display[col] = df_ing_display[col].apply(lambda x: f"🟢 $ {formato_clp(x)}")
            else:
                df_ing_display[col] = df_ing_display[col].apply(lambda x: f"$ {formato_clp(x)}")

    cols_a_texto = ['Unidad', 'Período', 'Año', 'Mes']
    for col in cols_a_texto:
        if col in df_ing_display.columns:
            df_ing_display[col] = df_ing_display[col].astype(str)

    st.dataframe(
        df_ing_display, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "Unidad": st.column_config.Column(width="small"),
            "Período": st.column_config.Column(width="small"),
            "Pago Total del Mes": st.column_config.Column(
                label="Pago Total del Mes",
                width="medium",
                help="Monto total recibido (🟢 indica ingreso)"
            ),
        }
    )

    total_ing_mes = df_ing_mes['Pago Total del Mes'].sum()
    
    col_ti1, col_ti2 = st.columns([3, 1])
    with col_ti2:
        st.markdown(f"### **Total: 🟢 $ {formato_clp(total_ing_mes)}**")
        st.caption("ℹ️ *Este monto incluye el saldo del mes anterior.*")

    st.markdown("---")
    
    # 3. LÓGICA DE EVOLUCIÓN HISTÓRICA
    st.write("### 📊 Evolución de Ingresos (Últimos Períodos)")

    try:
        # --- MODIFICACIÓN PARA HISTÓRICO: Filtrar excluidos mes a mes ---
        registros_hist_limpios = []
        for p in df_ing['Período'].unique():
            excl_p = calcular_excluidos_del_mes(p, DICCIONARIO_EXCLUSIONES)
            temp_p = df_ing[df_ing['Período'] == p].copy()
            # Quitamos los que no debían pagar en ese mes para que la barra histórica sea coherente
            temp_p = temp_p[~temp_p['Unidad'].astype(str).str.replace(".0","").isin([str(e).replace(".0","") for e in excl_p])]
            registros_hist_limpios.append(temp_p)
        
        df_ing_historico_filtrado = pd.concat(registros_hist_limpios) if registros_hist_limpios else df_ing
        
        df_ing_hist = df_ing_historico_filtrado[df_ing_historico_filtrado['Tipo de Ingreso'] != 'Saldo Mes Anterior'].groupby(['Período', 'Tipo de Ingreso'])['Pago Total del Mes'].sum().reset_index()
        # ---------------------------------------------------------------

        periodos_unicos_ing = df_ing['Período'].unique()
        try:
            periodos_ord_ing = sorted(periodos_unicos_ing, key=lambda x: datetime.strptime(str(x), '%m/%Y'))
        except:
            periodos_ord_ing = sorted(periodos_unicos_ing)
        
        ultimos_6_ing = periodos_ord_ing[-6:]
        df_plot_ing_hist = df_ing_hist[df_ing_hist['Período'].isin(ultimos_6_ing)]

        fig_ing_evol = px.bar(
            df_plot_ing_hist,
            x='Período',
            y='Pago Total del Mes',
            color='Tipo de Ingreso',
            title=None, 
            labels={'Pago Total del Mes': 'Total Recaudado ($)', 'Período': 'Mes/Año'},
            category_orders={"Período": ultimos_6_ing},
            barmode='stack',
            color_discrete_sequence=paleta_azul
        )
        
        fig_ing_evol.update_layout(
            separators=",.", 
            margin=dict(l=10, r=10, t=30, b=100),
            height=500,
            xaxis_tickangle=0, 
            font=dict(size=14, family="Source Sans Pro, sans-serif"),
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.4,
                xanchor="center",
                x=0.5,
                title_text=None,
                font=dict(size=14, color='black'),
                itemsizing='constant'
            ),
            xaxis=dict(
                tickfont=dict(size=14, color='black'),
                title=None,
                showline=True,
                linecolor='#E0E0E0'
            ),
            yaxis=dict(
                tickfont=dict(size=14, color='black'),
                tickformat="$,.0f", 
                title=None,
                gridcolor='rgba(200, 200, 200, 0.2)',
                zeroline=False
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )

        fig_ing_evol.update_traces(
            marker_line_color='white', 
            marker_line_width=1.5,
            hovertemplate="<b>%{x}</b><br>%{fullData.name}<br>Monto: $ %{y:,.0f}<extra></extra>",
            texttemplate='<b>$ %{y:,.0f}</b>', 
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=13, color='white'),
            cliponaxis=True 
        )

        # --- APLICACIÓN DEL CONTENEDOR ESTILIZADO (PORTAL) ---
        with st.container(border=True):
            # Subtítulo interno
            st.markdown(
                """
                <div style="margin-bottom: 5px;">
                    <p style="color: #6B7280; font-size: 0.99rem; font-weight: bold; margin-bottom: 8px;">
                        Análisis comparativo de ingresos: Últimos 6 períodos
                    </p>
                    <hr style="margin: 0; border: none; border-top: 1px solid #F3F4F6;">
                </div>
                """, 
                unsafe_allow_html=True
            )
            
            # Gráfico dentro del contenedor
            st.plotly_chart(fig_ing_evol, use_container_width=True, config={'displayModeBar': False})

            # NOTA INFORMATIVA CON ESTILO AZUL CLARO
            st.info("ℹ️ **Nota:** Cifras basadas en recaudación efectiva mensual (no incluye saldos de arrastre).")

    except Exception as e:
        st.info("No hay suficientes datos históricos de ingresos para mostrar la tendencia.")


# --- PESTAÑA EGRESOS ---    
    with tab_egr:
        st.write(f"### 📉 Análisis de Egresos - {mes_seleccionado}")
# ---
# --- NUEVA SECCIÓN: EQUILIBRIO OPERATIVO (ANTES DE GRÁFICOS) ---
        with st.container(border=True):
            # 1. Obtenemos la Recaudación Real (usando df_ing_mes que ya tienes cargado)
            # Filtramos por "Gastos Comunes" con la lógica robusta de una sola palabra
            df_ggcc_rec = df_ing_mes[
                df_ing_mes["Tipo de Ingreso"].astype(str).str.strip().str.upper() == "GASTOS COMUNES"
            ].copy()

            # --- MODIFICACIÓN: FILTRADO DE EXCLUSIONES PARA RECAUDACIÓN REAL ---
            excluidos_egr = calcular_excluidos_del_mes(mes_seleccionado, DICCIONARIO_EXCLUSIONES)
            # Solo consideramos la recaudación de quienes debían pagar
            df_ggcc_rec_validos = df_ggcc_rec[~df_ggcc_rec['Unidad'].astype(str).str.replace(".0","").isin([str(e).replace(".0","") for e in excluidos_egr])]
            
            recaudacion_real = df_ggcc_rec_validos["Pago Total del Mes"].sum()
            # -----------------------------------------------------------------
            
            # 2. Gasto Total del Mes (de la pestaña actual)
            gasto_total = df_egr_mes["Gasto Mensual"].sum()
            
            # 3. Cálculos de Resumen Financiero
            balance_mes = recaudacion_real - gasto_total
            cobertura = (recaudacion_real / gasto_total * 100) if gasto_total > 0 else 0
            
            st.write("#### ⚖️ Equilibrio Mensual (Recaudación vs. Gastos)")
            
            # Nota informativa si hay exclusiones que afecten el equilibrio
            if excluidos_egr:
                st.caption(f"ℹ️ El análisis de equilibrio considera la recaudación de unidades activas ({len(df_ggcc_rec_validos)} unidades).")

            c1, c2, c3, c4 = st.columns(4)
            
            # --- MODIFICACIÓN: LÓGICA UNIFICADA DE EQUILIBRIO ---
            # 1. Calculamos el valor absoluto para la métrica
            diferencia_abs = abs(balance_mes)
            
            # 2. Definimos etiquetas y colores (Superávit/Déficit es equivalente a A favor/En contra)
            if balance_mes >= 0:
                etiqueta_balance = "Superávit (A favor)"
                color_balance = "normal"  # Verde
            else:
                etiqueta_balance = "Déficit (En contra)"
                color_balance = "inverse" # Rojo

            c1.metric("Recaudado (GGCC)", f"$ {formato_clp(recaudacion_real)}")
            c2.metric("Gastado Total", f"$ {formato_clp(gasto_total)}")
            
            # Columna 3: Diferencia limpia de signos negativos
            c3.metric(
                "Diferencia", 
                f"$ {formato_clp(diferencia_abs)}", 
                delta=etiqueta_balance,
                delta_color=color_balance
            )
            
            # 4. Estado de Cobertura (Unificado con el lenguaje del Portal)
            if cobertura >= 100:
                estado_eco = "✅ SANO"
            elif cobertura >= 85:
                estado_eco = "⚠️ AJUSTADO"
            else:
                estado_eco = "🚨 ALERTA"
                
            c4.metric("Estado", estado_eco, delta=f"{cobertura:.1f}%")
            # --- FIN MODIFICACIÓN: LÓGICA UNIFICADA DE EQUILIBRIO ---

        st.write("") # Espacio estético entre el análisis y los gráficos

# ---
        
        # 1. Gráfico del Mes Actual
        col_e1, col_e2 = st.columns(2)
        
        with col_e1:
            fig_pie_egr = px.pie(
                df_egr_mes, 
                values='Gasto Mensual', 
                names='Clasificación', 
                hole=0.4,
                title="Distribución del Mes"
            )
            fig_pie_egr.update_layout(
                separators=",.", # Estandarización de miles (punto) y decimales (coma)
                margin=dict(l=0, r=0, t=40, b=0), 
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5),
                # CONFIGURACIÓN DE FUENTE
                title_font=dict(size=17, family="Source Sans Pro, sans-serif") 
            )
            # Hovertemplate para que muestre el monto con puntos en miles
            fig_pie_egr.update_traces(
                textposition='inside', 
                textinfo='percent+label',
                hovertemplate="<b>%{label}</b><br>Gasto: $ %{value:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_pie_egr, use_container_width=True, config={'displayModeBar': False})

        with col_e2:
            fig_bar_egr = px.bar(
                df_egr_mes, 
                x='Clasificación', 
                y='Gasto Mensual', 
                color='Tipo de Egreso',
                title="Detalle por Categoría"
            )
            fig_bar_egr.update_layout(
                separators=",.", # Estandarización de miles (punto) y decimales (coma)
                margin=dict(l=0, r=0, t=40, b=0),
                height=400,
                xaxis_title=None,
                yaxis_title=None,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                showlegend=True,
                # AGREGAMOS ESTO PARA IGUALAR EL TAMAÑO
                title_font=dict(size=17, family="Source Sans Pro, sans-serif") 
            )
            # Hovertemplate para que las barras también muestren el formato correcto
            fig_bar_egr.update_traces(
                hovertemplate="<b>%{x}</b><br>%{fullData.name}<br>Monto: $ %{y:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_bar_egr, use_container_width=True, config={'displayModeBar': False})

        # --- SEPARADOR ELEGANTE ---
        st.markdown("---")

# --- Análisis de Concentración
# --- LÓGICA PARA EL TOP 3 DE GASTOS (PARETO) ---
        try:
            # 1. Agrupamos y ordenamos de mayor a menor
            df_top = df_egr_mes.groupby('Tipo de Egreso')['Gasto Mensual'].sum().reset_index()
            df_top = df_top.sort_values(by='Gasto Mensual', ascending=False)
            
            # 2. Calculamos el total y los porcentajes
            total_egr_mes = df_top['Gasto Mensual'].sum()
            df_top['Porcentaje'] = (df_top['Gasto Mensual'] / total_egr_mes) * 100
            
            # 3. Tomamos los 3 primeros
            top_3 = df_top.head(3)
            suma_top_3_pct = top_3['Porcentaje'].sum()

            # 4. Diseño del mensaje
            st.write("🔍 **Análisis de Concentración de Gastos:**")
            
            # Usamos columnas pequeñas para un look de "etiquetas" profesionales
            cols_top = st.columns(3)
            for i, (index, row) in enumerate(top_3.iterrows()):
                with cols_top[i]:
                    st.caption(f"**{i+1}°. {row['Tipo de Egreso']}**")
                    st.write(f" {row['Porcentaje']:.1f}% del total")
            
            # Mensaje de conclusión tipo Pareto
            if suma_top_3_pct > 60:
                st.info(f"💡 Estas 3 categorías concentran el **{suma_top_3_pct:.1f}%** de los gastos este mes.")
            else:
                st.write(f"Tu gasto está distribuido: el Top 3 representa el {suma_top_3_pct:.1f}%.")

        except Exception as e:
            st.write("No hay datos suficientes para el análisis Top 3.")

        # --- SEPARADOR ELEGANTE ---
        st.markdown("---")

# --- BLOQUE DE BÚSQUEDA PLEGABLE ESPECÍFICA DE EGRESOS---
        # Título de la sección
        st.write("### 🟥 **Detalle y Búsqueda de Movimientos (Gastos)**")

        # El expander permite que el buscador no robe espacio visual innecesario
        with st.expander("🔍 Buscar un gasto específico (Proveedor, Tipo o Clasificación)"):
            busqueda_egr = st.text_input(
                "Filtre los movimientos del mes:", 
                placeholder="Ej: Chilquinta, Sueldos, Reparación...",
                key="search_egresos_input"
            ).strip()

        # 1. Lógica de Filtrado Reactivo
        if busqueda_egr:
            # Buscamos coincidencias en las columnas clave
            mask = (
                df_egr_mes['Tipo de Egreso'].astype(str).str.contains(busqueda_egr, case=False) |
                df_egr_mes['Clasificación'].astype(str).str.contains(busqueda_egr, case=False)
            )
            df_egr_filtrado = df_egr_mes[mask]
        else:
            df_egr_filtrado = df_egr_mes.copy()

        # 2. Preparación de Visualización (Copia del filtrado)
        df_egr_display = df_egr_filtrado.copy()
        
        # Formateo visual 🔴
        df_egr_display['Gasto Mensual'] = df_egr_display['Gasto Mensual'].apply(
            lambda x: f"🔴 $ {formato_clp(x)}"
        )

        # Ajuste de strings para alineación
        for col in ['Año', 'Período']:
            if col in df_egr_display.columns:
                df_egr_display[col] = df_egr_display[col].astype(str)

        # 3. Renderizado de la Tabla
        st.dataframe(
            df_egr_display,
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Gasto Mensual": st.column_config.Column(width="medium"),
                "Tipo de Egreso": st.column_config.Column(width="medium"),
                "Clasificación": st.column_config.Column(width="medium")
            }
        )

        # 4. Total Dinámico (Suma solo lo filtrado)
        total_egr_final = df_egr_filtrado['Gasto Mensual'].sum()
        
        col_te1, col_te2 = st.columns([3, 1])
        with col_te2:
            st.markdown(f"### **Total: 🔴 $ {formato_clp(total_egr_final)}**")

        # --- SEPARADOR ELEGANTE ---
    #     st.markdown("---")

#---- FIN BLOQUE DE BÚSQUEDA PLEGABLE
        
        # 3. LÓGICA DE PROYECCIÓN / EVOLUCIÓN HISTÓRICA
        st.write("### 📊 Evolución de Gastos (Últimos 6 Períodos)")
 
        # 🟢 CONFIGURACIÓN DE PALETA
        paleta = ["#023e8a", "#0077b6", "#0096c7", "#00b4d8", "#48cae4", "#90e0ef"] 

        # Agrupamos todos los egresos por Período y Tipo de Egreso (Cambiado de Clasificación)
        df_egr_hist = df_egr.groupby(['Período', 'Tipo de Egreso'])['Gasto Mensual'].sum().reset_index()
        
        # Ordenamiento cronológico
        try:
            periodos_ordenados = sorted(df_egr['Período'].unique(), key=lambda x: datetime.strptime(str(x), '%m/%Y'))
        except Exception:
            periodos_ordenados = sorted(df_egr['Período'].unique())
            
        ultimos_6 = periodos_ordenados[-6:] 
        df_plot_hist = df_egr_hist[df_egr_hist['Período'].isin(ultimos_6)]

        # --- CAMBIO A GRÁFICO DE BARRAS APILADAS POR TIPO DE EGRESO ---
        fig_evolucion = px.bar(
            df_plot_hist, 
            x='Período', 
            y='Gasto Mensual', 
            color='Tipo de Egreso', # Cambiado para agrupar por categoría macro
            title=None, 
            labels={'Gasto Mensual': 'Monto ($)', 'Período': 'Mes/Año'},
            category_orders={"Período": ultimos_6},
            color_discrete_sequence=paleta 
        )
        
        # Aplicamos la estética unificada
        fig_evolucion.update_layout(
            separators=",.", 
            margin=dict(l=10, r=10, t=30, b=100), # Ajustado para el contenedor
            height=550, 
            xaxis_tickangle=0,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(size=14, family="Source Sans Pro, sans-serif"),
            
            showlegend=True,
            legend=dict(
                orientation="h", 
                yanchor="top", 
                y=-0.25, 
                xanchor="center", 
                x=0.5,
                title_text=None,
                font=dict(size=14, color='black'), 
                itemsizing='constant'
            ),
            
            xaxis=dict(
                tickfont=dict(size=14, color='black'),
                title=None,
                showline=True,
                linecolor='#E0E0E0'
            ),
            yaxis=dict(
                tickfont=dict(size=14, color='black'),
                tickformat="$,.0f", 
                title=None,
                gridcolor='rgba(200, 200, 200, 0.2)',
                zeroline=False
            )
        )
        
        fig_evolucion.update_traces(
            marker_line_color='white', 
            marker_line_width=1.5,
            texttemplate='<b>$ %{y:,.0f}</b>',
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=12, color='white'),
            cliponaxis=True,
            hovertemplate="<b>%{x}</b><br>%{fullData.name}<br>Monto: $ %{y:,.0f}<extra></extra>"
        )
         
        # --- APLICACIÓN DEL CONTENEDOR ESTILIZADO (IGUAL QUE EN INGRESOS) ---
        with st.container(border=True):
            # Subtítulo interno descriptivo (Actualizado a 'Tipo de Egreso')
            st.markdown(
                """
                <div style="margin-bottom: 5px;">
                    <p style="color: #6B7280; font-size: 0.99rem; font-weight: bold; margin-bottom: 8px;">
                        Análisis comparativo de egresos: Distribución por tipo de gasto histórico
                    </p>
                    <hr style="margin: 0; border: none; border-top: 1px solid #F3F4F6;">
                </div>
                """, 
                unsafe_allow_html=True
            )

            # Renderizado del gráfico
            st.plotly_chart(fig_evolucion, use_container_width=True, config={'displayModeBar': False})

            # NOTA INFORMATIVA CON ESTILO AZUL CLARO
            st.info("ℹ️ **Nota:** Este gráfico consolida los gastos por categoría principal (Tipo de Egreso), permitiendo identificar las áreas de mayor impacto en el presupuesto mensual.")
  
# --- PESTAÑA CONTROL MOROSIDAD ---     
with tab_mor:
    st.write(f"### 🚦 Control de Morosidad - {mes_seleccionado}")
    
    if UNIDADES_EXCLUIDAS:
        st.info(f"ℹ️ Nota: La unidad {', '.join(UNIDADES_EXCLUIDAS)} está excluida por acuerdos administrativos.")

    # --- LÓGICA DE DATOS ---
    # 1. Definimos el universo de departamentos
    universo = df_ing[df_ing['Tipo de Unidad'].str.contains('Departamento', case=False, na=False)][['Unidad', 'Residente']].drop_duplicates('Unidad')
    universo_efectivo = universo[~universo['Unidad'].isin(UNIDADES_EXCLUIDAS)]
    
    # --- [CORRECCIÓN DE LÓGICA PARA EL MES ACTUAL] ---
    # Copiamos la tabla de control y limpiamos los nombres de las columnas
    df_h = df_control.copy()
    df_h.columns = [str(c).strip() for c in df_h.columns]
    col_id = "Departamento N°"

    unidades_pendientes_mes = []

    # Recorremos el universo efectivo buscando quién NO tiene un "1" en la columna del mes seleccionado
    if mes_seleccionado in df_h.columns:
        for _, fila in universo_efectivo.iterrows():
            u = str(fila['Unidad']).replace(".0", "")
            
            # Buscamos la fila de la unidad en la planilla de control
            fila_control = df_h[df_h[col_id].astype(str).str.replace(".0", "") == u]
            
            if not fila_control.empty:
                val_raw = fila_control[mes_seleccionado].iloc[0]
                val = str(val_raw).strip().lower() if pd.notna(val_raw) else ""
                
                # Si la celda está vacía o no es "1"/"1.0", significa que está PENDIENTE
                if val not in ["1", "1.0"]:
                    unidades_pendientes_mes.append(u)

    # Filtramos nuestro universo efectivo para quedarnos SOLO con los que se detectaron pendientes
    df_morosos_mes = universo_efectivo[universo_efectivo['Unidad'].astype(str).str.replace(".0", "").isin(unidades_pendientes_mes)].copy()


    # --- VENTANA 1: MOROSIDAD DEL MES ACTUAL ---
    st.subheader(f"📅 Pendientes del Mes: {mes_seleccionado}")
    
    if not df_morosos_mes.empty:
        # 1. Traemos el monto de la tabla de ingresos del mes
        df_morosos_mes = df_morosos_mes.merge(
            df_ing_mes[['Unidad', 'Monto Actual']].drop_duplicates('Unidad'), 
            on='Unidad', 
            how='left'
        )
        
        # 2. LÓGICA DE RESPALDO: Si el monto es 0 o NaN (común en meses nuevos), busca en el histórico
        def obtener_monto_seguro(fila):
            m_actual = pd.to_numeric(fila['Monto Actual'], errors='coerce')
            if pd.isna(m_actual) or m_actual == 0:
                # Buscamos en el DataFrame general de ingresos (df_ing)
                respaldo = df_ing[df_ing['Unidad'] == fila['Unidad']]['Monto Actual']
                if not respaldo.empty:
                    return pd.to_numeric(respaldo.iloc[0], errors='coerce')
            return m_actual

        df_morosos_mes['Monto Actual'] = df_morosos_mes.apply(obtener_monto_seguro, axis=1).fillna(0)
        
        df_morosos_mes['Estado'] = "🔴 Pendiente"
        
        # 3. Preparación para visualización
        df_display_mes = df_morosos_mes[['Unidad', 'Residente', 'Monto Actual', 'Estado']].copy()
        
        # 4. Formateo seguro
        df_display_mes['Monto Actual'] = df_display_mes['Monto Actual'].apply(lambda x: f"$ {formato_clp(x)}")
        
        # Mostramos la tabla
        st.dataframe(
            df_display_mes.sort_values('Unidad'), 
            use_container_width=True, 
            hide_index=True
        )
        
        # 5. Cálculo del Total
        monto_total_mes = df_morosos_mes['Monto Actual'].sum()
        st.metric("Total por Recaudar (Mes)", f"$ {formato_clp(monto_total_mes)}")
      
    else:
        st.success("✅ ¡Excelente! No hay pendientes para el mes actual.")

    st.divider()

    # --- VENTANA 2: MOROSIDAD HISTÓRICA ACUMULADA ---
    st.subheader("🕵️ Análisis de Deuda Histórica (Saldos Anteriores)")
    
    # 1. PREPARAR COLUMNAS DE MESES DISPONIBLES
    df_h = df_control.copy()
    df_h.columns = [str(c).strip() for c in df_h.columns]
    cols_meses_total = [c for c in df_h.columns if ("/" in str(c) or "-" in str(c))]
    
    # --- [NUEVA LÓGICA DE SINCRONIZACIÓN] ---
    # Buscamos el índice del mes que seleccionaste en la sidebar (mes_seleccionado)
    # dentro de la lista de meses de la tabla de control (cols_meses_total)
    try:
        # Intentamos encontrar el índice exacto
        indice_defecto = cols_meses_total.index(mes_seleccionado)
    except ValueError:
        # Si por alguna razón el formato no coincide, volvemos al último disponible
        indice_defecto = len(cols_meses_total) - 1 if len(cols_meses_total) > 0 else 0
    
    # 2. INTERFAZ DE FILTROS
    col_f1, col_f2 = st.columns([1, 1])
    with col_f1:
        mes_corte = st.selectbox(
            "Considerar deuda hasta el mes (inclusive):",
            cols_meses_total,
            index=indice_defecto, # <-- AHORA USA EL MES DE LA SIDEBAR POR DEFECTO
            help="Los meses posteriores a este no se sumarán a la deuda histórica."
        )

    with col_f2:
        criterio_orden = st.radio(
            "Ordenar lista por:",
            ["N° Departamento", "Monto de Deuda"],
            horizontal=True
        )

    try:
        col_id = "Departamento N°"
        idx_corte = cols_meses_total.index(mes_corte)
        cols_a_evaluar = cols_meses_total[:idx_corte + 1]
        
        st.caption(f"🔎 Evaluando deuda en los periodos: {', '.join(cols_a_evaluar)}")

        resumen_historico = []

        for _, fila in universo_efectivo.iterrows():
            u = str(fila['Unidad']).replace(".0", "") # Limpieza de unidad
            res = fila['Residente']
            fila_control = df_h[df_h[col_id].astype(str).str.replace(".0", "") == u]
            
            if not fila_control.empty:
                meses_pendientes = []
                for col in cols_a_evaluar:
                    # --- NUEVA REGLA DE PROTECCIÓN HISTÓRICA ---
                    # Si la unidad está en la lista de excluidos PARA ESE MES específico, saltamos
                    excluidos_ese_mes = calcular_excluidos_del_mes(col, DICCIONARIO_EXCLUSIONES)
                    if u in excluidos_ese_mes:
                        continue # No sumamos este mes a la deuda
                    
                    # Lógica original de detección de pago (1 o vacío)
                    val_raw = fila_control[col].iloc[0]
                    val = str(val_raw).strip().lower() if pd.notna(val_raw) else ""
                    if val not in ["1", "1.0"]:
                        meses_pendientes.append(col)
                
                # Solo agregamos al resumen si después del filtro de exclusión quedaron meses pendientes
                if meses_pendientes:
                    # --- SOLUCIÓN PLAN B: BUSCAR MONTO EN UNIVERSO ---
                    monto_u_query = df_ing_mes[df_ing_mes['Unidad'].astype(str).str.replace(".0", "") == u]['Monto Actual']
                    monto_u = pd.to_numeric(monto_u_query.iloc[0], errors='coerce') if not monto_u_query.empty else 0
                    
                    if pd.isna(monto_u) or monto_u == 0:
                        respaldo = df_ing[df_ing['Unidad'].astype(str).str.replace(".0", "") == u]['Monto Actual']
                        monto_u = pd.to_numeric(respaldo.iloc[0], errors='coerce') if not respaldo.empty else 0
                    
                    deuda_est = len(meses_pendientes) * monto_u
                    
                    resumen_historico.append({
                        "Unidad": u, "Residente": res, "Meses Deuda": len(meses_pendientes),
                        "Detalle Meses": ", ".join(meses_pendientes), "Deuda Total Est.": deuda_est
                    })

        if resumen_historico:
            df_hist_final = pd.DataFrame(resumen_historico)
            df_hist_final['Unidad_Num'] = pd.to_numeric(df_hist_final['Unidad'], errors='coerce')
            
            def color_morosidad(cant):
                return "🔴 Crítico" if cant >= 3 else "🟡 Riesgo"
            
            df_hist_final['Gravedad'] = df_hist_final['Meses Deuda'].apply(color_morosidad)
            
            if "Monto de Deuda" in criterio_orden:
                df_view_hist = df_hist_final.sort_values('Deuda Total Est.', ascending=False)
            else:
                df_view_hist = df_hist_final.sort_values('Unidad_Num', ascending=True)
            
            df_display = df_view_hist.copy()
            df_display['Deuda Total Est.'] = df_display['Deuda Total Est.'].apply(lambda x: f"$ {formato_clp(x)}")
            df_display['Unidad'] = df_display['Unidad'].astype(str)
            df_display['Meses Deuda'] = df_display['Meses Deuda'].astype(str) 
            
            st.dataframe(
                df_display[['Unidad', 'Residente', 'Meses Deuda', 'Deuda Total Est.', 'Gravedad', 'Detalle Meses']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Unidad": st.column_config.Column(width="small"),
                    "Meses Deuda": st.column_config.Column("Meses", width="small"),
                    "Deuda Total Est.": st.column_config.Column(width="medium"),
                    "Gravedad": st.column_config.Column(width="small")
                }
            )

 # --- Inicio exportación a Excel
            # Botón de Descarga
            # --- VERSIÓN: MOTOR NATIVO CALIBRI 11 + AUTOAJUSTE REAL ---
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. CREAMOS EL DATAFRAME FILTRADO (Sin Unidad_Num)
                cols_finales = [c for c in df_view_hist.columns if c != 'Unidad_Num']
                df_excel = df_view_hist[cols_finales].copy()
                
                # Exportamos la base de datos al excel
                df_excel.to_excel(writer, index=False, sheet_name='Morosidad', startrow=2)

                workbook  = writer.book
                worksheet = writer.sheets['Morosidad']
                
                # ESTILOS
                format_titulo = workbook.add_format({'bold': True, 'size': 14, 'align': 'center'})
                format_header = workbook.add_format({
                    'bold': True, 'bg_color': '#1f4e78', 'font_color': 'white',
                    'border': 1, 'align': 'center', 'valign': 'vcenter'
                })
                format_base = workbook.add_format({'border': 1, 'valign': 'vcenter', 'size': 11})
                format_centrado = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'size': 11})
                format_moneda = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'right', 'size': 11})
                format_critico = workbook.add_format({'font_color': '#FF0000', 'bold': True, 'border': 1, 'size': 11})
                format_riesgo = workbook.add_format({'font_color': '#FF8C00', 'bold': True, 'border': 1, 'size': 11})

                # 2. TÍTULO (Usamos df_excel.columns para el ancho)
                last_col = len(df_excel.columns) - 1
                worksheet.merge_range(0, 0, 0, last_col, f"REPORTE DE MOROSIDAD HISTÓRICA - CORTE {mes_corte}", format_titulo)

                # SUBTÍTULO
                from datetime import datetime
                fecha_extraccion = datetime.now().strftime("%d/%m/%Y %H:%M")
                format_subtitulo = workbook.add_format({'size': 9, 'italic': True, 'align': 'right', 'valign': 'vcenter'})
                worksheet.merge_range(1, 0, 1, last_col, f"Fecha del reporte: {fecha_extraccion}", format_subtitulo)
                
                # 3. ESCRITURA DE ENCABEZADOS (Usando df_excel)
                for col_num, value in enumerate(df_excel.columns.values):
                    worksheet.write(2, col_num, value, format_header)

                # 4. ESCRITURA DE FILAS (Usando df_excel)
                for row_num in range(len(df_excel)):
                    for col_num, col_name in enumerate(df_excel.columns):
                        val = df_excel.iloc[row_num, col_num]
                        
                        if col_name in ["Unidad", "Meses Deuda"]:
                            worksheet.write(row_num + 3, col_num, val, format_centrado)
                        elif col_name == "Gravedad":
                            fmt = format_critico if "Crítico" in str(val) else format_riesgo
                            worksheet.write(row_num + 3, col_num, val, fmt)
                        elif "Deuda Total" in col_name:
                            # --- SOLUCIÓN DEFINITIVA DESDE EL DATAFRAME ORIGINAL ---
                            try:
                                # Buscamos la fila correspondiente en el DataFrame original antes del formateo visual
                                # Usamos la unidad actual como llave de búsqueda
                                unidad_actual = df_excel.iloc[row_num]['Unidad']
                                
                                # Extraemos el valor numérico crudo desde df_view_hist
                                fila_original = df_view_hist[df_view_hist['Unidad'] == unidad_actual]
                                
                                if not fila_original.empty:
                                    # Accedemos al valor original (que en Python es un float o int puro)
                                    val_original = fila_original[col_name].iloc[0]
                                    
                                    # Si por casualidad ya era un string con formato en df_view_hist, hacemos una limpieza agresiva:
                                    if isinstance(val_original, str):
                                        # Eliminamos todo lo que no sea un dígito numérico
                                        import re
                                        solo_numeros = re.sub(r'[^\d]', '', val_original)
                                        n_val = int(solo_numeros) if solo_numeros else 0
                                    else:
                                        # Si es un número puro, lo convertimos a entero de forma segura
                                        n_val = int(float(val_original))
                                else:
                                    # Si no lo encuentra, limpiamos el string actual removiendo el punto final .0 si existiera
                                    s_val = str(val).split('.')[0] # Corta cualquier decimal .0
                                    import re
                                    solo_numeros = re.sub(r'[^\d]', '', s_val)
                                    n_val = int(solo_numeros) if solo_numeros else 0

                                # Escribimos el número limpio en Excel
                                worksheet.write_number(row_num + 3, col_num, n_val, format_moneda)
                                
                            except Exception as e:
                                # Respaldo absoluto: si todo lo anterior falla, limpia el string removiendo el $ y puntos
                                try:
                                    s_val = str(val).split('.')[0] # Evita arrastrar decimales ocultos
                                    import re
                                    solo_numeros = re.sub(r'[^\d]', '', s_val)
                                    worksheet.write_number(row_num + 3, col_num, int(solo_numeros), format_moneda)
                                except:
                                    worksheet.write(row_num + 3, col_num, val, format_moneda)
                        else:
                            worksheet.write(row_num + 3, col_num, str(val), format_base)

                # 5. AUTOAJUSTE (Usando df_excel)
                for i, col in enumerate(df_excel.columns):
                    max_len = max(df_excel[col].astype(str).map(len).max(), len(str(col)))
                    worksheet.set_column(i, i, max_len + 2)

                # 6. CONFIGURACIÓN DE PÁGINA ---
                worksheet.set_paper(1)                # Tamaño Carta
                worksheet.set_landscape()             # Horizontal
                worksheet.fit_to_pages(1, 1)          # Forzar a 1 página
                worksheet.center_horizontally()       # Centrar en la hoja
                worksheet.hide_gridlines(2)           # Eliminar cuadrícula
                worksheet.set_margins(0.3, 0.3, 0.5, 0.5) # Márgenes estrechos

           
            total_historico = df_hist_final['Deuda Total Est.'].sum()
            st.metric(f"Deuda Total Acumulada (al {mes_corte})", f"$ {formato_clp(total_historico)}")
            

# --- SECCIÓN DE EXPORTACIÓN (Reemplaza tu st.download_button antiguo) ---
            st.markdown("---")
            st.markdown("""
                <div style="background-color: #fdfefe; padding: 15px; border-radius: 10px; border: 1px solid #e1e4e8; border-left: 5px solid #1f4e78; margin-bottom: 20px;">
                    <h4 style="margin-top: 0; color: #1f4e78;">📥 Descargar Reportes de Morosidad</h4>
                    <p style="font-size: 0.9rem; color: #444;">
                        Los archivos contienen el listado detallado de unidades con deuda, ordenados según su preferencia actual y filtrados hasta el mes de corte <b>( {0} )</b>.
                    </p>
                </div>
            """.format(mes_corte), unsafe_allow_html=True)

            col_exp1, col_exp2 = st.columns(2)

            with col_exp1:
                # Botón de Excel (Mantiene tu lógica actual)
                st.download_button(
                    label="❎ Descargar Reporte (Excel)",
                    data=output.getvalue(),
                    file_name=f"Reporte_Mora_Edificio_{mes_corte.replace('/','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_exp2:
                # Botón de PDF (Nuevo)
                try:
                    pdf_mora_bytes = bytes(generate_pdf_morosidad_historica(df_view_hist, mes_corte))
                    st.download_button(
                        label="📕 Descargar Reporte (PDF)",
                        data=pdf_mora_bytes,
                        file_name=f"Reporte_Mora_Edificio_{mes_corte.replace('/','_')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error al generar PDF: {e}")

 # --- Fin exportación a Excel y PDF
            
        else:
            st.info(f"✅ No se detectó deuda histórica hasta {mes_corte}.")

    except Exception as e:
        st.error(f"Error al procesar historial: {e}")

# --- PESTAÑA CONTROL PAGOS ---    
    with tab_control:
        st.write("### 🗓️ Matriz de Pagos por Departamento")

        # 1. Agregamos nota con tono amarillo para mantener la identidad visual
        st.warning(
            f"**Recordatorio:** Este historial incluye todos los abonos registrados a la fecha, "
            f"incluyendo pagos anticipados realizados para meses posteriores a **{mes_seleccionado}**.",
            icon="⚠️"
        )

        # 2. Nota de Exclusión (Azul) - Ajustada para coincidir en formato
        if UNIDADES_EXCLUIDAS:
            unidades_str = ", ".join(UNIDADES_EXCLUIDAS)
            st.info(
                f"**Exclusión:** La unidad {unidades_str} no se considera en la validación de morosidad.",
                icon="ℹ️"
            )
        
        if df_control is not None:
            df_visual = df_control.copy()
            col_unid_original = df_visual.columns[0]
            
            # Renombramos para el efecto de 2 pisos en la web
            df_visual = df_visual.rename(columns={col_unid_original: "Departamento N°"})
            col_unid = "Departamento N°" 
            
            # --- OPCIÓN DE VISTA COMPACTA ---
            # Colocamos el checkbox justo antes de la tabla
            vista_compacta = st.checkbox("🔍 Ver solo los últimos 16 meses (Vista Compacta)", value=True)
            
            # Identificamos las columnas que son meses
            cols_meses_total = [c for c in df_visual.columns if c != col_unid]
            
            if vista_compacta:
                # Limitamos a los últimos 16 meses
                cols_meses = cols_meses_total[-16:]
                df_visual = df_visual[[col_unid] + list(cols_meses)]
            else:
                # Mostramos todo el historial
                cols_meses = cols_meses_total
            
            df_visual[col_unid] = df_visual[col_unid].astype(str).str.strip()

            # --- CSS MAESTRO (Tu diseño actual) ---
            st.markdown("""
                <style>
                    .stTable { display: block !important; overflow-x: auto !important; }
                    .stTable th:first-child, .stTable td:first-child { display: none !important; }
                    .stTable th {
                        background-color: #2F5597 !important;
                        color: white !important;
                        text-align: center !important;
                        font-weight: bold !important;
                        padding: 8px 10px !important;
                        white-space: nowrap !important;
                    }
                    .stTable td {
                        text-align: center !important;
                        vertical-align: middle !important;
                        padding: 8px !important;
                        background-color: white;
                        white-space: nowrap !important;
                    }
                    .stTable td:nth-child(2), .stTable th:nth-child(2) {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 3 !important;
                        width: 110px !important; 
                        min-width: 120px !important;
                        max-width: 120px !important;
                        border-right: 2px solid #2F5597 !important;
                    }
                    .stTable th:nth-child(2) {
                        white-space: normal !important; 
                        word-wrap: break-word !important;
                        line-height: 1.1 !important;
                        background-color: #2F5597 !important;
                        z-index: 4 !important;
                    }
                    .stTable td:nth-child(2) { background-color: #f8f9fa !important; color: #000 !important; font-weight: bold !important; }
                </style>
            """, unsafe_allow_html=True)

            # Renderizado de la tabla
            # --- LÓGICA DE FORMATEO ROBUSTA ---
            def formatear_celda(val):
                v_str = str(val).strip().lower()
                if v_str in ["1", "1.0"]:
                    return "✅"
                elif v_str in ["0", "0.0", "nan", "none", ""]:
                    return ""
                else:
                    # Si escribieron cualquier otra cosa (letras, otros números, etc.)
                    return f"❓ {val}" 

            def estilizar_celda(val):
                v_str = str(val).strip().lower()
                if v_str in ["1", "1.0"]:
                    return 'background-color: #d4edda; color: #155724; font-weight: bold;'
                elif v_str in ["0", "0.0", "nan", "none", ""]:
                    return ''
                else:
                    # Color de advertencia para datos inesperados
                    return 'background-color: #fff3cd; color: #856404; font-weight: bold;'

            # Renderizado de la tabla con la nueva lógica
            st.table(
                df_visual.style
                .applymap(estilizar_celda, subset=cols_meses)
                .format(formatear_celda, subset=cols_meses)
            )

            # --- LÓGICA DE EXPORTACIÓN A EXCEL ---
            st.write("---") 
            import io
            import xlsxwriter # <--- IMPORTANTE: Aseguramos la importación aquí
            from datetime import datetime

            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # --- AJUSTE: LIMITAR A ÚLTIMOS 16 MESES (Manteniendo columna de ID) ---
                # Identificamos columnas que son meses (tienen / o -)
                cols_meses_ex = [c for c in df_visual.columns if "/" in str(c) or "-" in str(c)]
                # Tomamos la columna de identificación (la primera) y los últimos 16 meses
                cols_a_exportar = [df_visual.columns[0]] + cols_meses_ex[-16:]
                
                df_export = df_visual[cols_a_exportar].copy()
                workbook  = writer.book
                worksheet = workbook.add_worksheet('Matriz_Pagos')

                # --- CONFIGURACIÓN DE PÁGINA Y CENTRADO ---
                worksheet.hide_gridlines(2)
                worksheet.set_paper(1)                # Tamaño Carta
                worksheet.set_landscape()             # Horizontal
                worksheet.fit_to_pages(1, 1)          # FORZAR A 1 PÁGINA
                worksheet.center_horizontally()       # CENTRAR EN LA HOJA
                worksheet.set_margins(0.2, 0.2, 0.5, 0.5)

                # --- 2. DEFINICIÓN DE FORMATOS ---
                fmt_title = workbook.add_format({
                    'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'font_color': '#2F5597'
                })
                fmt_header = workbook.add_format({
                    'bold': True, 'bg_color': '#2F5597', 'font_color': 'white',
                    'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
                })
                fmt_check = workbook.add_format({
                    'bg_color': '#d4edda', 'font_color': '#155724',
                    'align': 'center', 'valign': 'vcenter', 'bold': True, 'border': 1
                })
                fmt_normal = workbook.add_format({
                    'align': 'center', 'valign': 'vcenter', 'border': 1
                })

                # --- 3. INSERTAR TÍTULO ---
                last_col_index = len(df_export.columns) - 1
                titulo_texto = f"REPORTE DE CONTROL DE PAGOS - {datetime.now().strftime('%d/%m/%Y')}"
                worksheet.merge_range(0, 0, 0, last_col_index, titulo_texto, fmt_title)
                worksheet.set_row(0, 30)

                # --- 4. ESCRIBIR ENCABEZADOS ---
                for col_num, value in enumerate(df_export.columns.values):
                    clean_header = str(value).replace("<br>", " ")
                    worksheet.write(1, col_num, clean_header, fmt_header)
                    ancho = max(len(clean_header), 12)
                    worksheet.set_column(col_num, col_num, ancho)

                # --- 5. ESCRIBIR DATOS ---
                for row_num, row_data in enumerate(df_export.values):
                    for col_num, cell_value in enumerate(row_data):
                        display_value = cell_value
                        current_fmt = fmt_normal
                        
                        if col_num == 0:
                            try:
                                # Si es convertible a número, lo pasamos como tal
                                display_value = float(cell_value) if '.' in str(cell_value) else int(cell_value)
                            except:
                                display_value = str(cell_value)
                        
                        elif col_num > 0:
                            if str(cell_value).strip() in ["1", "1.0"] or cell_value == 1:
                                display_value = "✅"
                                current_fmt = fmt_check
                            elif cell_value == "" or pd.isna(cell_value) or str(cell_value).strip() in ["0", "0.0"]:
                                display_value = ""
                        
                        worksheet.write(row_num + 2, col_num, display_value, current_fmt)

                # --- 6. ELIMINAR ADVERTENCIAS (CON LA RUTA CORREGIDA) ---
                # Usamos xlsxwriter.utility para obtener la letra de la última columna
                letra_col = xlsxwriter.utility.xl_col_to_name(last_col_index)
                rango_ignorar = f'A1:{letra_col}{len(df_export)+2}'
                worksheet.ignore_errors({'number_stored_as_text': rango_ignorar})

            # --- SECCIÓN DE BOTONES DE EXPORTACIÓN ---            
            # Mensaje de contexto para el usuario
            st.markdown("""
                <div style="background-color: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 5px solid #2f5597; margin-bottom: 20px;">
                    <h4 style="margin-top: 0; color: #2f5597;">📥 Exportar Reporte de Pagos</h4>
                    <p style="font-size: 0.9rem; color: #444;">
                        Seleccione el formato deseado para descargar la <b>Matriz de Pagos</b>. 
                        Ambos archivos incluyen el resumen de los últimos 16 meses para los departamentos seleccionados.
                    </p>
                </div>
            """, unsafe_allow_html=True)

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                # Botón de Excel
                st.download_button(
                    label="❎ Descargar Planilla Excel",
                    data=output.getvalue(),
                    file_name=f"Matriz_de_Pagos_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_btn2:
                # Botón de PDF
                try:
                    pdf_bytes = bytes(generate_pdf_morosidad(df_export))
                    st.download_button(
                        label="📕 Descargar Reporte PDF",
                        data=pdf_bytes,
                        file_name=f"Matriz_de_Pagos_{datetime.now().strftime('%d-%m-%Y')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="btn_pdf_morosidad"
                    )
                except Exception as e:
                    st.error(f"No se pudo generar el PDF: {e}")

# --- PESTAÑA GENERADOR DE AVISOS ---     
    with tab_avisos:
        st.write("### ✉️ Centro de Cobranza Mensual")

        if UNIDADES_EXCLUIDAS:
            st.info(f"ℹ️ La unidad {', '.join(UNIDADES_EXCLUIDAS)} está excluida de la gestión de cobranza.")
        
        # 1. OBTENER LISTA DE UNIDADES DESDE EL UNIVERSO (No desde df_ing_mes)
        # Esto asegura que el selector SIEMPRE tenga todos los departamentos, aunque el mes esté vacío.
        unidades_totales = sorted([
            str(u).strip() for u in universo_efectivo['Unidad'].unique()
        ])
        
        u_seleccionada = st.selectbox(
            "Seleccione Departamento:", 
            unidades_totales, 
            key="selector_aviso_final_v6"
        )

        if u_seleccionada:
            # --- LÓGICA DE OBTENCIÓN DE DATOS (PLAN A y PLAN B) ---
            # Intentamos buscar en el mes actual (febrero, por ejemplo)
            datos_u = df_ing_mes[df_ing_mes['Unidad'].astype(str).str.strip() == str(u_seleccionada)].copy()
            mes_key = str(mes_seleccionado).replace("/", "_")
            
            # Inicializamos variables por defecto
            residente = ""
            monto_mes_actual = 0

            # Verificamos si hay datos en el mes actual y si el monto es válido
            monto_check = pd.to_numeric(datos_u['Monto Actual'].iloc[0], errors='coerce') if not datos_u.empty else 0

            if not datos_u.empty and monto_check > 0:
                # PLAN A: El mes tiene datos
                residente = datos_u['Residente'].iloc[0]
                monto_mes_actual = monto_check
            else:
                # PLAN B: El mes está vacío o en 0. Rescatamos del historial general
                # Buscamos en el df_ing completo (maestro histórico)
                hist_u = df_ing[(df_ing['Unidad'].astype(str).str.strip() == str(u_seleccionada)) & (df_ing['Monto Actual'] > 0)]
                
                if not hist_u.empty:
                    # Traemos el último registro válido
                    residente = hist_u['Residente'].iloc[-1]
                    monto_mes_actual = pd.to_numeric(hist_u['Monto Actual'].iloc[-1], errors='coerce')
                else:
                    # Si no hay historial, usamos el Universo Efectivo para el nombre
                    u_maestro = universo_efectivo[universo_efectivo['Unidad'].astype(str).str.strip() == str(u_seleccionada)]
                    residente = u_maestro['Residente'].iloc[0] if not u_maestro.empty else "Residente"
                    monto_mes_actual = 0

            # --- BLOQUE 1: RENDERIZADO DEL AVISO ---
            fecha_hoy = datetime.now().strftime('%d/%m/%Y')
            st.markdown(f"#### 📅 Cobro Mes Actual: {u_seleccionada}")
            
            # Construcción del mensaje para WhatsApp
            msj_mes = f"🏠 *AVISO DE COBRO - EDIFICIO CREATIVA II*\n"
            msj_mes += f"==========================================\n\n"
            msj_mes += f"📍 *Departamento:* {u_seleccionada}\n"
            msj_mes += f"👤 *Residente:* {residente}\n"
            msj_mes += f"📅 *Mes:* {mes_seleccionado}\n\n"
            msj_mes += f"💰 *TOTAL MES: $ {formato_clp(monto_mes_actual)}*\n"
            msj_mes += f"------------------------------------------\n"
            msj_mes += f"*DATOS PARA TRANSFERENCIA:*\n"
            msj_mes += f"👤 *Titular:* Edificio Creativa II\n"
            msj_mes += f"🆔 *RUT:* 65.696.770-6\n"
            msj_mes += f"🏦 *Banco:* Scotiabank\n"
            msj_mes += f"💳 *Tipo de Cuenta:* Cuenta Corriente\n"
            msj_mes += f"🔢 *N° de Cuenta:* 992591625\n"
            msj_mes += f"📧 *Comprobante:* leylajamett@gmail.com\n\n"
            msj_mes += f"📌 *Nota:* Favor enviar comprobante indicando unidad.\n"
            msj_mes += f"_Atentamente, Administración._"

            col1, col2 = st.columns([2, 1])
            with col1:
                st.text_area("Copia Aviso Mes:", value=msj_mes, height=250, key=f"txt_mes_{u_seleccionada}_{mes_key}")
            with col2:
                import urllib.parse
                link_mes = f"https://api.whatsapp.com/send?text={urllib.parse.quote_plus(msj_mes)}"
                st.markdown(f'<a href="{link_mes}" target="_blank" style="text-decoration:none;"><div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; text-align:center; font-weight:bold; cursor:pointer;">📲 Enviar Mes</div></a>', unsafe_allow_html=True)
                st.metric("Total Mes", f"$ {formato_clp(monto_mes_actual)}")


            # --- BLOQUE 2: DEUDA HISTÓRICA (CON MEMORIA Y ORDEN INVERSO) ---
            st.divider()
            st.write("### ⚠️ Estado de Deuda Histórica")

            try:
                df_h = df_control.copy()
                df_h.columns = [str(c).strip() for c in df_h.columns]
                col_id = "Departamento N°"
                
                u_busqueda = str(u_seleccionada).strip()
                fila_h = df_h[df_h[col_id].astype(str).str.strip() == u_busqueda]

                if not fila_h.empty:
                    # 1. Obtener meses y ORDENAR de mayor a menor (más reciente primero)
                    cols_meses_raw = [c for c in df_h.columns if "/" in str(c) or "-" in str(c)]
                    # Ordenamos cronológicamente inverso
                    cols_meses_raw.sort(key=lambda x: pd.to_datetime(x, errors='coerce'), reverse=True)
                    
                    # 2. LÓGICA DE PERSISTENCIA (Session State)
                    # Creamos una llave única para este selector en la memoria
                    key_memoria_mes = "mes_corte_hist_global"
                    
                    # Si no existe en memoria, lo inicializamos con el de la sidebar
                    if key_memoria_mes not in st.session_state:
                        st.session_state[key_memoria_mes] = mes_seleccionado

                    # Buscamos el índice actual de lo que está en memoria dentro de nuestra lista invertida
                    try:
                        indice_default_h = cols_meses_raw.index(st.session_state[key_memoria_mes])
                    except ValueError:
                        indice_default_h = 0 # Por defecto el primero (el más reciente)

                    def format_mm_yyyy(c):
                        try:
                            return pd.to_datetime(c, errors='coerce').strftime('%m/%Y')
                        except:
                            return str(c)

                    col_sel, _ = st.columns([1, 1])
                    with col_sel:
                        mes_corte_orig = st.selectbox(
                            "Considerar deuda hasta el mes (inclusive):",
                            options=cols_meses_raw,
                            format_func=format_mm_yyyy,
                            index=indice_default_h, 
                            key="selector_corte_dinamico" # Llave fija para que no se resetee al cambiar dpto
                        )
                        # Actualizamos la memoria cada vez que el usuario cambia este selector
                        st.session_state[key_memoria_mes] = mes_corte_orig
                    
                    # 3. Cálculo de deuda (Debemos volver al orden original para el rango de meses)
                    # Para el cálculo técnico, usamos la lista original (ascendente) para capturar el rango correcto
                    cols_calculo = [c for c in df_h.columns if "/" in str(c) or "-" in str(c)]
                    cols_calculo.sort(key=lambda x: pd.to_datetime(x, errors='coerce'))
                    
                    idx_limite = cols_calculo.index(mes_corte_orig)
                    meses_a_revisar = cols_calculo[:idx_limite + 1]
                    
                    meses_deuda_lista = []
                    for col in meses_a_revisar:
                        mes_actual_revisado = format_mm_yyyy(col)
                        excluidos_del_periodo = calcular_excluidos_del_mes(mes_actual_revisado, DICCIONARIO_EXCLUSIONES)
                        
                        if u_busqueda.replace(".0", "") in [str(e).replace(".0", "") for e in excluidos_del_periodo]:
                            continue
                            
                        valor = fila_h[col].iloc[0]
                        val_str = str(valor).strip().lower()
                        if val_str not in ["1", "1.0"]:
                            meses_deuda_lista.append(mes_actual_revisado)
                    
                    # --- RENDERIZADO DE RESULTADOS ---
                    cant_meses = len(meses_deuda_lista)
                    meses_texto = ""
                    deuda_total = 0
                    mes_corte_tit = format_mm_yyyy(mes_corte_orig)

                    if cant_meses > 0:
                        meses_texto = ", ".join(meses_deuda_lista)
                        deuda_total = cant_meses * monto_mes_actual
                        
                        st.error(f"🔴 Se detectaron {cant_meses} meses pendientes hasta {mes_corte_tit}.")
                        
                        msj_h = f"🚨 *AVISO DE DEUDA PENDIENTE - {u_seleccionada}*\n"
                        msj_h += f"==========================================\n\n"
                        msj_h += f"Informamos que, revisado el historial al mes de *{mes_corte_tit}*, no registramos el pago de:\n\n"
                        msj_h += f"📅 *Meses:* {meses_texto}\n\n"
                        msj_h += f"💰 *DEUDA TOTAL ESTIMADA: $ {formato_clp(deuda_total)}*\n\n"
                        msj_h += f"Si ya realizó los pagos, por favor ignore este mensaje.\n"
                        msj_h += f"------------------------------------------\n"
                        msj_h += f"*ADMINISTRACIÓN EDIFICIO CREATIVA II*"

                        col_msg, col_btn = st.columns([2, 1])
                        with col_msg:
                            st.text_area("Aviso de Morosidad Histórica:", value=msj_h, height=260, key=f"area_h_{u_seleccionada}_{mes_corte_orig}")
                        
                        with col_btn:
                            import urllib.parse
                            link_h = f"https://api.whatsapp.com/send?text={urllib.parse.quote_plus(msj_h)}"
                            st.markdown(f'''
                                <a href="{link_h}" target="_blank" style="text-decoration:none;">
                                    <div style="background-color:#EA8273; color:white; padding:10px; border-radius:8px; text-align:center; font-weight:bold; cursor:pointer;">
                                        📲 Enviar Cobro Crítico
                                    </div>
                                </a>
                            ''', unsafe_allow_html=True)
                            st.metric("Total Deuda Histórica", f"$ {formato_clp(deuda_total)}")
                            st.write(f"ℹ️ Basado en {cant_meses} meses.")
                    else:
                        st.success(f"✅ El departamento {u_seleccionada} se encuentra al día hasta {mes_corte_tit}.")
                else:
                    st.info(f"No se encontró información para el departamento {u_seleccionada}.")

            except Exception as e:
                st.warning("Error al procesar la deuda histórica.")
                st.exception(e)
#--- Fin del código pestaña: GENERADOR DE AVISOS

#---

# --- NUEVA PESTAÑA: SIMULADOR DE REAJUSTE ---

            # =================================================================
            # BLOQUE: SIMULADOR DE REAJUSTES (Pestaña tab_utilidad)
            # =================================================================
            with tab_utilidad:
                st.write(f"### ⚖️ Simulador de Proyección de Gastos Comunes")
                st.markdown("Analice el impacto de cambiar o reajustar el modelo de cobro mensual.")
                
                # 1. FILTROS DE SEGMENTACIÓN
                with st.expander("🔍 Selección de Datos Base", expanded=True):
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        if 'Período' in df_ing.columns:
                            periodos_unicos = df_ing['Período'].unique()
                            lista_ordenada = sorted(
                                periodos_unicos, 
                                key=lambda x: pd.to_datetime(x, format='%m/%Y'),
                                reverse=True
                            )
                            p_sel = st.selectbox("Período base:", options=lista_ordenada, index=0, key="p_sim_final_v5")
                        else:
                            st.error("No se encontró la columna 'Período'")
                            p_sel = None
                            
                    with col_f2:
                        df_ing['Tipo de Ingreso'] = df_ing['Tipo de Ingreso'].astype(str).str.strip()
                        tipos_disp = sorted(df_ing['Tipo de Ingreso'].unique())
                        
                        opciones_default = [t for t in tipos_disp if t.strip().lower() in ["gastos comunes", "gasto común", "gasto comun"]]
                        
                        f_tipo = st.multiselect(
                            "Concepto a simular:", 
                            options=tipos_disp, 
                            default=opciones_default if opciones_default else [tipos_disp[0]], 
                            key="t_sim_final_v5"
                        )

                if p_sel:
                    # 2. CONFIGURACIÓN DEL MODELO DE COBRO
                    with st.expander("⚙️ Configuración del Modelo de Cobro", expanded=True):
                        # --- NUEVA LÓGICA: DETECCIÓN DE EXCLUSIONES PARA RE-INCORPORAR ---
                        excluidos_periodo_base = calcular_excluidos_del_mes(p_sel, DICCIONARIO_EXCLUSIONES)
                        unidades_a_reincorporar = []
                        
                        if excluidos_periodo_base:
                            st.markdown("---")
                            st.warning(f"🕵️ **Exclusiones detectadas:** Existen {len(excluidos_periodo_base)} unidades excluidas en {p_sel}.")
                            unidades_a_reincorporar = st.multiselect(
                                "Seleccione unidades para DESHABILITAR exclusión (re-incluir en el cobro):",
                                options=excluidos_periodo_base,
                                help="Use esta opción si proyecta que estas unidades volverán a pagar en el período simulado.",
                                key="reinc_excluidas_sim"
                            )
                        st.markdown("---")

                        modelo = st.radio(
                            "Seleccione el modelo a simular:",
                            ["Mantener Prorrateo Actual (Aplicar % o $ extra)", "Abolir Prorrateo (Monto Único Igualitario)"],
                            horizontal=True,
                            key="mod_cobro_final_v5"
                        )
                        
                        c_val1, c_val2 = st.columns([1, 2])
                        with c_val1:
                            if modelo == "Mantener Prorrateo Actual (Aplicar % o $ extra)":
                                tipo_ajuste = st.selectbox("Tipo de ajuste:", ["Porcentaje (%)", "Monto Fijo ($)"], key="t_aj_final_v5")
                                label_input = "Valor del ajuste:"
                            else:
                                tipo_ajuste = "Monto Único"
                                label_input = "Defina el nuevo Monto Fijo Mensual ($):"
                        
                        with c_val2:
                            if tipo_ajuste == "Porcentaje (%)":
                                placeholder_txt = "Ej: 10.5"
                            elif tipo_ajuste == "Monto Fijo ($)":
                                placeholder_txt = "Ej: 5000"
                            else:
                                placeholder_txt = "Ej: 80000"

                            val_texto = st.text_input(
                                label_input, 
                                value="", 
                                placeholder=placeholder_txt, 
                                key=f"v_sim_txt_{modelo}_{tipo_ajuste}"
                            )
                            
                            try:
                                if val_texto.strip() == "":
                                    v_sim = 0.0
                                else:
                                    limpio = val_texto.replace(".", "").replace(",", ".")
                                    v_sim = float(limpio)
                                    
                                    if tipo_ajuste == "Porcentaje (%)":
                                        st.markdown(f"📈 **Aumento aplicado: {v_sim}%**")
                                    else:
                                        st.markdown(f"✅ **Monto ingresado: $ {formato_clp(int(v_sim))}**")
                            except ValueError:
                                st.warning("⚠️ Ingrese un valor numérico válido.")
                                v_sim = 0.0

                    # 3. PROCESAMIENTO
                    df_s = df_ing[(df_ing['Período'] == p_sel) & (df_ing['Tipo de Ingreso'].isin(f_tipo))].copy()
                    
                    # --- MODIFICACIÓN: EXCLUSIÓN DINÁMICA CON OPCIÓN DE RE-INCORPORACIÓN ---
                    # Quitamos de la lista de 'excluidos_periodo_base' aquellas unidades que el usuario marcó para re-incorporar
                    excluidos_finales_sim = [e for e in excluidos_periodo_base if e not in unidades_a_reincorporar]
                    
                    palabras_bloqueadas = ['TOTAL', 'EDIFICIO', 'COMUNIDAD', 'SUBTOTAL', 'RESUMEN']
                    df_s = df_s[
                        (~df_s['Unidad'].astype(str).str.upper().str.contains('|'.join(palabras_bloqueadas))) & 
                        (df_s['Unidad'].notna()) & 
                        (~df_s['Unidad'].astype(str).str.replace(".0","").isin([str(e).replace(".0","") for e in excluidos_finales_sim]))
                    ]
                    # --- FIN MODIFICACIÓN ---
                    
                    df_s['Base'] = pd.to_numeric(df_s['Monto Actual'], errors='coerce').fillna(0)
                    
                    # 4. LÓGICA DE CÁLCULO
                    if modelo == "Mantener Prorrateo Actual (Aplicar % o $ extra)":
                        if tipo_ajuste == "Porcentaje (%)":
                            df_s['Nuevo'] = (df_s['Base'] * (1 + (v_sim / 100))).round(0).astype(int)
                        else:
                            df_s['Nuevo'] = (df_s['Base'] + v_sim).astype(int)
                    else:
                        df_s['Nuevo'] = int(v_sim)
                        
                    df_s['Dif'] = df_s['Nuevo'] - df_s['Base']

                    # 5. PANEL DE MÉTRICAS
                    st.write(f"### 📊 Resumen de Impacto ({p_sel})")
                    
                    try:
                        gastos_mes = df_egr.groupby('Período')['Gasto Mensual'].sum().reset_index()
                        if not gastos_mes.empty:
                            mediana_anual = gastos_mes['Gasto Mensual'].median()
                            umbral = mediana_anual * 1.3
                            gastos_norm = gastos_mes[gastos_mes['Gasto Mensual'] <= umbral]
                            gasto_base = gastos_norm['Gasto Mensual'].mean()
                            meses_extra = len(gastos_mes) - len(gastos_norm)
                        else:
                            gasto_base = 0
                            meses_extra = 0
                    except:
                        gasto_base = 0
                        meses_extra = 0

                    t_b, t_n, t_d = int(df_s['Base'].sum()), int(df_s['Nuevo'].sum()), int(df_s['Dif'].sum())
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Recaudación Actual", f"$ {formato_clp(t_b)}")
                    m2.metric("Nueva Recaudación", f"$ {formato_clp(t_n)}", delta=f"{formato_clp(t_d)}")
                    m3.metric("Diferencia Mensual", f"$ {formato_clp(t_d)}")

                    if gasto_base > 0:
                        cobertura = (t_n / gasto_base) * 100
                        st.info(f"💡 **Análisis de Sostenibilidad:** El gasto operativo base (promedio normalizado) es de **$ {formato_clp(int(gasto_base))}**.")
                        if cobertura < 95:
                            st.error(f"⚠️ El modelo cubre el {cobertura:.1f}% del gasto base. Se mantendrá el déficit operativo.")
                        elif 95 <= cobertura <= 105:
                            st.success(f"✅ El modelo cubre el {cobertura:.1f}% del gasto base. Equilibrio financiero logrado.")
                        else:
                            st.success(f"🚀 El modelo cubre el {cobertura:.1f}% del gasto base. Permite ahorro o fondo de reserva.")

                    # 6. TABLA DE DETALLE
                    st.markdown("---")
                    st.write("### 📋 Detalle por Departamento")
                    df_v = df_s[['Unidad', 'Tipo de Unidad', 'Base', 'Nuevo', 'Dif']].copy()
                    df_v = df_v.sort_values(by='Unidad')
                    
                    for col in ['Base', 'Nuevo', 'Dif']:
                        df_v[col] = df_v[col].apply(formato_clp)
                    
                    st.dataframe(df_v, use_container_width=True, hide_index=True,
                                 column_config={
                                     "Base": "Actual (m2)",
                                     "Nuevo": "Simulado",
                                     "Dif": "Variación"
                                 })
                    
                    # --- MODIFICACIÓN: PIE DE TABLA DINÁMICO ---
                    texto_excluidas = ', '.join(excluidos_finales_sim) if excluidos_finales_sim else "Ninguna"
                    st.caption(f"ℹ️ Analizando {len(df_s)} unidades activas en {p_sel}. Excluidas en este periodo: {texto_excluidas}")
                    if unidades_a_reincorporar:
                        st.success(f"🔄 Unidades re-incorporadas manualmente para simulación: {', '.join(unidades_a_reincorporar)}")

#------ 

# --- PESTAÑA: RESUMEN FINANCIERO (Análisis 12 Meses) ---
# =================================================================
            with tab_salud:
                st.write(f"### 🌊 Análisis de Flujo y Sostenibilidad")
                st.write("Comparativa de los últimos 12 meses: Recaudación Comunidad vs. Gastos Operacionales.")

                # 1. PREPARACIÓN DE DATOS (EXCLUSIÓN DE DEUDA/SALDOS)
                df_i = df_ing.copy()
                # Lista negra: lo que NO es aporte fresco del mes
                excluir = ["Saldo Mes Anterior", "Otros Ingresos", "Préstamos", "Préstamo"]
                
                df_i['Tipo_Limpio'] = df_i['Tipo de Ingreso'].astype(str).str.strip()
                df_i_real = df_i[~df_i['Tipo_Limpio'].isin(excluir)].copy()
                
                # --- MODIFICACIÓN PARA MANTENER LÓGICA DE EXCLUSIÓN ---
                registros_validos = []
                for periodo_f in df_i_real['Período'].unique():
                    # Obtener unidades excluidas para este mes específico
                    excluidos_mes = calcular_excluidos_del_mes(periodo_f, DICCIONARIO_EXCLUSIONES)
                    temp_df = df_i_real[df_i_real['Período'] == periodo_f].copy()
                    
                    # Filtrar unidades que no debían pagar en este periodo
                    temp_df = temp_df[~temp_df['Unidad'].astype(str).str.replace(".0","").isin([str(e).replace(".0","") for e in excluidos_mes])]
                    registros_validos.append(temp_df)
                
                if registros_validos:
                    df_i_real = pd.concat(registros_validos)
                # -----------------------------------------------------

                df_i_real['Monto_Limpio'] = pd.to_numeric(df_i_real['Pago Total del Mes'], errors='coerce').fillna(0)
                res_i = df_i_real.groupby('Período')['Monto_Limpio'].sum().reset_index(name='Monto')
                res_i['Tipo'] = 'Recaudación Comunidad'

                # Egresos
                df_e = df_egr.copy()
                df_e['Monto_Egr'] = pd.to_numeric(df_e['Gasto Mensual'], errors='coerce').fillna(0)
                res_e = df_e.groupby('Período')['Monto_Egr'].sum().reset_index(name='Monto')
                res_e['Tipo'] = 'Gastos Reales'

                # 2. ORDENAMIENTO CRONOLÓGICO PARA 12 MESES
                try:
                    periodos_ordenados = sorted(df_ing['Período'].unique(), 
                                               key=lambda x: datetime.strptime(str(x), '%m/%Y'))
                except Exception:
                    periodos_ordenados = sorted(df_ing['Período'].unique())
                
                # CAMBIO AQUÍ: Tomamos los últimos 12
                ultimos_12 = periodos_ordenados[-12:]
                
                # Preparamos dataframes para el gráfico mixto
                df_res_i = res_i[res_i['Período'].isin(ultimos_12)].set_index('Período').reindex(ultimos_12).reset_index().fillna(0)
                df_res_e = res_e[res_e['Período'].isin(ultimos_12)].set_index('Período').reindex(ultimos_12).reset_index().fillna(0)

                # 3. GRÁFICO MIXTO (RECAUDACIÓN EN BARRAS VS GASTOS EN LÍNEA)
                import plotly.graph_objects as go
                fig_comp = go.Figure()

                # Barras de Recaudación (Verde)
                fig_comp.add_trace(go.Bar(
                    x=df_res_i['Período'],
                    y=df_res_i['Monto'],
                    name='Recaudación Comunidad',
                    marker_color='#2ECC71',
                    opacity=0.8,
                    # Formato con puntos en miles para el hover
                    hovertemplate="Recaudación: $ %{y:,.0f}<extra></extra>"
                ))

                # Línea de Gastos Reales (Rojo)
                fig_comp.add_trace(go.Scatter(
                    x=df_res_e['Período'],
                    y=df_res_e['Monto'],
                    name='Gastos Reales (Límite)',
                    mode='lines+markers',
                    line=dict(color='#E74C3C', width=4),
                    marker=dict(size=10, symbol='circle'),
                    # Formato con puntos en miles para el hover
                    hovertemplate="Gastos: $ %{y:,.0f}<extra></extra>"
                ))
                
                fig_comp.update_layout(
                    separators=",.", # Fuerza el uso de punto para miles y coma para decimales
                    # TÍTULO MÁS GRANDE
                    title=dict(
                        text="Histórico de Flujo de Caja (12 Meses)",
                        font=dict(size=22)
                    ),
                    # Aumentamos el margen izquierdo para que quepan los números largos
                    margin=dict(l=80, r=20, t=60, b=120), 
                    height=600, # Un poco más de altura para legibilidad
                    paper_bgcolor='rgba(0,0,0,0)', 
                    plot_bgcolor='rgba(0,0,0,0)',
                    hovermode="x unified",
                    
                    xaxis=dict(
                        title_text="Meses Analizados",
                        tickmode='linear',
                        tickangle=-30, 
                        automargin=True,
                        categoryorder='array',
                        categoryarray=ultimos_12,
                        tickfont=dict(size=14),
                        title_font=dict(size=16),
                        dtick=1 
                    ),

                    yaxis=dict(
                        tickformat="$,.0f", # Aplica el formato de moneda en el eje vertical
                        tickfont=dict(size=14, color='black'),
                        gridcolor='rgba(200, 200, 200, 0.2)'
                    ),
                    
                    legend=dict(
                        orientation="h",
                        font=dict(size=16, color='black'), 
                        yanchor="bottom", 
                        y=-0.45, 
                        xanchor="center", 
                        x=0.5
                    )
                )

                # 4. CÁLCULO DE SOSTENIBILIDAD
                # Reutilizamos df_plot para los cálculos de texto de abajo
                df_tendencia = pd.concat([res_i, res_e])
                df_plot = df_tendencia[df_tendencia['Período'].isin(ultimos_12)]
                
                df_calc = df_plot.pivot(index='Período', columns='Tipo', values='Monto').fillna(0)
                if 'Recaudación Comunidad' not in df_calc.columns: df_calc['Recaudación Comunidad'] = 0
                if 'Gastos Reales' not in df_calc.columns: df_calc['Gastos Reales'] = 1
                
                df_calc['Capacidad'] = (df_calc['Recaudación Comunidad'] / df_calc['Gastos Reales'] * 100)
                avg_capacidad = df_calc['Capacidad'].mean()

                # Visualización
                st.plotly_chart(fig_comp, use_container_width=True, config={'displayModeBar': False})

                # --- AJUSTE AQUÍ: Inicializamos diag_msg para que siempre exista ---
                diag_msg = "Análisis de sostenibilidad de los últimos 12 meses."
                
                st.write(f"### 🛡️ Diagnóstico del Año: {avg_capacidad:.1f}% de Sostenibilidad")
                if avg_capacidad < 100:
                    st.error(f"En promedio, el edificio gasta más de lo que recauda. Faltó un {100-avg_capacidad:.1f}% para cubrir el año con ingresos propios.")
                else:
                    st.success(f"El edificio es financieramente sano. Recaudó un {avg_capacidad-100:.1f}% por sobre sus gastos en el último año.")
                
                st.caption("ℹ️ Nota: El análisis de recaudación solo considera unidades obligadas al pago según el calendario de exclusiones.")


                # --- BONUS DE TRANSPARENCIA: ¿A DÓNDE VA EL DINERO? ---
                st.divider()
                st.write("### 🍰 ¿Cómo se distribuye el gasto de la Comunidad?")
                st.markdown("Análisis histórico de los últimos 12 meses por tipo de egreso principal.")

                if not df_egr.empty:
                    try:
                        # 1. Definición de columnas según tu jerarquía
                        col_analisis = "Tipo de Egreso" 
                        col_monto = "Gasto Mensual"

                        # 2. Definimos los últimos 12 períodos disponibles
                        periodos_disponibles = sorted(
                            df_egr['Período'].unique(), 
                            key=lambda x: pd.to_datetime(x, format='%m/%Y')
                        )
                        ultimos_12_analisis = periodos_disponibles[-12:]

                        # 3. Agrupamos por la clase mayor (Tipo de Egreso) filtrando esos 12 meses
                        df_pie = df_egr[df_egr['Período'].isin(ultimos_12_analisis)].groupby(col_analisis)[col_monto].sum().reset_index()
                        
                        # 4. Creamos el gráfico de torta (Donut)
                        import plotly.express as px
                        fig_pie = px.pie(
                            df_pie, 
                            values=col_monto, 
                            names=col_analisis,
                            hole=0.4,
                            color_discrete_sequence=px.colors.qualitative.Pastel
                        )
                        
                        # MODIFICACIÓN: Hover con formato de punto para miles y quitar etiqueta extra
                        fig_pie.update_traces(
                            textposition='inside', 
                            textinfo='percent+label',
                            insidetextfont=dict(size=16, color='black'), # Aquí ajustas el tamaño (ej: 16)
                            hovertemplate="<b>%{label}</b><br>Monto: $ %{value:,.0f}<extra></extra>"
                        )

                        fig_pie.update_layout(
                            separators=",.", # Establece el punto como separador de miles
                            margin=dict(l=20, r=20, t=30, b=50),
                            showlegend=True,
                            legend=dict(orientation="h", y=-0.2, font=dict(size=14)),
                            height=500
                        )

                        # 5. Mostrar en dos columnas: Gráfico + Explicación Pedagógica
                        col_chart, col_txt = st.columns([2, 1])
                        
                        with col_chart:
                            st.plotly_chart(fig_pie, use_container_width=True)
                        
                        with col_txt:
                            st.write("#### 💰 Por cada $1.000 que pagas:")
                            total_gasto_anual = df_pie[col_monto].sum()
                            df_pie_sorted = df_pie.sort_values(by=col_monto, ascending=False)
                            
                            for _, row in df_pie_sorted.iterrows():
                                proporcion_pesos = (row[col_monto] / total_gasto_anual) * 1000
                                if proporcion_pesos >= 1:
                                    # MODIFICACIÓN: Fuente más grande y valor en negrita mediante markdown
                                    st.markdown(f"<span style='font-size: 1.1rem;'>{row[col_analisis]}: **${int(proporcion_pesos)}**</span>", unsafe_allow_html=True)
                            
                            st.caption(f"Nota: Basado en {len(ultimos_12_analisis)} meses de historial.")
                            
                    except Exception as e:
                        st.error(f"Error al procesar el desglose de gastos: {e}")
                else:
                    st.warning("No hay datos de egresos suficientes para generar el desglose.")


                # 5. TABLA RESUMEN ORDENADA
                # =================================================================
                st.markdown("---")
                st.write(f"### 📋 Desglose Mensual: Recaudación vs. Gastos")

                # NOTA INFORMATIVA CON ESTILO AZUL CLARO
                st.info("ℹ️ **Nota:** Este cuadro compara el total de ingresos recibidos frente a los egresos ejecutados en cada período. La diferencia positiva o negativa, representa el excedente o déficit operativo mensual de la comunidad. (No considera saldos de arrastre de meses anteriores)")
                
                # 1. Calculamos la diferencia numérica en df_calc (esto lo usa el PDF)
                df_calc['Diferencia'] = df_calc['Recaudación Comunidad'] - df_calc['Gastos Reales']
                df_view = df_calc.reset_index()
                
                # 2. Ordenamiento cronológico
                df_view['Período'] = pd.Categorical(df_view['Período'], categories=ultimos_12, ordered=True)
                df_view = df_view.sort_values('Período')

                # 3. Función para el semáforo visual de pantalla
                def agregar_indicador(monto):
                    # Usamos TU función formato_clp
                    monto_formateado = formato_clp(monto)
                    if monto > 0: return f"🟢 {monto_formateado}"
                    elif monto < 0: return f"🔴 {monto_formateado}"
                    else: return f"⚪ {monto_formateado}"

                # 4. Creamos df_final (Mantenemos este nombre para que no de NameError más abajo)
                df_final = df_view[['Período', 'Recaudación Comunidad', 'Gastos Reales', 'Diferencia']].copy()
                
                # Convertimos a texto para la visualización web
                df_final['Recaudación Comunidad'] = df_final['Recaudación Comunidad'].apply(formato_clp)
                df_final['Gastos Reales'] = df_final['Gastos Reales'].apply(formato_clp)
                
                # IMPORTANTE: Aplicamos el indicador sobre el valor numérico de df_view
                df_final['Diferencia'] = df_view['Diferencia'].apply(agregar_indicador)

                # 5. Mostramos la tabla en Streamlit
                st.dataframe(df_final, use_container_width=True, hide_index=True)

#------ 
# --- DENTRO DE tab_salud (al final del bloque) ---
                import plotly.offline as opy
                
                # 1. Convertimos los gráficos de Resumen Financiero a DIVs de HTML
                div_grafico_tendencia = opy.plot(fig_comp, auto_open=False, output_type='div', include_plotlyjs='cdn')
                
                try:
                    div_grafico_torta = opy.plot(fig_pie, auto_open=False, output_type='div', include_plotlyjs='cdn')
                except NameError:
                    div_grafico_torta = "<div>No hay datos de distribución disponibles.</div>"

                # 2. Guardamos el diagnóstico y el color para la interfaz web
                salud_web_data = {
                    'avg_sostenibilidad': f"{avg_capacidad:.1f}%",
                    'status_msg': "El edificio es financieramente sano." if avg_capacidad >= 100 else "El edificio gasta más de lo que recauda.",
                    'color_salud': "#27ae60" if avg_capacidad >= 100 else "#e74c3c"
                }


# --- CAPTURA DE DATOS PARA EL PORTAL WEB ---
                explicacion_mil = {}
                if 'df_pie_sorted' in locals():
                    total_g_anual = df_pie_sorted[col_monto].sum()
                    for _, row in df_pie_sorted.iterrows():
                        prop = (row[col_monto] / total_g_anual) * 1000
                        if prop >= 1:
                            explicacion_mil[row[col_analisis]] = int(prop)
                
                # Guardamos la tabla final de salud
                df_salud_web = df_final.copy()


#------ 
                # =================================================================
                # MOTOR DE REPORTES PDF - ANÁLISIS INTEGRAL
                # =================================================================
                try:
                    from fpdf import FPDF
                    from datetime import datetime
                    
                    # 1. CÁLCULOS PREVIOS (Silenciosos)
                    tipos_en_datos = df_ing['Tipo de Ingreso'].unique()
                    tiene_ce = "Cuotas extraordinarias" in tipos_en_datos
                    tiene_pr = "Otros Ingresos" in tipos_en_datos
                    
                    # --- LÓGICA DE FECHAS PARA EL ENCABEZADO ---
                    lista_periodos = df_view['Período'].tolist()
                    # Convertimos a objetos datetime para encontrar los extremos reales
                    objetos_fecha = [datetime.strptime(p, "%m/%Y") for p in lista_periodos]
                    objetos_fecha.sort()
                    
                    mes_inicio = objetos_fecha[0].strftime("%m/%Y") if objetos_fecha else "N/A"
                    mes_fin = objetos_fecha[-1].strftime("%m/%Y") if objetos_fecha else "N/A"
                    total_meses = len(objetos_fecha)
                    
                    g_base_val = 0
                    m_atipicos = 0
                    try:
                        g_mes = df_egr.groupby('Período')['Gasto Mensual'].sum().reset_index()
                        if not g_mes.empty:
                            mediana_a = g_mes['Gasto Mensual'].median()
                            u_limite = mediana_a * 1.3
                            g_norm = g_mes[g_mes['Gasto Mensual'] <= u_limite]
                            g_base_val = int(g_norm['Gasto Mensual'].mean())
                            m_atipicos = len(g_mes) - len(g_norm)
                    except:
                        pass

                    class PDFReport(FPDF):
                        def __init__(self, periodo_actual, *args, **kwargs):
                            super().__init__(*args, **kwargs)
                            self.periodo_actual = periodo_actual # Guardamos el mes aquí

                        def header(self):
                            self.set_font('Helvetica', 'B', 15)
                            self.set_text_color(44, 62, 80)
                            self.cell(190, 10, 'INFORME DE GESTION FINANCIERA - CREATIVA II', 0, 1, 'C')
                            
                            self.set_font('Helvetica', 'I', 9)
                            self.cell(190, 5, f'Generado: {datetime.now().strftime("%d/%m/%Y")}', 0, 1, 'C')
                            
                            self.set_font('Helvetica', 'B', 10)
                            self.set_text_color(70, 70, 70)
                            
                            # Ahora usamos self.periodo_actual que ya vive dentro de la clase
                            if self.page_no() == 1:
                                texto_header = f'Analisis Historico de {total_meses} meses | Periodo: {mes_inicio} a {mes_fin}'
                            else:
                                texto_header = f'Balance Mensual de Ejecucion | Mes: {self.periodo_actual}'
                            
                            self.cell(190, 7, texto_header.encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'C')
                            self.ln(5)

                            # --- LA LÍNEA SUTIL DE SEPARACIÓN ---
                            # Dibujamos la línea justo después del texto del encabezado
                            self.set_draw_color(44, 62, 80) # Mismo color azul oscuro del título
                            self.set_line_width(0.1)        # Grosor sutil
                            # line(x1, y1, x2, y2)
                            self.line(10, 34, 200, 34)      # Posición horizontal de margen a margen
                            
                            self.ln(8) # Espacio después de la línea para que el contenido no quede pegado

                        def footer(self):
                            self.set_y(-15)
                            self.set_font('Helvetica', 'I', 8)
                            self.set_text_color(128, 128, 128)
                            self.cell(190, 10, f'Pagina {self.page_no()}', 0, 0, 'C')

                    def generar_reporte_maestro(df, cap, msg_original, c_ext, prest, gb, m_ext, ult_p):
                        pdf = PDFReport(ult_p, format='letter') # Definimos pagina tamaño Carta
                        pdf.set_auto_page_break(auto=True, margin=15)
                        pdf.add_page()

                        # --- SECCIÓN 1: NOTAS DINÁMICAS (CUADRO AMARILLO) ---
                        if c_ext or prest:
                            pdf.set_font('Helvetica', 'B', 10)
                            pdf.set_fill_color(255, 243, 205)
                            pdf.set_text_color(133, 100, 4)
                            
                            notas = ["NOTA DE COMPOSICION FINANCIERA:"]
                            if c_ext: notas.append("- Se incluyen Cuotas Extraordinarias que explican desviaciones al alza.")
                            if prest: notas.append("- Se detectan 'Otros Ingresos' (Prestamos), indicando deficit operativo.")
                            
                            pdf.multi_cell(190, 6, "\n".join(notas).encode('latin-1', 'replace').decode('latin-1'), 1, 'L', True)
                            pdf.ln(5)

                        # --- SECCIÓN 2: GASTO BASE INTEGRADO ---
                        if gb > 0:
                            pdf.set_font('Helvetica', 'B', 12)
                            pdf.set_text_color(44, 62, 80)
                            pdf.cell(190, 10, '1. ANALISIS DE GASTO OPERATIVO BASE', 0, 1, 'L')
                            
                            pdf.set_font('Helvetica', '', 11)
                            pdf.set_text_color(0, 0, 0)
                            frase = f"El gasto operativo base (excluyendo {m_ext} meses de gastos atipicos) se determina en: "
                            pdf.write(6, frase.encode('latin-1', 'replace').decode('latin-1'))
                            
                            pdf.set_font('Helvetica', 'B', 12)
                            pdf.set_text_color(0, 51, 102)
                            pdf.write(6, f"$ {formato_clp(gb)}")
                            pdf.ln(12)

                        # --- SECCIÓN 3: RESUMEN DE SOSTENIBILIDAD (ANÁLISIS RESTAURADO) ---
                        pdf.set_font('Helvetica', 'B', 12)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 10, '2. SOSTENIBILIDAD DEL PERIODO', 0, 1, 'L')
                        
                        bg_color = (254, 235, 235) if cap < 100 else (235, 254, 235)
                        pdf.set_fill_color(*bg_color)
                        pdf.set_font('Helvetica', 'B', 13)
                        
                        t_color = (150, 0, 0) if cap < 100 else (0, 100, 0)
                        pdf.set_text_color(*t_color)
                        pdf.cell(190, 12, f"  Resumen Financiero: {cap:.1f}% de Sostenibilidad", 0, 1, 'L', True)
                        
                        pdf.ln(2)
                        pdf.set_font('Helvetica', 'I', 11)
                        pdf.set_text_color(0, 0, 0)
                        
                        # Lógica de diagnóstico restaurada
                        if cap < 100:
                            falta = 100 - cap
                            msg_analisis = (f"Aun considerando los ingresos extraordinarios, el edificio gasta mas de lo que recauda. "
                                           f"Falto un {falta:.1f}% para cubrir el período. Esto evidencia que los pagos regulares "
                                           f"son insuficientes para la operacion actual.")
                        else:
                            msg_analisis = "La recaudacion total ha logrado equilibrar y sostener el gasto del periodo analizado."
                            
                        pdf.multi_cell(190, 7, msg_analisis.encode('latin-1', 'replace').decode('latin-1'))
                        pdf.ln(5)

                        # --- SECCIÓN 4: TABLA DE MOVIMIENTOS ---
                        pdf.set_font('Helvetica', 'B', 12)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 10, '3. DETALLE DE MOVIMIENTOS MENSUALES', 0, 1, 'L')
                        
                        pdf.set_font('Helvetica', 'B', 10)
                        pdf.set_fill_color(52, 73, 94)
                        pdf.set_text_color(255, 255, 255)
                        
                        w = [35, 50, 50, 55]
                        columnas = ["Mes", "Recaudacion Tot.", "Gastos Reales", "Diferencia"]
                        for i, col in enumerate(columnas):
                            pdf.cell(w[i], 8, col, 1, 0, 'C', True)
                        pdf.ln()

                        pdf.set_font('Helvetica', '', 9)
                        pdf.set_text_color(0, 0, 0)
                        for _, row in df.iterrows():
                            pdf.cell(w[0], 7, str(row['Período']), 1, 0, 'C')
                            pdf.cell(w[1], 7, formato_clp(int(row['Recaudación Comunidad'])), 1, 0, 'R')
                            pdf.cell(w[2], 7, formato_clp(int(row['Gastos Reales'])), 1, 0, 'R')
                            
                            dif = int(row['Diferencia'])
                            if dif < 0: pdf.set_text_color(200, 0, 0)
                            else: pdf.set_text_color(0, 120, 0)
                            
                            pdf.cell(w[3], 7, formato_clp(dif), 1, 1, 'R')
                            pdf.set_text_color(0, 0, 0)


                        # --- PÁGINA 2: BALANCE DE CIERRE (RESUMEN EJECUTIVO) ---
                        pdf.add_page()
                        
                        # Encabezado de Página
                        pdf.set_font('Helvetica', 'B', 14)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 10, f'BALANCE DE CIERRE: {ult_p}', 0, 1, 'L')
                        pdf.set_draw_color(44, 62, 80)
                        pdf.ln(8) 

                        # 1. SECCIÓN: ANÁLISIS DE RECAUDACIÓN
                        pdf.set_font('Helvetica', 'B', 11)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 8, "1. EFICIENCIA DE RECAUDACION (Gastos Comunes)", 0, 1, 'L')
                        pdf.ln(2)
                        
                        # Tabla de Eficiencia
                        pdf.set_font('Helvetica', 'B', 9); pdf.set_fill_color(245, 245, 245); pdf.set_text_color(0, 0, 0)
                        pdf.cell(47.5, 7, "Esperado", 1, 0, 'C', True)
                        pdf.cell(47.5, 7, "Recaudado", 1, 0, 'C', True)
                        pdf.cell(47.5, 7, "Diferencia", 1, 0, 'C', True)
                        pdf.cell(47.5, 7, "Estado", 1, 1, 'C', True)

                        pdf.set_font('Helvetica', '', 10)
                        pdf.cell(47.5, 9, f"$ {formato_clp(monto_esperado)}", 1, 0, 'C')
                        pdf.cell(47.5, 9, f"$ {formato_clp(monto_recaudado)}", 1, 0, 'C')
                        
                        # Lógica para evitar signos negativos en saldos a favor
                        val_diff = abs(monto_pendiente)
                        txt_diff = f"$ {formato_clp(val_diff)}"
                        if monto_pendiente < 0:
                            txt_diff += " (A favor)"
                        
                        pdf.cell(47.5, 9, txt_diff, 1, 0, 'C')
                        
                        # Lógica de color para Estado
                        pdf.set_font('Helvetica', 'B', 10)
                        if "Excelente" in estado_salud or "Satisfactorio" in estado_salud:
                            pdf.set_text_color(0, 120, 0)
                        elif "Atencion" in estado_salud:
                            pdf.set_text_color(200, 150, 0)
                        else:
                            pdf.set_text_color(200, 0, 0)
                        
                        estado_pdf = estado_salud.replace("✅ ", "").replace("🟢 ", "").replace("⚠️ ", "").replace("🔴 ", "")
                        pdf.cell(47.5, 9, estado_pdf.upper(), 1, 1, 'C')
                        
                        pdf.set_font('Helvetica', 'I', 9); pdf.set_text_color(100, 100, 100)
                        pdf.cell(190, 6, f"    Nota: La recaudacion representa un {pct_recaudacion:.1f}% del total de dinero que debía ingresar este mes.", 0, 1, 'L')
                        pdf.ln(8) # Espacio entre secciones

                        # 2. SECCIÓN: FLUJO DE CAJA
                        pdf.set_font('Helvetica', 'B', 11); pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 8, "2. MOVIMIENTO DE CAJA REAL (Flujo Total)", 0, 1, 'L')
                        pdf.ln(2)
                        
                        pdf.set_font('Helvetica', 'B', 10); pdf.set_fill_color(240, 240, 240); pdf.set_text_color(0, 0, 0)
                        pdf.cell(47.5, 8, "Saldo Inicial", 1, 0, 'C', True)
                        pdf.cell(47.5, 8, "Ingresos (+)", 1, 0, 'C', True)
                        pdf.cell(47.5, 8, "Egresos (-)", 1, 0, 'C', True)
                        pdf.cell(47.5, 8, "Saldo Final", 1, 1, 'C', True)
                        
                        pdf.set_font('Helvetica', '', 11)
                        pdf.cell(47.5, 11, f"$ {formato_clp(saldo_inicial)}", 1, 0, 'C')
                        pdf.cell(47.5, 11, f"$ {formato_clp(ingresos_puros)}", 1, 0, 'C')
                        pdf.cell(47.5, 11, f"$ {formato_clp(gastos_puros)}", 1, 0, 'C')
                        pdf.set_font('Helvetica', 'B', 11); pdf.cell(47.5, 11, f"$ {formato_clp(saldo_final)}", 1, 1, 'C')
                        pdf.ln(6)

                        # 3. SECCIÓN: ANÁLISIS DE SOLVENCIA
                        if variacion_pct > 0:
                            pdf.set_text_color(150, 0, 0); pdf.set_fill_color(255, 245, 245)
                            msg_v = f" RESULTADO DEL MES: DEFICIT DEL {variacion_pct:.1f}%"
                        else:
                            pdf.set_text_color(0, 100, 0); pdf.set_fill_color(245, 255, 245)
                            msg_v = f" RESULTADO DEL MES: SUPERAVIT DEL {abs(variacion_pct):.1f}%"
                        
                        pdf.set_font('Helvetica', 'B', 11)
                        pdf.cell(190, 10, msg_v, 0, 1, 'L', True)
                        pdf.ln(3)

                        pdf.set_font('Helvetica', '', 10); pdf.set_text_color(0, 0, 0)
                        if variacion_pct > 0:
                            dif_monto = abs(ingresos_puros - gastos_puros)
                            nota_caja = (f"Los egresos superaron la recaudacion en un {variacion_pct:.1f}%. "
                                        f"La diferencia de $ {formato_clp(dif_monto)} se cubrio con el "
                                        f"saldo del mes anterior (Saldo Inicial).")
                            pdf.multi_cell(190, 7, nota_caja.encode('latin-1', 'replace').decode('latin-1'), 0, 'L')
                        pdf.ln(5)

                        # 4. DISPONIBILIDAD DE FONDOS
                        recursos_totales = saldo_inicial + ingresos_puros
                        disp_total = (saldo_final / recursos_totales * 100) if recursos_totales > 0 else 0
                        pdf.set_font('Helvetica', 'B', 11); pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 8, f"Disponibilidad Total de Fondos: {disp_total:.1f}%", 0, 1, 'L')
                        pdf.set_font('Helvetica', 'I', 9); pdf.set_text_color(100, 100, 100)
                        pdf.cell(190, 5, "Porcentaje de dinero remanente sobre el total de recursos que circularon en el mes.", 0, 1, 'L')
                        pdf.ln(10) # Salto amplio antes de la tabla final

                        # 5. TABLA RESUMEN DE EGRESOS (Con Clasificaciones únicas)
                        pdf.set_font('Helvetica', 'B', 11); pdf.set_text_color(44, 62, 80)
                        pdf.cell(190, 8, "3. DETALLE RESUMIDO DE EGRESOS", 0, 1, 'L')
                        pdf.ln(2)
                        
                        pdf.set_font('Helvetica', 'B', 9); pdf.set_fill_color(52, 73, 94); pdf.set_text_color(255, 255, 255)
                        pdf.cell(130, 8, " Tipo de Egreso (Clasificaciones)", 1, 0, 'L', True)
                        pdf.cell(60, 8, "Total Ejecutado ", 1, 1, 'R', True)
                        
                        pdf.set_font('Helvetica', '', 9); pdf.set_text_color(0, 0, 0)
                        
                        df_egr_mes = df_egr[df_egr['Período'] == ult_p]
                        if not df_egr_mes.empty:
                            # Agrupamos montos y concatenamos clasificaciones únicas
                            egr_resumen = df_egr_mes.groupby('Tipo de Egreso').agg({
                                'Gasto Mensual': 'sum',
                                'Clasificación': lambda x: ", ".join(sorted(set(x.astype(str))))
                            }).reset_index().sort_values('Gasto Mensual', ascending=False)
                            
                            for _, r in egr_resumen.iterrows():
                                # Texto con paréntesis y limpieza de duplicados
                                texto_tipo = f" {r['Tipo de Egreso']} ({r['Clasificación']})"
                                
                                # Si es muy largo, cortamos para no romper la tabla
                                if len(texto_tipo) > 85:
                                    texto_tipo = texto_tipo[:82] + "..."
                                    
                                pdf.cell(130, 8, texto_tipo.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'L')
                                pdf.cell(60, 8, f"$ {formato_clp(int(r['Gasto Mensual']))} ", 1, 1, 'R')
                        
                        # Cierre de tabla
                        pdf.set_font('Helvetica', 'B', 10); pdf.set_fill_color(240, 240, 240)
                        pdf.cell(130, 9, " TOTAL GASTOS EJECUTADOS EN EL MES", 1, 0, 'L', True)
                        pdf.cell(60, 9, f"$ {formato_clp(int(gastos_puros))} ", 1, 1, 'R', True)

                        return bytes(pdf.output())

                    # 2. EJECUCIÓN Y DESCARGA
                    # Agregamos 'mes_fin' como el octavo argumento (que la función recibirá como 'ult_p')
                    pdf_final_bytes = generar_reporte_maestro(
                        df_view, 
                        avg_capacidad, 
                        diag_msg, 
                        tiene_ce, 
                        tiene_pr, 
                        g_base_val, 
                        m_atipicos,
                        mes_fin  # <--- ESTE ES EL ARGUMENTO QUE FALTA
                    )
 
# --- SECCIÓN DE EXPORTACIÓN FINAL ---
                    st.markdown("---")
                    st.markdown("""
                        <div style="background-color: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 5px solid #2f5597; margin-bottom: 20px;">
                            <h4 style="margin-top: 0; color: #2f5597;">📊 Informe de Gestión Financiero</h4>
                            <p style="font-size: 0.9rem; color: #444;">
                                El documento contiene el <b>resumen consolidado</b> de ingresos/egresos y el <b>detalle cronológico</b> de todos los movimientos del periodo seleccionado.
                            </p>
                        </div>
                    """, unsafe_allow_html=True)

                    st.download_button(
                        label="📥 Descargar Informe de Gestión",
                        data=pdf_final_bytes,
                        file_name=f"Informe_Gestión_Creativa_II_{datetime.now().strftime('%m_%Y')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"Error al generar reporte: {e}")
#------ 

        else:
            st.error("No se encontró la hoja 'ControlPagoxDeptos'.")


# --- GATILLO FINAL PARA GENERAR EL HTML ---
try:
    import base64  # Importamos para codificar el PDF
    from datetime import datetime # Importamos para la estampa de tiempo

    # Capturamos la fecha y hora actual
    fecha_actualizacion = datetime.now().strftime("%d/%m/%Y %H:%M")

    # 1. Verificamos si las variables existen (Balance)
    r_display = df_resumen_display if 'df_resumen_display' in locals() else pd.DataFrame()
    m_web = mis_metricas_web if 'mis_metricas_web' in locals() else {'ingresos': '0', 'egresos': '0'}
    
    # 2. Datos de Morosidad (Pestaña 3)
    m_mes = df_display_mes if 'df_display_mes' in locals() else pd.DataFrame()
    m_hist = df_display if 'df_display' in locals() else pd.DataFrame()
    m_salud_dict = {
        'total_historico': formato_clp(df_hist_final['Deuda Total Est.'].sum()) if 'df_hist_final' in locals() else '0'
    }
    
    # 3. Datos de Resumen Financiero (Pestaña 2)
    d_tendencia = div_grafico_tendencia if 'div_grafico_tendencia' in locals() else "<div>Cargando análisis...</div>"
    d_torta = div_grafico_torta if 'div_grafico_torta' in locals() else "<div>Cargando distribución...</div>"
    s_data = salud_web_data if 'salud_web_data' in locals() else {'avg_sostenibilidad': '0%', 'status_msg': 'Cargando...', 'color_salud': '#cccccc'}
    
    # --- NUEVOS DATOS PARA SALUD Y CONTROL ---
    df_s_f = df_salud_web if 'df_salud_web' in locals() else pd.DataFrame()
    e_mil = explicacion_mil if 'explicacion_mil' in locals() else {}
    d_control = df_control if 'df_control' in locals() else None

    # --- NUEVO: PREPARACIÓN DE INFORMES PDF POR SEPARADO ---
    pdf_b64_estado = ""
    pdf_b64_gestion = ""

    if 'pdf_raw' in locals() and pdf_raw is not None:
        try:
            raw_data = pdf_raw.encode('latin-1') if isinstance(pdf_raw, str) else pdf_raw
            pdf_b64_estado = base64.b64encode(raw_data).decode('utf-8')
        except:
            pdf_b64_estado = ""

    if 'pdf_final_bytes' in locals() and pdf_final_bytes is not None:
        try:
            gestion_data = pdf_final_bytes.encode('latin-1') if isinstance(pdf_final_bytes, str) else pdf_final_bytes
            pdf_b64_gestion = base64.b64encode(gestion_data).decode('utf-8')
        except:
            pdf_b64_gestion = ""

    # 4. LLAMADA ACTUALIZADA
    generar_portal_web(
        df_resumen=r_display,                                          # 1
        df_maestra=df_maestra_display if 'df_maestra_display' in locals() else pd.DataFrame(), # 2
        metricas=m_web,                                                # 3
        df_mora_mes=m_mes,                                             # 4
        df_mora_hist=m_hist,                                           # 5
        m_salud=m_salud_dict,                                          # 6
        div_tendencia=d_tendencia,                                     # 7
        div_torta=d_torta,                                             # 8
        salud_data=s_data,                                             # 9
        mes_actual=mes_seleccionado,                                   # 10
        fig_gauge=fig_gauge if 'fig_gauge' in locals() else None,      # 11
        df_salud_final=df_s_f,                                         # 12
        explicacion_mil=e_mil,                                         # 13
        df_control=d_control,                                          # 14
        pdf_b64=pdf_b64_estado,                                        # 15
        pdf_gestion_b64=pdf_b64_gestion,                               # 16
        unidades_excluidas=UNIDADES_EXCLUIDAS if 'UNIDADES_EXCLUIDAS' in locals() else [], # 17
        fecha_actualizacion=fecha_actualizacion,                       # 18

        # --- ESTA SON LAS VARIABLES PARA APAGAR ALGUNA DE LAS PESTAÑAS EN EL PORTAL (True / False) ---
        ver_balance=True, 
        ver_resumen_fin=True, 
        ver_morosidad=True, 
        ver_control=True, 
        ver_pdfs=True
    )

    # --- LÓGICA DE VISIBILIDAD (CON RE-GENERACIÓN ACTIVA) ---
    if 'div_grafico_tendencia' in locals() and 'df_display_mes' in locals() and 'df_control' in locals():
        
        # 1. Grabamos la fecha en la variable temporal si no existe
        if 'fecha_temporal_sincro' not in st.session_state:
            st.session_state['fecha_temporal_sincro'] = datetime.now().strftime("%d/%m/%Y %H:%M")
            st.rerun()
        
        # 2. Sincronizamos el estado lógico para el botón
        if not st.session_state.get('sincro_exitosa', False):
            st.session_state['sincro_exitosa'] = True
            st.rerun()
            
        # NOTA: Ya no usamos espacio_servidor.markdown ni espacio_fecha.markdown
        # El éxito se visualiza automáticamente en la Sidebar gracias a st.session_state['sincro_exitosa']
            
    else:
        # Si faltan datos, limpiamos todo para evitar inconsistencias
        if st.session_state.get('sincro_exitosa', True):
            st.session_state['sincro_exitosa'] = False
            st.session_state.pop('fecha_temporal_sincro', None)
            st.rerun()

except Exception as e:
    st.sidebar.error(f"Error en sincronización web: {e}")
