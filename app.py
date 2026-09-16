from weasyprint import HTML
import streamlit as st
import pandas as pd
import urllib.parse
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

# Configurar la página en modo ancho para aprovechar las dos columnas
st.set_page_config(layout="wide")

st.title("GRUPO TRM LOGISTIC - Cotizador")

excel_file = "TRM_Cotizador_.xlsm"

# Leer la hoja de Excel omitiendo el encabezado superior
df_raw = pd.read_excel(excel_file, sheet_name="Rutas", skiprows=1)

def limpiar_numero(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('$', '').replace(',', '').strip()
    try:
        return float(s)
    except:
        return 0.0

# Crear DataFrame de Rutas de forma limpia y segura
df_rutas = pd.DataFrame()
df_rutas['Ruta'] = df_raw.iloc[:, 0].astype(str).str.strip()
df_rutas['Costo Base'] = df_raw.iloc[:, 1].apply(limpiar_numero)
df_rutas['Distancia'] = df_raw.iloc[:, 2].apply(limpiar_numero)
df_rutas['Casetas Base'] = df_raw.iloc[:, 3].apply(limpiar_numero)

df_rutas = df_rutas.dropna(subset=['Ruta'])
df_rutas = df_rutas[df_rutas['Ruta'] != 'nan']
df_rutas = df_rutas[df_rutas['Ruta'] != '']

# Crear DataFrame de Unidades de forma limpia y segura
df_unidades = pd.DataFrame()
df_unidades['Tipo de Unidad'] = df_raw.iloc[:, 4].astype(str).str.strip()
df_unidades['Mult. Flete'] = df_raw.iloc[:, 5].apply(limpiar_numero)
df_unidades['Rendimiento'] = df_raw.iloc[:, 6].apply(limpiar_numero)
df_unidades['Mult. Casetas'] = df_raw.iloc[:, 7].apply(limpiar_numero)

df_unidades = df_unidades.dropna(subset=['Tipo de Unidad'])
df_unidades = df_unidades[df_unidades['Tipo de Unidad'] != 'nan']
df_unidades = df_unidades[df_unidades['Tipo de Unidad'] != '']

# Función inteligente basada en los dos primeros dígitos del CP en México
def buscar_ruta_por_cp(cp):
    cp = cp.strip()
    if len(cp) < 2:
        return None
    
    prefijo = cp[:2]
    
    # Mapeo de prefijos postales de México a palabras clave de tus rutas
    if prefijo in ["64", "65", "66", "67"]:
        return "Monterrey"
    elif prefijo == "76":
        return "Querétaro"
    elif prefijo == "37":
        return "León"
    elif prefijo in ["44", "45", "46", "47", "48"]:
        return "Guadalajara"
    elif prefijo == "78":
        return "San Luis Potosi"
    elif prefijo in ["50", "51", "52", "53", "54", "55", "56", "57"]:
        return "Tecamac" # O Estado de México / CDMX según tus nombres de ruta
    elif prefijo in ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16"]:
        return "Ciudad de Mexico"
    
    return None

# Inicializar Folio en session_state si no existe
if "folio_num" not in st.session_state:
    st.session_state.folio_num = 1401

# Dividir la interfaz en dos columnas (Izquierda: Cotizador | Derecha: Mapa)
col_izq, col_der = st.columns([1.1, 0.9])

with col_izq:
    st.header("1. Datos del Envío y Cliente")

    # Campos de datos del cliente limpios para captura libre
    cliente_empresa = st.text_input("Empresa / Cliente:", "", placeholder="Ej. Empresa SA de CV", key="trm_cte")
    cliente_contacto = st.text_input("Contacto:", "", placeholder="Ej. Juan Pérez", key="trm_cont")
    cliente_correo = st.text_input("Correo:", "", placeholder="Ej. correo@empresa.com", key="trm_email")
    cliente_telefono_w = st.text_input("Teléfono WhatsApp (10 dígitos):", "", placeholder="Ej. 4611234567", key="trm_tel")

    st.markdown("---")
    
    # Opción de búsqueda por Código Postal o Selección Manual
    modo_seleccion = st.radio("¿Cómo deseas capturar el destino?", ["Seleccionar Ruta de la lista", "Buscar por Código Postal (CP)"], index=0, key="trm_modo_dest")

    rutas_list = sorted(df_rutas['Ruta'].unique().tolist())
    
    if modo_seleccion == "Buscar por Código Postal (CP)":
        cp_ingresado = st.text_input("Ingresa el Código Postal de destino (5 dígitos):", "", max_chars=5, placeholder="Ej. 66634", key="trm_input_cp")
        
        ruta_sugerida = ""
        if len(cp_ingresado) == 5:
            palabra_clave = buscar_ruta_por_cp(cp_ingresado)
            if palabra_clave:
                coincidencias = [r for r in rutas_list if palabra_clave.lower() in r.lower()]
                if coincidencias:
                    ruta_sugerida = coincidencias[0]
                    st.success(f"✅ CP {cp_ingresado} detectado ({palabra_clave}). Ruta asignada: **{ruta_sugerida}**")
                else:
                    st.warning(f"⚠️ Identificamos la zona ({palabra_clave}), pero no hay una ruta exacta con ese nombre en el Excel. Selecciona abajo.")
            else:
                st.info("ℹ️ Código postal no reconocido en el catálogo automático. Selecciona la ruta manualmente.")

        if ruta_sugerida and ruta_sugerida in rutas_list:
            indice_default = rutas_list.index(ruta_sugerida)
        else:
            indice_default = 0

        ruta_sel = st.selectbox("Ruta Asociada:", rutas_list, index=indice_default, key="trm_ruta_cp")
    else:
        ruta_sel = st.selectbox("Seleccionar Ruta:", rutas_list, key="trm_ruta")

    unidades_list = sorted(df_unidades['Tipo de Unidad'].unique().tolist())
    unidad_sel = st.selectbox("Tipo de Unidad:", unidades_list, key="trm_unidad")

    tipo_combustible = st.radio("Tipo de Combustible:", ["Gasolina ($23.68)", "Diésel ($27.03)"], index=0, key="trm_combustible")
    if "Diésel" in tipo_combustible:
        precio_litro = 27.03
    else:
        precio_litro = 23.68

    combustible_ida_vuelta = st.radio("Combustible y caseta ida y vuelta:", ["Sí", "No"], index=0, key="trm_iv")
    gastos_operativos = 1500.0

    tipo_viaje = st.radio("Tipo de Viaje:", ["Sencillo", "Redondo"], key="trm_tipo_viaje")

    precio_manual = st.number_input("PRECIO MANUAL (OPCIONAL) ($):", value=0.0, step=100.0, format="%.2f", key="trm_precio_manual")

    # Separar origen y destino de la ruta de manera limpia
    if "-" in ruta_sel:
        partes_ruta = ruta_sel.split("-")
        origen_str = partes_ruta[0].strip()
        destino_str = partes_ruta[1].strip()
    else:
        origen_str = "Celaya"
        destino_str = ruta_sel

    # URLs limpias de navegación para Google Maps especificando México
    url_maps = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(origen_str + ', Guanajuato, Mexico')}&destination={urllib.parse.quote(destino_str + ', Mexico')}"
    
    st.markdown(f'<a href="{url_maps}" target="_blank"><button style="background-color: #4CAF50; color: white; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-bottom: 15px;">🗺️ Abrir Ruta Exacta en Google Maps</button></a>', unsafe_allow_html=True)

    # Obtener datos exactos de la ruta seleccionada
    fila_ruta = df_rutas[df_rutas['Ruta'] == ruta_sel]
    if not fila_ruta.empty:
        costo_base_ruta = float(fila_ruta.iloc[0]['Costo Base'])
        distancia_base = float(fila_ruta.iloc[0]['Distancia'])
        casetas_base_ruta = float(fila_ruta.iloc[0]['Casetas Base'])
    else:
        costo_base_ruta, distancia_base, casetas_base_ruta = 0.0, 0.0, 0.0

    fila_unidad = df_unidades[df_unidades['Tipo de Unidad'] == unidad_sel]
    if not fila_unidad.empty:
        mult_flete = float(fila_unidad.iloc[0]['Mult. Flete'])
        rendimiento = float(fila_unidad.iloc[0]['Rendimiento'])
        mult_casetas = float(fila_unidad.iloc[0]['Mult. Casetas'])
    else:
        mult_flete, rendimiento, mult_casetas = 1.0, 1.0, 1.0

    factor_viaje = 2.0 if combustible_ida_vuelta == "Sí" else 1.0
    distancia_total = distancia_base * factor_viaje
    flete_base_dinamico = costo_base_ruta * mult_flete
    consumo_combustible_litros = distancia_total / rendimiento if rendimiento > 0 else 0
    costo_combustible_total = consumo_combustible_litros * precio_litro
    casetas_total = casetas_base_ruta * mult_casetas * factor_viaje

    costo_viaje_redondo = ((flete_base_dinamico + costo_combustible_total + casetas_total + gastos_operativos) * 0.5) if tipo_viaje == "Redondo" else 0.0

subtotal_calculado = flete_base_dinamico + costo_combustible_total + casetas_total + gastos_operativos + costo_viaje_redondo
subtotal = precio_manual if precio_manual > 0 else subtotal_calculado
iva = subtotal * 0.16
retencion_iva = subtotal * 0.04
total_neto = subtotal + iva - retencion_iva

st.markdown("---")
st.header("2. Desglose de Cotización")
st.write(f"**Ruta Detectada:** {ruta_sel} | Costo Base: ${costo_base_ruta:,.2f} | Distancia Base: {distancia_base:,.1f} km")
st.write(f"**Unidad Detectada:** {unidad_sel} | Mult Flete: {mult_flete} | Rendimiento: {rendimiento}")
st.write(f"**SUBTOTAL APLICADO:** ${subtotal:,.2f}")
st.write(f"**IVA (16%):** ${iva:,.2f}")
st.write(f"**IVA de retención (4%):** ${retencion_iva:,.2f}")
st.write(f"**TOTAL NETO:** ${total_neto:,.2f}")

with col_der:
    st.header("📍 Visualización de Ruta")
    st.write(f"Mostrando ruta de: **{origen_str}** a **{destino_str}**")
    url_embed = f"https://maps.google.com/maps?saddr={urllib.parse.quote(origen_str + ', Guanajuato, Mexico')}&daddr={urllib.parse.quote(destino_str + ', Mexico')}&t=&z=8&ie=UTF8&output=embed"
    st.markdown(f'<iframe src="{url_embed}" width="100%" height="600" style="border:0; border-radius: 8px;" allowfullscreen="" loading="lazy"></iframe>', unsafe_allow_html=True)

# Sección de Formato TRM y Acciones (PDF, Nuevo Folio, WhatsApp)
st.markdown("---")
st.header("3. Formato TRM - Vista Previa y Acciones")

fecha_actual = datetime.now().strftime("%d/%m/%Y")
folio_str = f"TRM-2026-{st.session_state.folio_num}"

st.info(f"**COTIZACIÓN** | Fecha: {fecha_actual} | Folio: **{folio_str}**")

col_f1, col_f2 = st.columns(2)
with col_f1:
    st.write(f"**Empresa:** {cliente_empresa if cliente_empresa else '(Sin especificar)'}")
    st.write(f"**Contacto:** {cliente_contacto if cliente_contacto else '(Sin especificar)'} ({cliente_correo})")
    st.write(f"**Origen / Destino:** {ruta_sel}")
with col_f2:
    st.write(f"**Tipo de Unidad:** {unidad_sel}")
    st.write(f"**Importe Subtotal:** ${subtotal:,.2f}")
    st.write(f"**Total Neto (con IVA y Retención):** ${total_neto:,.2f}")

# Función para generar PDF en memoria
def generar_pdf_cotizacion(
    folio,
    fecha,
    cliente,
    contacto,
    correo,
    validez,
    ruta_sel,
    unidad_sel,
    subtotal,
    iva,
    retencion_iva,
    total_neto,
    tipo_viaje,
):
  html_content = f"""<!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4 landscape;
                margin: 10mm 12mm;
                background-color: #ffffff;
            }}
            * {{
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }}
            body {{
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                color: #222;
                font-size: 11pt;
                line-height: 1.3;
            }}
            .header-container {{
                display: table;
                width: 100%;
                border-bottom: 2px solid #1a365d;
                padding-bottom: 10px;
                margin-bottom: 12px;
            }}
            .header-left {{
                display: table-cell;
                width: 40%;
                vertical-align: middle;
            }}
            .company-title {{
                font-size: 18pt;
                font-weight: bold;
                color: #1a365d;
                letter-spacing: 0.5px;
            }}
            .company-subtitle {{
                font-size: 9pt;
                color: #555;
                text-transform: uppercase;
                letter-spacing: 1px;
                margin-top: 3px;
            }}
            .header-center {{
                display: table-cell;
                width: 25%;
                text-align: center;
                vertical-align: middle;
            }}
            .logo-box {{
                font-size: 16pt;
                font-weight: 900;
                color: #d97706;
                letter-spacing: 1px;
            }}
            .logo-sub {{
                font-size: 6.5pt;
                color: #555;
                text-transform: uppercase;
            }}
            .header-right {{
                display: table-cell;
                width: 35%;
                text-align: right;
                vertical-align: middle;
            }}
            .cotizacion-title {{
                font-size: 18pt;
                font-weight: bold;
                color: #1a365d;
                text-transform: uppercase;
            }}
            .meta-table {{
                width: 100%;
                font-size: 9pt;
                margin-top: 4px;
            }}
            .meta-table td {{
                padding: 1px 0;
            }}
            .meta-label {{
                text-align: right;
                color: #555;
                padding-right: 8px;
            }}
            .meta-value {{
                font-weight: bold;
                color: #111;
                text-align: right;
            }}
            
            .section-title {{
                background-color: #1a365d;
                color: white;
                font-size: 9.5pt;
                font-weight: bold;
                padding: 4px 8px;
                margin-top: 10px;
                margin-bottom: 6px;
                text-transform: uppercase;
            }}

            .info-grid {{
                display: table;
                width: 100%;
                margin-bottom: 10px;
                font-size: 9.5pt;
            }}
            .info-col {{
                display: table-cell;
                width: 50%;
                vertical-align: top;
                padding-right: 15px;
            }}
            .info-row {{
                display: table;
                width: 100%;
                margin-bottom: 3px;
            }}
            .info-key {{
                display: table-cell;
                width: 35%;
                color: #555;
                font-weight: 600;
            }}
            .info-val {{
                display: table-cell;
                width: 65%;
                color: #111;
            }}

            .items-table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 8px;
                margin-bottom: 15px;
            }}
            .items-table th {{
                background-color: #1a365d;
                color: white;
                font-size: 9pt;
                text-transform: uppercase;
                padding: 6px 8px;
                text-align: left;
            }}
            .items-table th.right, .items-table td.right {{
                text-align: right;
            }}
            .items-table td {{
                padding: 7px 8px;
                font-size: 9.5pt;
                border-bottom: 1px solid #e2e8f0;
            }}

            .totals-container {{
                width: 100%;
                margin-bottom: 15px;
            }}
            .totals-table {{
                width: 320px;
                margin-left: auto;
                border-collapse: collapse;
                font-size: 10pt;
            }}
            .totals-table td {{
                padding: 4px 8px;
            }}
            .totals-label {{
                text-align: right;
                color: #333;
                font-weight: 600;
            }}
            .totals-val {{
                text-align: right;
                font-weight: bold;
                color: #111;
            }}
            .total-neto-row td {{
                border-top: 2px solid #1a365d;
                border-bottom: 2px solid #1a365d;
                background-color: #f1f5f9;
                font-size: 11pt;
                color: #1a365d;
                padding: 6px 8px;
            }}

            .terms-section {{
                border-top: 1px solid #cbd5e1;
                padding-top: 8px;
                margin-top: 10px;
            }}
            .terms-title {{
                font-size: 9pt;
                font-weight: bold;
                color: #1a365d;
                margin-bottom: 4px;
                text-transform: uppercase;
            }}
            .terms-list {{
                font-size: 8pt;
                color: #475569;
                line-align: 1.4;
            }}
            .terms-list li {{
                margin-left: 15px;
                margin-bottom: 2px;
            }}
        </style>
    </head>
    <body>

        <div class="header-container">
            <div class="header-left">
                <div class="company-title">GRUPO TRM LOGISTIC</div>
                <div class="company-subtitle">Autotransporte Federal & Logística</div>
            </div>
<div class="header-center">
    <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAQDAwQDAwQEAwQFBAQFBgoHBgYGBg0JCggKDw0QEA8NDw4RExgUERIXEg4PFRwVFxkZGxsbEBQdHx0aHxgaGxr/2wBDAQQFBQYFBgwHBwwaEQ8RGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhr/wgARCATpBQADASIAAhEBAxEB/8QAHAABAAEFAQEAAAAAAAAAAAAAAAECBQYHCAQD/8QAGwEBAAIDAQEAAAAAAAAAAAAAAAEFAgQGAwf/2gAMAwEAAhADEAAAAd/gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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" style="max-height: 60px; width: auto;">
</div>                
<div class="cotizacion-title">Cotización</div>=
            <div class="header-right">
                <table class="meta-table">
                    <tr>
                        <td class="meta-label">Fecha:</td>
                        <td class="meta-value">{fecha}</td>
                    </tr>
                    <tr>
                        <td class="meta-label">Folio:</td>
                        <td class="meta-value">{folio}</td>
                    </tr>
                    <tr>
                        <td class="meta-label">Teléfono:</td>
                        <td class="meta-value">4612324099</td>
                    </tr>
                </table>
            </div>
        </div>

        <div class="section-title">Datos del Cliente</div>
        <div class="info-grid">
            <div class="info-col">
                <div class="info-row">
                    <div class="info-key">Empresa / Cliente:</div>
                    <div class="info-val">{cliente}</div>
                </div>
                <div class="info-row">
                    <div class="info-key">Contacto:</div>
                    <div class="info-val">{contacto}</div>
                </div>
                <div class="info-row">
                    <div class="info-key">Correo:</div>
                    <div class="info-val">{correo}</div>
                </div>
                <div class="info-row">
                    <div class="info-key">Validez:</div>
                    <div class="info-val">{validez}</div>
                </div>
            </div>
            <div class="info-col">
                <div class="info-row">
                    <div class="info-key">Atención a Cliente:</div>
                    <div class="info-val">Atención a Cliente General / Departamento de Logistica</div>
                </div>
                <div class="info-row">
                    <div class="info-key">Ruta:</div>
                    <div class="info-val">{ruta_sel}</div>
                </div>
                <div class="info-row">
                    <div class="info-key">Tipo de Viaje:</div>
                    <div class="info-val">{tipo_viaje}</div>
                </div>
            </div>
        </div>

        <div class="section-title">Especificaciones del Servicio</div>
        
        <table class="items-table">
            <thead>
                <tr>
                    <th style="width: 8%;">Item</th>
                    <th style="width: 32%;">Descripción del Servicio</th>
                    <th style="width: 25%;">Ruta / Detalles</th>
                    <th style="width: 20%;">Tipo de Unidad</th>
                    <th class="right" style="width: 15%;">Importe</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>01</td>
                    <td>Servicio de Flete Terrestre de Carga</td>
                    <td>{ruta_sel}</td>
                    <td>{unidad_sel}</td>
                    <td class="right">${subtotal:,.2f}</td>
                </tr>
            </tbody>
        </table>

        <div class="totals-container">
            <table class="totals-table">
                <tr>
                    <td class="totals-label">Subtotal:</td>
                    <td class="totals-val">${subtotal:,.2f}</td>
                </tr>
                <tr>
                    <td class="totals-label">IVA (16%):</td>
                    <td class="totals-val">${iva:,.2f}</td>
                </tr>
                <tr>
                    <td class="totals-label">Retención IVA (4%):</td>
                    <td class="totals-val">-${retencion_iva:,.2f}</td>
                </tr>
                <tr class="total-neto-row">
                    <td class="totals-label">TOTAL NETO:</td>
                    <td class="totals-val">${total_neto:,.2f}</td>
                </tr>
            </table>
        </div>

        <div class="terms-section">
            <div class="terms-title">Condiciones Comerciales y Términos del Servicio</div>
            <ul class="terms-list">
                <li>Precios incluyen IVA</li>
                <li>Libre de Maniobras: Servicio libre de maniobras de carga y descarga (a cargo del cliente).</li>
                <li>Tiempos de Carga/Descarga: Incluye 3 hrs libres de carga y 3 hrs libres de descarga; tiempo extra genera estadía.</li>
                <li>Seguro de Mercancía: La carga viaja por cuenta y riesgo del cliente salvo contratación explícita de póliza.</li>
                <li>Condiciones de pago: Banco: BBVA BANCOMER Cuenta Cable: 012215001270598098 / Num de cuenta: 0127059809.</li>
            </ul>
        </div>

    </body>
    </html>
    """

  pdf_bytes = HTML(string=html_content).write_pdf()
  return pdf_bytes

pdf_data = generar_pdf_cotizacion(
    folio=folio_str,
    fecha=fecha_actual,
    cliente=cliente_empresa,
    contacto=cliente_contacto,
    correo=cliente_correo,
    validez="7 Días Naturales",
    ruta_sel=ruta_sel,
    unidad_sel=unidad_sel,
    subtotal=subtotal,
    iva=iva,
    retencion_iva=retencion_iva,
    total_neto=total_neto,
    tipo_viaje=tipo_viaje,
)

# Botones de Acción
col_b1, col_b2, col_b3 = st.columns(3)

with col_b1:
    st.download_button(
        label="📄 Guardar PDF",
        data=pdf_data,
        file_name=f"Cotizacion_{folio_str}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

with col_b2:
    if st.button("➕ NUEVO FOLIO", use_container_width=True):
        st.session_state.folio_num += 1
        st.rerun()

with col_b3:
    tel_limpio = cliente_telefono_w.strip() if cliente_telefono_w else "4610000000"
    mensaje_w = f"Hola {cliente_contacto if cliente_contacto else 'Estimado cliente'}, le compartimos la cotización {folio_str} para la ruta {ruta_sel} con unidad {unidad_sel}. Total Neto: ${total_neto:,.2f}. Quedamos a sus órdenes."
    mensaje_encoded = urllib.parse.quote(mensaje_w)
    url_whatsapp = f"https://wa.me/52{tel_limpio}?text={mensaje_encoded}"
    st.markdown(f'<a href="{url_whatsapp}" target="_blank"><button style="background-color: #25D366; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; width: 100%; font-weight: bold;">💬 Enviar WhatsApp</button></a>', unsafe_allow_html=True)
