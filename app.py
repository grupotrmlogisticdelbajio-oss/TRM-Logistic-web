import streamlit as st
import pandas as pd
import urllib.parse
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

# Configurar la página en modo ancho para aprovechar las dos columnas
st.set_page_config(layout="wide")

st.title("GRUPO TRM LOGISTIC - Cotizador con Códigos Postales y PDF")

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
def generar_pdf():
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    p.drawString(50, 750, "GRUPO TRM LOGISTIC - COTIZACIÓN DE FLETE")
    p.drawString(50, 730, f"Folio: {folio_str} | Fecha: {fecha_actual}")
    p.drawString(50, 700, f"Cliente: {cliente_empresa} - Atención: {cliente_contacto}")
    p.drawString(50, 680, f"Ruta: {ruta_sel} | Unidad: {unidad_sel}")
    p.drawString(50, 650, f"Subtotal: ${subtotal:,.2f}")
    p.drawString(50, 630, f"IVA (16%): ${iva:,.2f}")
    p.drawString(50, 610, f"Retención IVA (4%): ${retencion_iva:,.2f}")
    p.drawString(50, 590, f"TOTAL NETO: ${total_neto:,.2f}")
    p.showPage()
    p.save()
    buffer.seek(0)
    return buffer.getvalue()

pdf_data = generar_pdf()

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
