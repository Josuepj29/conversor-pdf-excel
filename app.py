import streamlit as st
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTChar
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Conversor Inteligente PDF a Excel", page_icon="📊", layout="centered")

# --- CSS AVANZADO (UI/UX) ---
st.markdown("""
<style>
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    
    .main .block-container {
        background: white;
        padding: 2.8rem;
        border-radius: 20px;
        box-shadow: 0 12px 32px rgba(15, 23, 42, 0.06);
        margin-top: 2rem;
        max-width: 850px;
        border: 1px solid #e2e8f0;
    }
    .stFileUploader > div > div {
        background-color: #f8fafc;
        border: 2px dashed #94a3b8;
        border-radius: 16px;
        padding: 28px;
        transition: all 0.25s ease;
    }
    .stFileUploader > div > div:hover {
        border-color: #2563eb;
        background-color: #eff6ff;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #1e293b, #0f172a);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 14px;
        font-size: 16px;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.18);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.28);
        color: #f8fafc;
    }
    .stDownloadButton>button {
        width: 100%;
        background: linear-gradient(135deg, #10b981, #059669);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 15px;
        font-size: 16px;
        font-weight: 700;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.25);
        transition: all 0.2s ease;
    }
    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 18px rgba(16, 185, 129, 0.35);
        color: white;
    }
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        background-color: #dbeafe;
        color: #1d4ed8;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# --- FUNCIONES DEL MOTOR DE EXTRACCIÓN POR COORDENADAS ---
def es_monto_numerico(val):
    return bool(re.match(r'^-?\d{1,3}(,\d{3})*(\.\d{2})$', str(val).strip()))

def convertir_valor(val):
    s = str(val).strip()
    if es_monto_numerico(s):
        return float(s.replace(',', ''))
    return s

def extraer_encabezado_superior(todas_las_filas, y_cabecera_tabla, num_cols):
    """Extrae el membrete de la página 1 (RUC, Cliente, Dirección, Cuenta, DOI, Oficina)
    usando el borde izquierdo (x0) para que quede alineado sin romper las columnas de la tabla."""
    filas_previas = [r for r in todas_las_filas if r['page'] == 0 and r['y'] > y_cabecera_tabla + 2.0]
    if not filas_previas:
        return []
        
    # Detectar párrafos de texto corrido (>330pt de ancho) para ignorar avisos legales ("Estimado cliente...")
    ys_parrafos = [
        r['y'] for r in filas_previas 
        if len(r['phrases']) == 1 and (r['phrases'][0][2] - r['phrases'][0][0]) > 330
    ]
    
    bloque_superior = []
    col_izq = 0
    col_etiq = min(4, max(1, num_cols - 3))
    col_val = min(5, max(2, num_cols - 2))
    
    for r in filas_previas:
        frases = r['phrases']
        y = r['y']
        if len(frases) == 1:
            ancho = frases[0][2] - frases[0][0]
            if ancho > 330 or any(abs(y - py) <= 18.0 for py in ys_parrafos):
                continue
                
        fila_out = [""] * num_cols
        for (x0, y0, x1, y1, txt) in frases:
            if x0 < 220:
                c_idx = col_izq
            elif x0 < 330:
                c_idx = max(1, col_etiq - 1)
            elif x0 < 395:
                c_idx = col_etiq
            else:
                c_idx = col_val
            fila_out[c_idx] = (fila_out[c_idx] + " " + txt).strip() if fila_out[c_idx] else txt
        if any(fila_out):
            bloque_superior.append(fila_out)
            
    return bloque_superior

def extraer_tablas_por_coordenadas(archivo_bytes):
    todas_las_filas = []
    
    # 1. Leer caracteres con sus coordenadas exactas (X, Y)
    for page_idx, page_layout in enumerate(extract_pages(io.BytesIO(archivo_bytes))):
        chars = []
        def recolectar_chars(container):
            for obj in container:
                if isinstance(obj, LTChar):
                    chars.append(obj)
                elif hasattr(obj, '__iter__'):
                    recolectar_chars(obj)
        recolectar_chars(page_layout)
        
        if not chars:
            continue
            
        chars.sort(key=lambda c: (-round(c.bbox[1], 1), c.bbox[0]))
        lineas = []
        for ch in chars:
            txt_ch = ch.get_text()
            if not txt_ch.strip() and txt_ch != ' ':
                continue
            ubicado = False
            for l in lineas:
                if abs(l['y'] - ch.bbox[1]) <= 3.0:
                    l['chars'].append(ch)
                    ubicado = True
                    break
            if not ubicado:
                lineas.append({'y': ch.bbox[1], 'chars': [ch]})
                
        # 2. Unir letras cercanas en frases indivisibles
        for l in lineas:
            l['chars'].sort(key=lambda c: c.bbox[0])
            frases = []
            cur_chars = []
            for ch in l['chars']:
                if not cur_chars:
                    cur_chars = [ch]
                else:
                    gap = ch.bbox[0] - cur_chars[-1].bbox[2]
                    if gap <= 6.5:
                        cur_chars.append(ch)
                    else:
                        txt = "".join(c.get_text() for c in cur_chars).strip()
                        if txt and "(cid:" not in txt:
                            frases.append((cur_chars[0].bbox[0], cur_chars[0].bbox[1], cur_chars[-1].bbox[2], cur_chars[-1].bbox[3], txt))
                        cur_chars = [ch]
            if cur_chars:
                txt = "".join(c.get_text() for c in cur_chars).strip()
                if txt and "(cid:" not in txt:
                    frases.append((cur_chars[0].bbox[0], cur_chars[0].bbox[1], cur_chars[-1].bbox[2], cur_chars[-1].bbox[3], txt))
            if frases:
                todas_las_filas.append({'page': page_idx, 'y': l['y'], 'phrases': frases})

    if not todas_las_filas:
        return None, None, []

    # 3. Aprender dónde están las columnas usando solo las filas densas (la tabla real)
    max_frases = max(len(r['phrases']) for r in todas_las_filas)
    min_denso = min(5, max(2, max_frases - 1))
    filas_tabla = [r['phrases'] for r in todas_las_filas if len(r['phrases']) >= min_denso]
    if not filas_tabla:
        filas_tabla = [r['phrases'] for r in todas_las_filas if len(r['phrases']) >= 2]
        
    max_x = int(max(c[2] for r in todas_las_filas for c in r['phrases'])) + 25
    cobertura_x = np.zeros(max_x, dtype=int)
    for r in filas_tabla:
        for (x0, y0, x1, y1, txt) in r:
            cobertura_x[max(0, int(np.floor(x0))):min(max_x, int(np.ceil(x1)) + 1)] += 1
            
    umbral = max(1, int(len(filas_tabla) * 0.03))
    es_columna = cobertura_x >= umbral
    
    intervalos_col = []
    en_col = False
    inicio = 0
    for x in range(max_x):
        if es_columna[x] and not en_col:
            en_col = True
            inicio = x
        elif not es_columna[x] and en_col:
            en_col = False
            if x - inicio >= 5:
                intervalos_col.append((inicio, x))

    if not intervalos_col:
        return None, None, []

    # 4. Asignar cada frase a su columna correspondiente
    filas_mapeadas = []
    for r in todas_las_filas:
        frases = r['phrases']
        if len(frases) == 1 and (frases[0][2] - frases[0][0]) > 350:
            continue
        celdas = [""] * len(intervalos_col)
        for (x0, y0, x1, y1, txt) in frases:
            cx = (x0 + x1) / 2.0
            mejor_col = min(
                range(len(intervalos_col)),
                key=lambda i: 0 if (intervalos_col[i][0] <= cx <= intervalos_col[i][1]) else min(abs(cx - intervalos_col[i][0]), abs(cx - intervalos_col[i][1]))
            )
            celdas[mejor_col] = (celdas[mejor_col] + " " + txt).strip() if celdas[mejor_col] else txt
        filas_mapeadas.append({'page': r['page'], 'y': r['y'], 'cells': celdas})

    # 5. Unir encabezados de tabla que vienen en 2 líneas (ej. FECHA arriba y OPER abajo)
    filas_unificadas = []
    i = 0
    while i < len(filas_mapeadas):
        cur = filas_mapeadas[i]
        tiene_num = any(es_monto_numerico(c) for c in cur['cells'])
        if not tiene_num and i + 1 < len(filas_mapeadas):
            grupo = [cur]
            j = i + 1
            while (j < len(filas_mapeadas) and 
                   filas_mapeadas[j]['page'] == cur['page'] and 
                   abs(grupo[-1]['y'] - filas_mapeadas[j]['y']) <= 5.0 and 
                   not any(es_monto_numerico(c) for c in filas_mapeadas[j]['cells'])):
                grupo.append(filas_mapeadas[j])
                j += 1
            if len(grupo) > 1 and sum(1 for g in grupo for c in g['cells'] if c) >= min_denso:
                combinadas = [""] * len(intervalos_col)
                for g in grupo:
                    for idx_c, val in enumerate(g['cells']):
                        if val:
                            combinadas[idx_c] = (combinadas[idx_c] + " " + val).strip() if combinadas[idx_c] else val
                filas_unificadas.append({'page': cur['page'], 'y': cur['y'], 'cells': combinadas, 'is_header': True})
                i = j
                continue
        cur['is_header'] = False
        filas_unificadas.append(cur)
        i += 1

    # 6. Separar "Tabla_Limpia", "Encabezado Superior" y "Documento_Completo"
    tabla_limpia = []
    header_agregado = False
    firma_header = None
    y_cabecera_detectada = None
    
    for r in filas_unificadas:
        celdas = r['cells']
        no_vacias = sum(1 for c in celdas if c != "")
        texto_fila = " ".join(celdas)
        if "Pág " in texto_fila or "Pag." in texto_fila:
            continue
        if r.get('is_header') or (not header_agregado and no_vacias >= min_denso and not any(es_monto_numerico(c) for c in celdas)):
            firma = tuple(celdas)
            if not header_agregado:
                tabla_limpia.append(celdas)
                header_agregado = True
                firma_header = firma
                y_cabecera_detectada = r['y']
            elif firma == firma_header:
                continue
        else:
            if no_vacias >= min_denso or (no_vacias >= 2 and es_monto_numerico(celdas[-1])):
                tabla_limpia.append([convertir_valor(c) for c in celdas])

    # Extraer el bloque superior de la página 1 sin afectar las columnas de la tabla
    bloque_encabezado = []
    if y_cabecera_detectada is not None:
        bloque_encabezado = extraer_encabezado_superior(todas_las_filas, y_cabecera_detectada, len(intervalos_col))

    if tabla_limpia and len(tabla_limpia) > 1:
        df_limpio = pd.DataFrame(tabla_limpia[1:], columns=tabla_limpia[0])
    else:
        df_limpio = pd.DataFrame([[convertir_valor(c) for c in r['cells']] for r in filas_unificadas])
        
    df_completo = pd.DataFrame([[convertir_valor(c) for c in r['cells']] for r in filas_unificadas])
    return df_limpio, df_completo, bloque_encabezado

# --- FUNCIÓN PARA CREAR EL EXCEL CON DISEÑO PROFESIONAL (OPENPYXL) ---
def generar_excel_estilizado(df_limpio, df_completo, bloque_encabezado, incluir_membrete=True):
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    meta_font = Font(bold=True, color="334155", size=9.5)
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    borde_fino = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    def estilizar_hoja(ws, df, tiene_cabecera=True, membrete=None):
         fila_actual = 1
         
         # 1. Si hay membrete superior, escribirlo arriba SIN combinar celdas
         if membrete:
             for fila_meta in membrete:
                 for col_idx, val in enumerate(fila_meta, 1):
                     if val != "":
                         c = ws.cell(row=fila_actual, column=col_idx, value=val)
                         c.font = meta_font
                 fila_actual += 1
             fila_actual += 1  # Dejar 1 fila en blanco de separación antes de la tabla
             
         fila_cabecera_tabla = fila_actual
         
         # 2. Escribir cabecera de la tabla
         if tiene_cabecera:
             for col_idx, col_name in enumerate(df.columns, 1):
                 cell = ws.cell(row=fila_actual, column=col_idx, value=str(col_name) if col_name != "" else f"COL_{col_idx}")
                 cell.fill = header_fill
                 cell.font = header_font
                 cell.alignment = Alignment(horizontal="center", vertical="center")
             fila_actual += 1
             
         # 3. Escribir datos de la tabla
         for row_data in df.values:
             for col_idx, val in enumerate(row_data, 1):
                 valor_celda = None if val == "" else val
                 cell = ws.cell(row=fila_actual, column=col_idx, value=valor_celda)
                 cell.border = borde_fino
                 if (fila_actual - fila_cabecera_tabla) % 2 == 1:
                     cell.fill = zebra_fill
                 if isinstance(valor_celda, float):
                     cell.number_format = '#,##0.00'
                     cell.alignment = Alignment(horizontal="right", vertical="center")
                 else:
                     cell.alignment = Alignment(vertical="center")
             fila_actual += 1

         # 4. Autoajustar ancho midiendo SOLO las filas de la tabla (para no deformar FECHA OPER)
         for col_idx in range(1, df.shape[1] + 1):
             max_len = max(
                 len(str(ws.cell(row=r, column=col_idx).value or ""))
                 for r in range(fila_cabecera_tabla, fila_actual)
             )
             col_letter = get_column_letter(col_idx)
             ws.column_dimensions[col_letter].width = max(12, min(max_len + 4, 48))

    # Hoja 1: Tabla Limpia (con o sin membrete según elija el usuario)
    ws1 = wb.active
    ws1.title = "Tabla_Limpia"
    estilizar_hoja(
        ws1, 
        df_limpio, 
        tiene_cabecera=True, 
        membrete=bloque_encabezado if incluir_membrete else None
    )

    # Hoja 2: Documento Completo
    ws2 = wb.create_sheet(title="Documento_Completo")
    estilizar_hoja(ws2, df_completo, tiene_cabecera=False, membrete=None)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

# --- INTERFAZ DE USUARIO ---
st.markdown("<div style='text-align: center;'><span class='badge'>⚡ CONVERSOR DE PDF A XLXS</span></div>", unsafe_allow_html=True)
st.markdown("<h1 style='text-align: center; color: #0f172a; font-weight: 800; margin-bottom: 0;'>Conversor PDF a EXCEL</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; font-size: 1.05rem; margin-top: 8px;'>Detecta columnas automáticamente. Exporta pdfs en formato xlsx.</p>", unsafe_allow_html=True)
st.markdown("<hr style='border: 1px solid #f1f5f9; margin: 1.8rem 0;'>", unsafe_allow_html=True)

archivo_pdf = st.file_uploader("📁 Sube o arrastra aquí tu estado de cuenta o reporte en PDF", type=["pdf"])
incluir_membrete = st.checkbox("📌 Incluir datos del encabezado (RUC, Cliente, Cuenta, CCI, Periodo) arriba de la tabla", value=True)

if archivo_pdf:
    if st.button("🚀 Estructurar y Convertir a Excel"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.markdown("<p style='text-align:center; color:#2563eb; font-weight:500;'>🔍 Mapeando coordenadas del pdf...</p>", unsafe_allow_html=True)
        for p in range(0, 45, 15):
            time.sleep(0.08)
            progress_bar.progress(p)
            
        try:
            bytes_pdf = archivo_pdf.read()
            df_limpio, df_completo, bloque_encabezado = extraer_tablas_por_coordenadas(bytes_pdf)
            
            status_text.markdown("<p style='text-align:center; color:#2563eb; font-weight:500;'>🎨 Aplicando formato xlxs y autoajustando columnas...</p>", unsafe_allow_html=True)
            for p in range(45, 95, 15):
                time.sleep(0.08)
                progress_bar.progress(p)
                
            if df_limpio is not None and not df_limpio.empty:
                excel_bytes = generar_excel_estilizado(df_limpio, df_completo, bloque_encabezado, incluir_membrete)
                progress_bar.progress(100)
                time.sleep(0.15)
                status_text.empty()
                progress_bar.empty()
                
                st.success(f"✅ ¡Conversión exitosa! Se detectaron **{df_limpio.shape[1]} columnas** y **{df_limpio.shape[0]} filas** limpias.")
                
                nombre_base = archivo_pdf.name.rsplit('.', 1)[0]
                st.download_button(
                    label="📥 Descargar Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{nombre_base}_Limpio.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                tab1, tab2 = st.tabs(["📊 Vista Previa: Tabla Limpia", "📄 Vista Previa: Encabezado"])
                with tab1:
                    st.dataframe(df_limpio.head(15), use_container_width=True)
                with tab2:
                    if bloque_encabezado:
                        st.dataframe(pd.DataFrame(bloque_encabezado), use_container_width=True)
                    else:
                        st.info("No se detectó un bloque de membrete superior en este documento.")
            else:
                progress_bar.empty()
                status_text.empty()
                st.error("⚠️ No se encontró texto seleccionable en el PDF (es posible que sea una imagen escaneada).")
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"Error al procesar el archivo: {e}")