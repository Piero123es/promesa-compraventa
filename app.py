"""
Generador de "Promesa de Compraventa" a partir de plantilla.docx
-----------------------------------------------------------------
Versión Streamlit (compatible con Streamlit Cloud).

Requisitos (requirements.txt):
    streamlit
    docxtpl
    python-docx

IMPORTANTE:
- Este script usa `docxtpl`, no `python-docx`, para reemplazar las
  etiquetas Jinja de plantilla.docx ({{comprador}}, {{dni}},
  {% for cuota in cuotas %}...{% endfor %}, etc.).
- El archivo "plantilla.docx" debe estar en el mismo repositorio,
  junto a este app.py (o ajusta RUTA_PLANTILLA).
- Después de rellenar la plantilla, se fuerza que TODO el texto
  quede en fuente Aptos, tamaño 11 (cuerpo, tablas, encabezados,
  pies de página y el estilo "Normal").
"""

import os
import io
import streamlit as st
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.oxml.ns import qn

# -----------------------------
# Configuración general
# -----------------------------
st.set_page_config(page_title="Promesa de Compraventa", page_icon="📄", layout="centered")

RUTA_PLANTILLA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla.docx")

NOMBRE_FUENTE = "Aptos"
TAMANO_FUENTE = 11


# -----------------------------
# Utilidades para forzar fuente en TODO el documento
# -----------------------------
def _set_run_font(run, nombre_fuente, tamano_pt, mayusculas=True):
    run.font.name = nombre_fuente
    run.font.size = Pt(tamano_pt)
    run.font.all_caps = mayusculas
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    for atributo in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(atributo), nombre_fuente)


def _procesar_parrafos(parrafos, nombre_fuente, tamano_pt, mayusculas=True):
    for p in parrafos:
        for run in p.runs:
            _set_run_font(run, nombre_fuente, tamano_pt, mayusculas)


def _procesar_tablas(tablas, nombre_fuente, tamano_pt, mayusculas=True):
    for tabla in tablas:
        for fila in tabla.rows:
            for celda in fila.cells:
                _procesar_parrafos(celda.paragraphs, nombre_fuente, tamano_pt, mayusculas)
                _procesar_tablas(celda.tables, nombre_fuente, tamano_pt, mayusculas)  # tablas anidadas


def aplicar_fuente_global(documento, nombre_fuente=NOMBRE_FUENTE, tamano_pt=TAMANO_FUENTE, mayusculas=True):
    """Fuerza que TODO el texto del documento use la fuente y tamaño indicados,
    en MAYÚSCULAS (formato visual de Word, no cambia el texto real): cuerpo,
    tablas, encabezados y pies de página. También actualiza el estilo 'Normal'."""
    _procesar_parrafos(documento.paragraphs, nombre_fuente, tamano_pt, mayusculas)
    _procesar_tablas(documento.tables, nombre_fuente, tamano_pt, mayusculas)

    for section in documento.sections:
        partes = (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        )
        for parte in partes:
            _procesar_parrafos(parte.paragraphs, nombre_fuente, tamano_pt, mayusculas)
            _procesar_tablas(parte.tables, nombre_fuente, tamano_pt, mayusculas)

    estilo_normal = documento.styles["Normal"]
    estilo_normal.font.name = nombre_fuente
    estilo_normal.font.size = Pt(tamano_pt)
    estilo_normal.font.all_caps = mayusculas
    rPr = estilo_normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    for atributo in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(atributo), nombre_fuente)


# -----------------------------
# Estado de las cuotas (equivalente a las filas dinámicas de customtkinter)
# -----------------------------
if "cuotas" not in st.session_state:
    st.session_state.cuotas = [{"nombre": "", "monto": "", "monto_letras": "", "fecha": ""}]


def agregar_cuota():
    st.session_state.cuotas.append({"nombre": "", "monto": "", "monto_letras": "", "fecha": ""})


def eliminar_cuota(idx):
    if len(st.session_state.cuotas) <= 1:
        st.warning("Debe haber al menos una cuota.")
        return
    st.session_state.cuotas.pop(idx)


# -----------------------------
# Interfaz
# -----------------------------
st.title("📄 Nueva Promesa de Compraventa")

st.subheader("Fecha del documento")
col1, col2 = st.columns(2)
dia = col1.text_input("Día")
mes = col2.text_input("Mes")

st.subheader("Datos del Promitente Comprador(a)")
comprador = st.text_input("Nombre Completo")
dni = st.text_input("DNI")
celular = st.text_input("Celular")
nacionalidad = st.text_input("Nacionalidad")
estado_civil = st.text_input("Estado Civil")
ocupacion = st.text_input("Ocupación")
direccion = st.text_input("Dirección")
provincia = st.text_input("Provincia")
departamento = st.text_input("Departamento")

st.subheader("Datos del Terreno y Precio")
area = st.text_input("Área (M2)")
lote = st.text_input("N° de Lote")
precio = st.text_input("Precio total (S/)")
precio_letras = st.text_input("Precio total en letras")
monto_adicional = st.text_input("Monto para escritura pública (S/)")
monto_adicional_letras = st.text_input("Monto para escritura pública (en letras)")

st.subheader("Cuotas de pago")
for i, cuota in enumerate(st.session_state.cuotas):
    with st.container(border=True):
        c1, c2, c3, c4, c5 = st.columns([2, 1.3, 2.4, 1.3, 0.6])
        cuota["nombre"] = c1.text_input("Concepto", value=cuota["nombre"], key=f"nombre_{i}", placeholder="Ej: Inicial / Cuota 1")
        cuota["monto"] = c2.text_input("Monto S/", value=cuota["monto"], key=f"monto_{i}")
        cuota["monto_letras"] = c3.text_input("Monto en letras", value=cuota["monto_letras"], key=f"letras_{i}")
        cuota["fecha"] = c4.text_input("Fecha", value=cuota["fecha"], key=f"fecha_{i}", placeholder="dd/mm/aaaa")
        c5.write("")
        c5.button("✕", key=f"del_{i}", on_click=eliminar_cuota, args=(i,))

st.button("+ Agregar cuota", on_click=agregar_cuota)

st.subheader("Contacto Adicional")
adicional = st.text_input("Nombre del contacto adicional")
parentesco = st.text_input("Parentesco")
celular_adicional = st.text_input("Celular del contacto adicional")

st.subheader("Ejecutivo Comercial")
ejecutivo = st.text_input("Nombre del ejecutivo")
dni_ejecutivo = st.text_input("DNI del ejecutivo")

st.divider()

if st.button("Generar Word", type="primary", use_container_width=True):
    obligatorios = {
        "Nombre del comprador": comprador,
        "DNI": dni,
        "Día": dia,
        "Mes": mes,
        "Precio": precio,
    }
    faltantes = [nombre for nombre, valor in obligatorios.items() if not valor.strip()]

    if faltantes:
        st.error("Completa: " + ", ".join(faltantes))
    elif not os.path.exists(RUTA_PLANTILLA):
        st.error(
            f"No se encontró plantilla.docx en:\n{RUTA_PLANTILLA}\n\n"
            "Sube el archivo plantilla.docx al mismo repositorio que app.py."
        )
    else:
        contexto = {
            "dia": dia.strip(),
            "mes": mes.strip(),
            "comprador": comprador.strip(),
            "dni": dni.strip(),
            "celular": celular.strip(),
            "nacionalidad": nacionalidad.strip(),
            "ocupacion": ocupacion.strip(),
            "estado_civil": estado_civil.strip(),
            "direccion": direccion.strip(),
            "provincia": provincia.strip(),
            "departamento": departamento.strip(),
            "area": area.strip(),
            "lote": lote.strip(),
            "precio": precio.strip(),
            "precio_letras": precio_letras.strip(),
            "monto_adicional": monto_adicional.strip(),
            "monto_adicional_letras": monto_adicional_letras.strip(),
            "cuotas": [
                {k: v.strip() for k, v in c.items()} for c in st.session_state.cuotas
            ],
            "adicional": adicional.strip(),
            "parentesco": parentesco.strip(),
            "celular_adicional": celular_adicional.strip(),
            "ejecutivo": ejecutivo.strip(),
            "dni_ejecutivo": dni_ejecutivo.strip(),
        }

        try:
            doc = DocxTemplate(RUTA_PLANTILLA)
            doc.render(contexto)
            aplicar_fuente_global(doc.docx, NOMBRE_FUENTE, TAMANO_FUENTE)

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            nombre_sugerido = f"Promesa_Compraventa_{comprador.strip().replace(' ', '_') or 'documento'}.docx"

            st.success("Documento generado correctamente.")
            st.download_button(
                label="⬇️ Descargar documento",
                data=buffer,
                file_name=nombre_sugerido,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Ocurrió un error al generar el documento:\n{e}")
