# -*- coding: utf-8 -*-
"""
Generador Maestro de Guiones del Presentador ("Torpedos Internos") — MMCall Academy (S&D)
Produce los 3 documentos oficiales en formato Word (.docx) y PDF (.pdf):
1. 01_General_y_Ventas_ESL -> Guion_del_Presentador_General_y_Ventas_ESL.docx / .pdf
2. 02_Gateway_Minew_ESL -> Guion_del_Presentador_Gateway_Minew_ESL.docx / .pdf
3. 03_ESL_Modo_Standalone_Celular -> Guion_del_Presentador_ESL_Modo_Standalone.docx / .pdf
"""

import os
import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = Path(r"C:\Users\vcayu\Desktop\REPOSITORIO MMCALL ACADEMY\SyD\ESL")
DIR_01 = BASE_DIR / "01_General_y_Ventas_ESL"
DIR_02 = BASE_DIR / "02_Gateway_Minew_ESL"
DIR_03 = BASE_DIR / "03_ESL_Modo_Standalone_Celular"

# Paleta MMCall Academy Estándar Verificada
COLOR_RED = RGBColor(227, 6, 19)       # #E30613
COLOR_LILA = RGBColor(156, 137, 245)   # #9C89F5
COLOR_DARK = RGBColor(64, 64, 64)      # #404040
COLOR_MED = RGBColor(127, 140, 141)    # #7F8C8D
COLOR_LIGHT = RGBColor(128, 128, 128)  # #808080
COLOR_BODY = RGBColor(51, 51, 51)      # #333333

HEX_RED = "E30613"
HEX_ZEBRA = "F2F2F2"
HEX_BORDER = "CCCCCC"


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def format_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{HEX_RED}"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def format_card_table_borders(table, border_color=HEX_RED):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'  <w:left w:val="single" w:sz="16" w:space="0" w:color="{border_color}"/>'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'  <w:insideH w:val="none"/>'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def create_base_doc():
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(1.0)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(1.0)
        sec.different_first_page_header_footer = True
    return doc


def setup_header_footer(doc, module_name: str):
    header = doc.sections[0].header
    htbl = header.add_table(rows=1, cols=2, width=Inches(6.5))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = htbl.cell(0, 0)
    c_right = htbl.cell(0, 1)

    c_left.text = ""
    p_r = c_right.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    r1 = p_r.add_run("MMCALL ")
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_DARK

    r2 = p_r.add_run("ACADEMY")
    r2.font.name = "Calibri"
    r2.font.size = Pt(10)
    r2.font.bold = True
    r2.font.color.rgb = COLOR_LILA

    r3 = p_r.add_run(f"  |  Guion del Presentador · {module_name}")
    r3.font.name = "Calibri"
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = COLOR_MED

    # Footer
    footer = doc.sections[0].footer
    p_f = footer.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rf = p_f.add_run("Soporte y Desarrollo (S&D) — MMCall Paging Services LTDA. — Documento de Uso Interno")
    rf.font.name = "Calibri"
    rf.font.size = Pt(9)
    rf.font.color.rgb = COLOR_MED


def add_cover(doc, module_title: str, subtitle: str, logo_path: Path):
    # 1. Header Title Monospace
    p0 = doc.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(36)
    p0.paragraph_format.space_after = Pt(6)
    r0 = p0.add_run("GUION DEL PRESENTADOR · USO INTERNO")
    r0.font.name = "Consolas"
    r0.font.size = Pt(10)
    r0.font.bold = True
    r0.font.color.rgb = COLOR_LILA

    # 2. Brand line
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(14)
    r1_a = p1.add_run("MMCALL ")
    r1_a.font.name = "Calibri"
    r1_a.font.size = Pt(26)
    r1_a.font.bold = True
    r1_a.font.color.rgb = COLOR_DARK
    r1_b = p1.add_run("ACADEMY")
    r1_b.font.name = "Calibri"
    r1_b.font.size = Pt(26)
    r1_b.font.bold = True
    r1_b.font.color.rgb = COLOR_LILA

    # 3. Logo
    if logo_path.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(6)
        p_logo.paragraph_format.space_after = Pt(18)
        p_logo.add_run().add_picture(str(logo_path), width=Inches(2.4))

    # 4. Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(8)
    rt = p_title.add_run(f"Guion del Presentador\n«{module_title}»")
    rt.font.name = "Calibri"
    rt.font.size = Pt(24)
    rt.font.bold = True
    rt.font.color.rgb = COLOR_RED

    # 5. Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    rs = p_sub.add_run(subtitle)
    rs.font.name = "Calibri"
    rs.font.size = Pt(12)
    rs.font.color.rgb = COLOR_LIGHT

    # 6. Aviso de Confidencialidad
    p_warn = doc.add_paragraph()
    p_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_warn.paragraph_format.space_before = Pt(8)
    p_warn.paragraph_format.space_after = Pt(24)
    rw = p_warn.add_run("Documento interno para relatores y técnicos comerciales. Úsalo en tu celular o impreso. NO proyectar ni compartir con la audiencia.")
    rw.font.name = "Calibri"
    rw.font.size = Pt(10)
    rw.font.italic = True
    rw.font.color.rgb = COLOR_RED

    # 7. Sello de Calidad
    p_seal = doc.add_paragraph()
    p_seal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_seal.paragraph_format.space_before = Pt(14)
    p_seal.paragraph_format.space_after = Pt(4)
    r_seal = p_seal.add_run("Sello de Calidad MMCall Academy")
    r_seal.font.name = "Calibri"
    r_seal.font.size = Pt(10.5)
    r_seal.font.italic = True
    r_seal.font.color.rgb = COLOR_LIGHT

    # 8. Footer Portada
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(0)
    p_foot.paragraph_format.space_after = Pt(0)
    rf = p_foot.add_run("Soporte y Desarrollo (S&D) - MMCall Paging Services LTDA. - Chile")
    rf.font.name = "Calibri"
    rf.font.size = Pt(9.5)
    rf.font.color.rgb = COLOR_MED

    # ÚNICO SALTO DE PÁGINA PERMITIDO EN TODO EL DOCUMENTO
    doc.add_page_break()

    # Repetición de Título en Página 2
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(8)
    p2.paragraph_format.space_after = Pt(12)
    rp2 = p2.add_run(f"Guion Oficial del Presentador — {module_title}")
    rp2.font.name = "Calibri"
    rp2.font.size = Pt(16)
    rp2.font.bold = True
    rp2.font.color.rgb = COLOR_RED


def add_heading_1(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = COLOR_RED


def add_heading_2(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_DARK


def add_body_p(doc, text: str, italic=False, color=COLOR_BODY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.italic = italic
    r.font.color.rgb = color
    return p


def add_bullet(doc, prefix: str, body: str):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    r_pre = p.add_run(prefix + " ")
    r_pre.font.name = "Calibri"
    r_pre.font.size = Pt(10.5)
    r_pre.font.bold = True
    r_pre.font.color.rgb = COLOR_DARK

    r_bod = p.add_run(body)
    r_bod.font.name = "Calibri"
    r_bod.font.size = Pt(10.5)
    r_bod.font.color.rgb = COLOR_BODY


def add_parada_card(doc, parada_num: int, total_paradas: int, title: str, duration: str, time_accum: str, 
                    img_path: Path, que_se_ve: str, que_decir: str, que_hacer: str, mensaje_clave: str):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    format_card_table_borders(table, border_color=HEX_RED)

    c_img = table.cell(0, 0)
    c_txt = table.cell(0, 1)
    c_img.width = Inches(2.2)
    c_txt.width = Inches(4.3)
    set_cell_margins(c_img, top=100, bottom=100, left=100, right=100)
    set_cell_margins(c_txt, top=100, bottom=100, left=120, right=100)
    set_cell_shading(c_img, "F9FAFC")
    set_cell_shading(c_txt, "FFFFFF")

    # Columna izquierda: Imagen o Badge
    p_badge = c_img.paragraphs[0]
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_badge.paragraph_format.space_after = Pt(4)
    rb = p_badge.add_run(f"PARADA {parada_num} / {total_paradas}\n")
    rb.font.name = "Consolas"
    rb.font.size = Pt(9.5)
    rb.font.bold = True
    rb.font.color.rgb = COLOR_RED

    rtm = p_badge.add_run(f"{duration} · Minuto {time_accum}")
    rtm.font.name = "Calibri"
    rtm.font.size = Pt(8.5)
    rtm.font.color.rgb = COLOR_MED

    if img_path.exists():
        p_pic = c_img.add_paragraph()
        p_pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pic.paragraph_format.space_before = Pt(4)
        p_pic.paragraph_format.space_after = Pt(4)
        try:
            p_pic.add_run().add_picture(str(img_path), width=Inches(1.9))
        except Exception:
            pass

    # Columna derecha: Título y Contenido
    p_t = c_txt.paragraphs[0]
    p_t.paragraph_format.space_before = Pt(0)
    p_t.paragraph_format.space_after = Pt(4)
    r_title = p_t.add_run(f"{parada_num}. {title}")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(12)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_DARK

    def add_subitem(label: str, text: str, bold_label_color=COLOR_RED):
        p_item = c_txt.add_paragraph()
        p_item.paragraph_format.space_before = Pt(0)
        p_item.paragraph_format.space_after = Pt(3)
        p_item.paragraph_format.line_spacing = 1.15
        rl = p_item.add_run(label + ": ")
        rl.font.name = "Calibri"
        rl.font.size = Pt(9.5)
        rl.font.bold = True
        rl.font.color.rgb = bold_label_color
        rtx = p_item.add_run(text)
        rtx.font.name = "Calibri"
        rtx.font.size = Pt(9.5)
        rtx.font.color.rgb = COLOR_BODY

    add_subitem("Qué se ve", que_se_ve, COLOR_MED)
    add_subitem("Qué decir", que_decir, COLOR_RED)
    add_subitem("Qué hacer", que_hacer, COLOR_LILA)
    add_subitem("Mensaje clave", mensaje_clave, COLOR_DARK)

    # Espacio tras la tarjeta
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)


def convert_docx_to_pdf(out_docx: Path, out_pdf: Path):
    try:
        from docx2pdf import convert
        print(f"Compilando {out_docx.name} a PDF...")
        convert(str(out_docx), str(out_pdf))
        print(f"[OK] Torpedo PDF generado en: {out_pdf}")
    except Exception as e:
        print(f"[WARN] docx2pdf falló ({e}), intentando win32com...")
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            wdoc = word.Documents.Open(str(out_docx))
            wdoc.SaveAs(str(out_pdf), FileFormat=17)
            wdoc.Close()
            word.Quit()
            print(f"[OK] Torpedo PDF generado vía win32com en: {out_pdf}")
        except Exception as e2:
            print(f"[ERROR] No se pudo generar PDF: {e2}")


# =============================================================================
# 1. GUION GENERAL Y VENTAS ESL (01)
# =============================================================================

def build_guion_01():
    out_docx = DIR_01 / "Guion_del_Presentador_General_y_Ventas_ESL.docx"
    out_pdf = DIR_01 / "Guion_del_Presentador_General_y_Ventas_ESL.pdf"
    logo_file = DIR_01 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_01 / "imagenes"

    doc = create_base_doc()
    setup_header_footer(doc, "General y Ventas ESL")
    add_cover(
        doc,
        module_title="Ecosistema de Etiquetas Electrónicas (ESL)",
        subtitle="Recorrido guiado de venta y capacitación técnica: qué decir y qué accionar en cada parada de la presentación interactiva y PowerPoint.",
        logo_path=logo_file
    )

    add_heading_1(doc, "1. Antes de la Demostración Comercial")
    add_bullet(doc, "Copia de Seguridad:", "Asegúrate de tener la carpeta completa '01_General_y_Ventas_ESL'. El archivo presentacion_autocontenida.html contiene todo Base64 y funciona 100% offline.")
    add_bullet(doc, "Navegador Recomendado:", "Abre en Google Chrome o Microsoft Edge. Presiona la tecla 'F' o el botón de pantalla completa para maximizar.")
    add_bullet(doc, "Prueba Previa de Simuladores:", "Ensaya al menos una vez la rotación con mouse de los tags 3D en la parada 4, el cambio de precios en la góndola de la parada 5, y el movimiento de sliders en la calculadora de ROI de la parada 9.")
    add_bullet(doc, "Dispositivo de Apoyo:", "Abre este documento en tu teléfono celular o tenlo impreso. Los números de parada coinciden exactamente con los badges en pantalla.")

    add_heading_1(doc, "2. Cronología del Recorrido (Duración: 25 a 30 Minutos)")
    add_body_p(doc, "Versión comercial express: 15 minutos (enfocarse en paradas 1, 3, 5, 8 y 9, omitiendo detalles exhaustivos de accesorios).")

    # Tabla resumen de paradas
    tbl_crono = doc.add_table(rows=1, cols=4)
    tbl_crono.alignment = WD_TABLE_ALIGNMENT.CENTER
    format_table_borders(tbl_crono)
    headers = ["Parada", "Título / Módulo", "Tiempo", "Minuto"]
    widths = [Inches(1.0), Inches(3.4), Inches(1.0), Inches(1.1)]
    for i, h in enumerate(headers):
        cell = tbl_crono.cell(0, i)
        cell.width = widths[i]
        set_cell_margins(cell, 80, 80, 100, 100)
        set_cell_shading(cell, HEX_RED)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    crono_data = [
        ("1", "Portada Oficial y Visión General", "1.5 min", "0:00"),
        ("2", "Tecnología E-ink Bi-estable", "2.0 min", "1:30"),
        ("3", "Precios Dinámicos y Cumplimiento SERNAC", "2.5 min", "3:30"),
        ("4", "Visor 3D de Hardware (MTag / STag / DS / RS)", "3.0 min", "6:00"),
        ("5", "Simulador Interactivo de Góndola y Pick-by-Light", "3.5 min", "9:00"),
        ("6", "Accesorios y Sistemas de Montaje", "2.0 min", "12:30"),
        ("7", "Infraestructura: Gateway Minew G1-E y Enlace MQTT", "2.5 min", "14:30"),
        ("8", "Plataforma Minew Cloud y App Cloud Tag", "3.0 min", "17:00"),
        ("9", "Calculadora Interactiva de ROI y Ahorro", "3.5 min", "20:00"),
        ("10", "Evaluación de Conocimientos (Quiz) y Cierre", "3.0 min", "23:30"),
    ]

    for row_idx, data in enumerate(crono_data, 1):
        row = tbl_crono.add_row()
        bg_col = HEX_ZEBRA if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(data):
            c = row.cells[col_idx]
            c.width = widths[col_idx]
            set_cell_margins(c, 60, 60, 80, 80)
            set_cell_shading(c, bg_col)
            p = c.paragraphs[0]
            if col_idx in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_BODY

    add_heading_1(doc, "3. Guion Parada por Parada (Qué Decir y Qué Accionar)")

    paradas_01 = [
        {
            "num": 1, "tot": 10, "dur": "1.5 min", "time": "0:00",
            "title": "Portada Oficial y Visión General del Ecosistema",
            "img": img_dir / "image1.jpg",
            "que_se_ve": "Fondo holográfico interactivo con partículas animadas, logo oficial MMCall Academy y badge de capacitación.",
            "que_decir": "«Buenos días. Hoy revisaremos cómo el ecosistema de etiquetas electrónicas ESL de MMCall revoluciona la sala de venta. No solo reemplazamos flejes de papel: conectamos en tiempo real el ERP con la góndola, eliminamos las discrepancias de precio y habilitamos una experiencia omnicanal completa.»",
            "que_hacer": "Mantener la portada en pantalla completa. Mover el cursor para demostrar que las partículas reaccionan dinámicamente, transmitiendo alta tecnología.",
            "mensaje_clave": "MMCall no solo vende hardware, entrega una solución integral certificada con soporte local."
        },
        {
            "num": 2, "tot": 10, "dur": "2.0 min", "time": "1:30",
            "title": "Tecnología de Tinta Electrónica E-ink Bi-estable",
            "img": img_dir / "image5.png",
            "que_se_ve": "Diagrama de microcápsulas electroforéticas y comparativa frente a pantallas LCD tradicionales.",
            "que_decir": "«La clave de nuestras etiquetas es la tecnología E-ink bi-estable. A diferencia de un televisor o celular que consume energía para iluminar cada píxel, la tinta electrónica solo consume una milésima de miliamperio durante los dos segundos que tarda en refrescarse. Una vez congelado el precio, el consumo es estrictamente cero. Esto permite una autonomía de cinco a diez años con pilas botón estándar.»",
            "que_hacer": "Resaltar el ángulo de visión de 180 grados sin reflejos molestos frente a la iluminación artificial de supermercado.",
            "mensaje_clave": "Cero consumo en reposo, legibilidad idéntica al papel y vida útil de varios años."
        },
        {
            "num": 3, "tot": 10, "dur": "2.5 min", "time": "3:30",
            "title": "Precios Dinámicos y Cumplimiento SERNAC",
            "img": img_dir / "image6.png",
            "que_se_ve": "Módulos de sincronización POS vs. Góndola y estadísticas de errores en retail.",
            "que_decir": "«En Chile, la Ley del Consumidor sanciona severamente la discrepancia entre el precio de la góndola y el cobrado en caja. Un cambio manual de flejes toma horas y siempre deja errores humanos. Con MMCall, una modificación de precio en el sistema central se refleja en miles de etiquetas en segundos, garantizando coherencia absoluta y cero multas.»",
            "que_hacer": "Mencionar cómo los supermercados pierden hasta el 2% de su margen en ajustes de caja por reclamos de clientes.",
            "mensaje_clave": "Sincronización instantánea = Blindaje legal y protección del margen comercial."
        },
        {
            "num": 4, "tot": 10, "dur": "3.0 min", "time": "6:00",
            "title": "Visor 3D Interactivo de Hardware",
            "img": img_dir / "image12.png",
            "que_se_ve": "Tarjetas 3D con tilt giroscópico interactivo para las series MTag29R, STag21R, DS37R y RS75R.",
            "que_decir": "«Disponemos de cuatro familias de dispositivos: la serie MTag, que es el caballo de batalla del retail con display tricolor; la serie STag, con un perfil ultra delgado ideal para cosmética y joyería; la serie DS para cabeceras y gran formato; y la serie RS de grado industrial hermética IP67, que soporta congeladores a menos veinte grados bajo cero.»",
            "que_hacer": "Mover el mouse sobre cada tarjeta en el HTML interactivo para activar el reflejo luminoso holográfico y la rotación en perspectiva.",
            "mensaje_clave": "Un modelo exacto para cada categoría de producto y condición ambiental."
        },
        {
            "num": 5, "tot": 10, "dur": "3.5 min", "time": "9:00",
            "title": "Simulador Interactivo de Góndola y Pick-by-Light",
            "img": img_dir / "image14.png",
            "que_se_ve": "Góndola virtual de productos con botones para cambiar precios y encender luces LED.",
            "que_decir": "«Veamos esto en acción. Voy a pulsar el botón 'Cambiar Precio a Oferta Flash'. Observen cómo la pantalla parpadea y congela el nuevo valor con el fondo rojo de promoción. Ahora activaré el 'Pick-by-Light LED': la luz frontal enciende en color verde. Esto permite que el personal de picking en dark stores o eCommerce prepare pedidos un cuarenta por ciento más rápido sin equivocarse de producto.»",
            "que_hacer": "Hacer clic en los botones del simulador de góndola: 'Normal', 'Oferta Flash' y 'Picking LED' mostrando los tres estados en vivo.",
            "mensaje_clave": "Pick-by-Light transforma la etiqueta de un simple cartel a una herramienta de productividad logística."
        },
        {
            "num": 6, "tot": 10, "dur": "2.0 min", "time": "12:30",
            "title": "Accesorios y Sistemas Universales de Fijación",
            "img": img_dir / "image13.png",
            "que_se_ve": "Muestrario de rieles de policarbonato/aluminio, clips para canastos y bases magnéticas.",
            "que_decir": "«Un punto crítico en toda instalación es la compatibilidad mecánica. Disponemos de rieles modulares que encajan en las ranuras de cualquier estantería metálica o de madera, clips articulados para canastos de fruta, fijaciones magnéticas para racks industriales y atriles acrílicos para vitrinas y mesones.»",
            "que_hacer": "Si tienes muestras físicas en la reunión, pasa a los asistentes un riel con una etiqueta montada para que aprecien la firmeza del encaje.",
            "mensaje_clave": "Instalación limpia, robusta y modular en cualquier infraestructura existente."
        },
        {
            "num": 7, "tot": 10, "dur": "2.5 min", "time": "14:30",
            "title": "Infraestructura: Gateway Minew G1-E y Enlace Nube",
            "img": img_dir / "G1-E.png",
            "que_se_ve": "Diagrama de red con el Gateway central, túnel MQTT SSL y radio BLE 5.0 hacia las etiquetas.",
            "que_decir": "«El Gateway Minew G1-E es el concentrador de datos. Con un solo gateway se pueden cubrir hasta cincuenta metros a la redonda y administrar hasta cinco mil etiquetas. Se alimenta cómodamente por el mismo cable de red mediante PoE estándar, y toda la transmisión hacia la nube viaja cifrada con protocolo MQTT bajo certificados SSL.»",
            "que_hacer": "Destacar la función de reinicio automático y watchdog: si la red falla, el equipo conmuta automáticamente o se reanuda sin intervención técnica.",
            "mensaje_clave": "Infraestructura de nivel corporativo, segura, escalable y de mínimo cableado."
        },
        {
            "num": 8, "tot": 10, "dur": "3.0 min", "time": "17:00",
            "title": "Plataforma Cloud Minew y App Móvil Cloud Tag",
            "img": img_dir / "image22.png",
            "que_se_ve": "Panel web con estado de batería y señal RSSI, más la app móvil escaneando un tag.",
            "que_decir": "«La administración se realiza desde nuestra consola Cloud. Permite diseñar plantillas visuales con la tipografía y logo del cliente, monitorear el nivel de batería de cada tag y programar actualizaciones. En terreno, el operador usa la app móvil Cloud Tag: simplemente escanea el código de barras del producto y de la etiqueta, y quedan vinculados al instante.»",
            "que_hacer": "Explicar que la plataforma dispone de APIs abiertas REST y Webhooks para integrarse directamente a SAP, Oracle u Odoo.",
            "mensaje_clave": "Integración nativa con los sistemas centrales y operación móvil sumamente intuitiva."
        },
        {
            "num": 9, "tot": 10, "dur": "3.5 min", "time": "20:00",
            "title": "Calculadora Interactiva de ROI y Ahorro Económico",
            "img": img_dir / "image18.png",
            "que_se_ve": "Calculadora de retorno con sliders interactivos de cantidad de etiquetas, horas/hombre y costo de papel.",
            "que_decir": "«El valor más importante para el gerente de finanzas es el retorno de inversión. En esta calculadora interactiva podemos ingresar el número de productos de su tienda. Al simular dos cambios de precio por semana, el ahorro en horas de personal y eliminación de insumos de papel amortiza el proyecto completo en un plazo de ocho a catorce meses. A partir del primer año, todo es ganancia operativa neta.»",
            "que_hacer": "Mover los controles deslizantes de la calculadora de ROI en el HTML frente al cliente mostrando la cifra de ahorro en pesos chilenos.",
            "mensaje_clave": "El proyecto se autofinancia en menos de un año y libera cientos de horas de personal para vender."
        },
        {
            "num": 10, "tot": 10, "dur": "3.0 min", "time": "23:30",
            "title": "Evaluación de Conocimientos (Quiz) y Cierre",
            "img": img_dir / "image20.jpeg",
            "que_se_ve": "Módulo de evaluación interactiva con 8 preguntas técnicas y pantalla de diploma final.",
            "que_decir": "«Para cerrar esta sesión técnica, invitamos al equipo a validar los conceptos en nuestro quiz evaluativo. MMCall Academy entrega una certificación a cada participante que supera el ochenta por ciento de aprobación, asegurando que su personal cuenta con el conocimiento óptimo para operar el sistema.»",
            "que_hacer": "Resolver un par de preguntas en pantalla activando los sonidos de acierto con el Web Audio API y mostrar el diploma digital.",
            "mensaje_clave": "MMCall Academy respalda a sus clientes con capacitación formal, repuestos y soporte continuo en Chile."
        },
    ]

    for p_info in paradas_01:
        add_parada_card(
            doc,
            parada_num=p_info["num"], total_paradas=p_info["tot"],
            title=p_info["title"], duration=p_info["dur"], time_accum=p_info["time"],
            img_path=p_info["img"], que_se_ve=p_info["que_se_ve"],
            que_decir=p_info["que_decir"], que_hacer=p_info["que_hacer"],
            mensaje_clave=p_info["mensaje_clave"]
        )

    add_heading_1(doc, "4. Si Preguntan… (Manejo de Objeciones Comerciales y Técnicas)")
    add_bullet(doc, "¿Qué pasa si se corta la energía eléctrica en la tienda?", "Las pantallas E-ink son bi-estables: no requieren corriente para mantener la información. El precio, logo y código de barras quedan congelados exactamente igual aunque no haya luz durante semanas.")
    add_bullet(doc, "¿Cómo se reemplaza la batería y cuánto cuesta?", "Utilizan pilas botón de litio comerciales (CR2450 o similares). El cambio es sumamente sencillo abriendo la tapa posterior y tiene un costo insignificante cada 5 años.")
    add_bullet(doc, "¿Soporta integración con nuestro software de punto de venta?", "Sí. La plataforma Cloud dispone de una API REST documentada y Webhooks. Cuando el ERP cambia un precio, se dispara automáticamente la actualización a la etiqueta sin intervención humana.")
    add_bullet(doc, "¿Las etiquetas resisten golpes de carros de supermercado?", "Están fabricadas en policarbonato de alta resistencia e incluyen protectores acrílicos frontales en los rieles para absorber impactos directos.")
    add_bullet(doc, "¿Cuál es la garantía y tiempo de respuesta de MMCall?", "Contamos con laboratorio técnico propio en Chile, repuestos de recambio inmediato y garantía oficial de doce meses extensible.")

    add_heading_1(doc, "5. Si Algo Falla en Vivo (Plan de Contingencia)")
    add_bullet(doc, "Sin Conexión a Internet:", "La presentación presentacion_autocontenida.html funciona 100% desconectada. No depende de servidores externos ni de acceso a internet.")
    add_bullet(doc, "Proyector con Resolución Baja:", "Presiona la tecla '+' o '-' en el navegador para escalar la presentación sin que se desborden las tarjetas.")
    add_bullet(doc, "Pantalla Completa:", "Presiona la tecla 'F' para ocultar las barras del navegador y dar aspecto de aplicación nativa.")

    doc.save(str(out_docx))
    print(f"[OK] Guion 01 generado en: {out_docx}")
    convert_docx_to_pdf(out_docx, out_pdf)


# =============================================================================
# 2. GUION GATEWAY MINEW ESL (02)
# =============================================================================

def build_guion_02():
    out_docx = DIR_02 / "Guion_del_Presentador_Gateway_Minew_ESL.docx"
    out_pdf = DIR_02 / "Guion_del_Presentador_Gateway_Minew_ESL.pdf"
    logo_file = DIR_02 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_02 / "imagenes"

    doc = create_base_doc()
    setup_header_footer(doc, "Gateway Minew ESL")
    add_cover(
        doc,
        module_title="Configuración y Diagnóstico del Gateway Minew G1-E",
        subtitle="Torpedo técnico de soporte y puesta en marcha: qué verificar, comandos de consola, solución al conflicto Option 50 y enlace MQTT.",
        logo_path=logo_file
    )

    add_heading_1(doc, "1. Antes de Conectar el Gateway")
    add_bullet(doc, "Herramientas Necesarias:", "Computador portátil con puerto Wi-Fi y puerto Ethernet, cable de red RJ45 Cat5e o Cat6, y un inyector PoE 802.3af o cargador DC 5V 2A.")
    add_bullet(doc, "IPs por Defecto:", "En modo Access Point (Wi-Fi de servicio), la IP es 192.168.1.1. En modo Ethernet cliente, toma IP por DHCP.")
    add_bullet(doc, "Credenciales Oficiales:", "Usuario: admin | Contraseña oficial entregada por MMCall S&D.")

    add_heading_1(doc, "2. Cronología del Taller Técnico (20 Minutos)")
    crono_data_02 = [
        ("1", "Arquitectura Física, LEDs y Alimentación PoE", "2.0 min", "0:00"),
        ("2", "Conexión Inicial por Modo Access Point y Login Web", "2.5 min", "2:00"),
        ("3", "Configuración de Red Cableada Ethernet LAN", "2.0 min", "4:30"),
        ("4", "Configuración de Red Inalámbrica Wi-Fi Cliente", "2.5 min", "6:30"),
        ("5", "Diagnóstico Avanzado: Rechazo DHCP Option 50", "3.0 min", "9:00"),
        ("6", "Configuración del Servicio MQTT SSL con Minew Cloud", "2.5 min", "12:00"),
        ("7", "Estrategia de Failover y Watchdog Auto-Reboot", "2.0 min", "14:30"),
        ("8", "Vinculación en la Plataforma Cloud y Estado Online", "2.0 min", "16:30"),
        ("9", "Descarga de Logs, Actualización OTA y Respaldo", "1.5 min", "18:30"),
    ]

    tbl_crono = doc.add_table(rows=1, cols=4)
    tbl_crono.alignment = WD_TABLE_ALIGNMENT.CENTER
    format_table_borders(tbl_crono)
    headers = ["Parada", "Módulo Técnico", "Tiempo", "Minuto"]
    widths = [Inches(1.0), Inches(3.4), Inches(1.0), Inches(1.1)]
    for i, h in enumerate(headers):
        cell = tbl_crono.cell(0, i)
        cell.width = widths[i]
        set_cell_margins(cell, 80, 80, 100, 100)
        set_cell_shading(cell, HEX_RED)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, data in enumerate(crono_data_02, 1):
        row = tbl_crono.add_row()
        bg_col = HEX_ZEBRA if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(data):
            c = row.cells[col_idx]
            c.width = widths[col_idx]
            set_cell_margins(c, 60, 60, 80, 80)
            set_cell_shading(c, bg_col)
            p = c.paragraphs[0]
            if col_idx in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_BODY

    add_heading_1(doc, "3. Guion de Capacitación Paso a Paso")

    paradas_02 = [
        {
            "num": 1, "tot": 9, "dur": "2.0 min", "time": "0:00",
            "title": "Interfaces Físicas, LEDs y Alimentación PoE",
            "img": img_dir / "02_status_top.png",
            "que_se_ve": "Panel frontal y posterior del Gateway G1-E con puerto Ethernet, botón de reset y LEDs de PWR, LAN y BLE.",
            "que_decir": "«Iniciamos la revisión del hardware. El Gateway G1-E admite alimentación por cable de datos PoE estándar 802.3af hasta cien metros, eliminando transformadores en techos o pilares. Los LEDs frontales indican claramente: PWR fijo en verde para energía, LAN parpadeando para datos de red, y BLE activo para radiofrecuencia hacia las etiquetas.»",
            "que_hacer": "Mostrar el puerto RJ45 PoE y explicar que la polaridad es automática en switches comerciales.",
            "mensaje_clave": "Un único cable Ethernet resuelve datos y alimentación de forma industrial."
        },
        {
            "num": 2, "tot": 9, "dur": "2.5 min", "time": "2:00",
            "title": "Conexión por Modo Access Point y Login Web",
            "img": img_dir / "01_login.png",
            "que_se_ve": "Pantalla de autenticación web en http://192.168.1.1 con campos de usuario y contraseña.",
            "que_decir": "«Cuando el equipo no tiene red, emite su propio Wi-Fi de servicio con el nombre del modelo y su MAC. Nos conectamos desde el notebook, abrimos el navegador e ingresamos a 192.168.1.1 con las credenciales de administrador MMCall. Desde aquí controlamos todas las interfaces de comunicación.»",
            "que_hacer": "Recomendar usar una pestaña de incógnito para evitar que el navegador guarde en caché páginas web previas.",
            "mensaje_clave": "El modo AP permite configurar el equipo en terreno sin requerir routers intermedios."
        },
        {
            "num": 3, "tot": 9, "dur": "2.0 min", "time": "4:30",
            "title": "Configuración de Red Cableada: Ethernet LAN",
            "img": img_dir / "04_network_ethernet.png",
            "que_se_ve": "Menú Network > Ethernet con selección entre DHCP e IP Estática.",
            "que_decir": "«Para instalaciones definitivas en supermercados y tiendas, siempre recomendamos conectar por cable Ethernet. El modo DHCP solicita automáticamente la IP a la red del cliente. Si las políticas de seguridad exigen IP fija, configuramos IP estática, máscara, puerta de enlace y los DNS de la empresa.»",
            "que_hacer": "Advertir que al guardar cambios en modo estático, el notebook debe configurarse en el mismo segmento para no perder acceso.",
            "mensaje_clave": "Ethernet es el método con mayor estabilidad, menor jitter y máxima velocidad de transmisión."
        },
        {
            "num": 4, "tot": 9, "dur": "2.5 min", "time": "6:30",
            "title": "Configuración de Red Inalámbrica: Wi-Fi Cliente",
            "img": img_dir / "05_network_wireless_scan.png",
            "que_se_ve": "Módulo de escaneo inalámbrico mostrando SSIDs 2.4 GHz, canal, señal dBm y tipo de seguridad.",
            "que_decir": "«En locales donde tirar cable es imposible, asociamos el gateway al Wi-Fi del cliente. Presionamos 'Scan', seleccionamos el SSID del local, ingresamos la clave WPA2 o WPA3 Personal y pulsamos Save & Apply. El equipo se unirá a la red inalámbrica en unos cuarenta segundos.»",
            "que_hacer": "Verificar que el nivel de señal del SSID seleccionado supere los menos setenta dBm (-70 dBm) para asegurar estabilidad.",
            "mensaje_clave": "Wi-Fi 2.4 GHz permite despliegues rápidos en locales pequeños o zonas protegidas arquitectónicamente."
        },
        {
            "num": 5, "tot": 9, "dur": "3.0 min", "time": "9:00",
            "title": "Diagnóstico Avanzado: Rechazo DHCP Option 50",
            "img": img_dir / "07_network_wireless_form.png",
            "que_se_ve": "Formulario de conexión y logs de kernel con rechazos DHCPNAK.",
            "que_decir": "«Un hallazgo crítico documentado por nuestro equipo: si el router Wi-Fi del cliente tiene control estricto de lease y el gateway pide la IP anterior mediante la opción cincuenta de DHCP, el router puede rechazar la conexión con un DHCPNAK repetitivo. La solución técnica es reiniciar el cliente udhcpc forzando un DISCOVER limpio, o fijar una IP estática libre en el segmento para garantizar conexión inmediata.»",
            "que_hacer": "Explicar que este procedimiento resolvió el caso del router en terreno y debe aplicarse ante cualquier bucle de desconexión.",
            "mensaje_clave": "Conocimiento técnico especializado para resolver incompatibilidades de red complejas."
        },
        {
            "num": 6, "tot": 9, "dur": "2.5 min", "time": "12:00",
            "title": "Configuración del Servicio MQTT SSL con Minew Cloud",
            "img": img_dir / "14_service_mqtt_top.png",
            "que_se_ve": "Módulo Services > MQTT con host, puerto 8883, Client ID y tópicos de telemetría.",
            "que_decir": "«Una vez que el gateway tiene internet, configuramos el túnel MQTT. El host oficial es mqtt.minewesl.com en el puerto seguro 8883. El tráfico viaja completamente encriptado bajo certificados TLS, garantizando que nadie en la red local pueda interceptar los datos ni alterar los precios.»",
            "que_hacer": "Verificar que la opción SSL esté activada con el checkbox correspondiente antes de guardar.",
            "mensaje_clave": "Seguridad de grado bancario para la protección de la información de precios de la empresa."
        },
        {
            "num": 7, "tot": 9, "dur": "2.0 min", "time": "14:30",
            "title": "Estrategia de Failover y Watchdog Auto-Reboot",
            "img": img_dir / "21_other_auto_reboot_on.png",
            "que_se_ve": "Pantalla Other > Strategy con temporizador de reinicio programado y conmutación de enlace.",
            "que_decir": "«Para evitar visitas a terreno por bloqueos de memoria o cortes momentáneos de internet, activamos el Auto-Reboot preventivo cada veinticuatro horas en horario nocturno. Además, el equipo cuenta con watchdog: si pierde contacto con la nube por más de diez minutos, intenta reiniciar la interfaz inalámbrica de forma autónoma.»",
            "que_hacer": "Mostrar cómo configurar el reinicio a las 04:00 AM para no interrumpir la operación comercial.",
            "mensaje_clave": "Alta disponibilidad garantizada con mínimo mantenimiento en terreno."
        },
        {
            "num": 8, "tot": 9, "dur": "2.0 min", "time": "16:30",
            "title": "Vinculación en Minew Cloud y Validación Online",
            "img": img_dir / "26_cloud_gateway_online.png",
            "que_se_ve": "Consola web de Minew Cloud mostrando el gateway con punto verde de estado 'Online'.",
            "que_decir": "«Ingresamos a cloud.minewesl.com, vamos a Base Stations, agregamos la MAC del equipo y lo asociamos a la tienda activa. En menos de un minuto el indicador cambia de gris a verde brillante, confirmando que la nube ya tiene control total sobre el gateway y sus etiquetas.»",
            "que_hacer": "Comprobar que la hora del sistema y el contador de paquetes recibidos comiencen a incrementarse.",
            "mensaje_clave": "Verificación final exitosa: el gateway está en servicio operativo activo."
        },
        {
            "num": 9, "tot": 9, "dur": "1.5 min", "time": "18:30",
            "title": "Descarga de Logs, Actualización OTA y Respaldo",
            "img": img_dir / "24_other_logs_reboot.png",
            "que_se_ve": "Menú de descarga de logs (klog/syslog) y actualización de firmware .bin.",
            "que_decir": "«Ante cualquier comportamiento anómalo, descargamos el archivo de logs del sistema desde esta pestaña para auditar los eventos con el equipo de soporte. Asimismo, las actualizaciones de firmware OTA se aplican con un solo clic cargando el paquete oficial entregado por MMCall.»",
            "que_hacer": "Recordar nunca desconectar la alimentación eléctrica durante un proceso de actualización de firmware.",
            "mensaje_clave": "Trazabilidad completa y mantenimiento preventivo estandarizado."
        },
    ]

    for p_info in paradas_02:
        add_parada_card(
            doc,
            parada_num=p_info["num"], total_paradas=p_info["tot"],
            title=p_info["title"], duration=p_info["dur"], time_accum=p_info["time"],
            img_path=p_info["img"], que_se_ve=p_info["que_se_ve"],
            que_decir=p_info["que_decir"], que_hacer=p_info["que_hacer"],
            mensaje_clave=p_info["mensaje_clave"]
        )

    add_heading_1(doc, "4. Preguntas Frecuentes y Diagnóstico Rápido")
    add_bullet(doc, "¿Qué hacer si el gateway no toma IP por DHCP?", "Conectar directo al puerto Ethernet con IP fija en el notebook (192.168.1.100), ingresar a la web local y revisar en 'Status > Network' si el cable físico está negociando a 100 Mbps.")
    add_bullet(doc, "¿Por qué el indicador no pasa a Online en la Nube?", "Revisar que la puerta de enlace tenga salida hacia internet y que el firewall corporativo no tenga bloqueado el puerto saliente 8883 TCP.")
    add_bullet(doc, "¿Cómo se restablece de fábrica si se olvidó la clave?", "Mantener presionado el botón físico de Reset durante 10 segundos con el equipo encendido hasta que los tres LEDs parpadeen juntos.")

    doc.save(str(out_docx))
    print(f"[OK] Guion 02 generado en: {out_docx}")
    convert_docx_to_pdf(out_docx, out_pdf)


# =============================================================================
# 3. GUION ESL MODO STANDALONE (03)
# =============================================================================

def build_guion_03():
    out_docx = DIR_03 / "Guion_del_Presentador_ESL_Modo_Standalone.docx"
    out_pdf = DIR_03 / "Guion_del_Presentador_ESL_Modo_Standalone.pdf"
    logo_file = DIR_03 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_03 / "imagenes"

    doc = create_base_doc()
    setup_header_footer(doc, "ESL Modo StandAlone")
    add_cover(
        doc,
        module_title="Operación de Etiquetas ESL en Modo StandAlone",
        subtitle="Guía táctica para inducción operativa de terreno: qué decir y qué mostrar al personal de tienda que actualiza etiquetas 1 a 1 por Bluetooth 5.0.",
        logo_path=logo_file
    )

    add_heading_1(doc, "1. Antes de Iniciar la Capacitación en Tienda")
    add_bullet(doc, "Requisitos Móviles:", "Smartphone con Bluetooth 5.0 encendido, permisos de ubicación activos (requerido por el stack BLE de Android e iOS) y app Cloud Tag instalada.")
    add_bullet(doc, "Credenciales de Sucursal:", "Tener a mano el usuario y clave de la tienda activa en el servidor cloud.minewesl.com.")
    add_bullet(doc, "Etiquetas de Prueba:", "Disponer de al menos 2 etiquetas físicas con batería instalada (ej. MTag de 2.9 pulgadas) para hacer la demostración presencial.")

    add_heading_1(doc, "2. Cronología del Módulo (15 Minutos)")
    crono_data_03 = [
        ("1", "Concepto StandAlone: Cero Gateways, 100% Celular", "1.5 min", "0:00"),
        ("2", "Descarga de Cloud Tag y Permisos de Radio BLE", "2.0 min", "1:30"),
        ("3", "Inicio de Sesión y Módulo 'Actualización Independiente'", "2.0 min", "3:30"),
        ("4", "Identificación Óptica por Código de Barras Frontal", "2.0 min", "5:30"),
        ("5", "Identificación Inalámbrica por Baliza BLE y RSSI", "2.0 min", "7:30"),
        ("6", "Selección de Plantilla y Asociación de Producto", "2.0 min", "9:30"),
        ("7", "Modificación de Valores en Vivo (Teclado Móvil)", "2.0 min", "11:30"),
        ("8", "Inyección por Radiofrecuencia y Refresco E-ink", "2.0 min", "13:30"),
        ("9", "Auditoría en Góndola y Buenas Prácticas", "1.5 min", "15:30"),
    ]

    tbl_crono = doc.add_table(rows=1, cols=4)
    tbl_crono.alignment = WD_TABLE_ALIGNMENT.CENTER
    format_table_borders(tbl_crono)
    headers = ["Parada", "Paso Operativo", "Tiempo", "Minuto"]
    widths = [Inches(1.0), Inches(3.4), Inches(1.0), Inches(1.1)]
    for i, h in enumerate(headers):
        cell = tbl_crono.cell(0, i)
        cell.width = widths[i]
        set_cell_margins(cell, 80, 80, 100, 100)
        set_cell_shading(cell, HEX_RED)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, data in enumerate(crono_data_03, 1):
        row = tbl_crono.add_row()
        bg_col = HEX_ZEBRA if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(data):
            c = row.cells[col_idx]
            c.width = widths[col_idx]
            set_cell_margins(c, 60, 60, 80, 80)
            set_cell_shading(c, bg_col)
            p = c.paragraphs[0]
            if col_idx in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_BODY

    add_heading_1(doc, "3. Guion de Inducción Paso a Paso")

    paradas_03 = [
        {
            "num": 1, "tot": 9, "dur": "1.5 min", "time": "0:00",
            "title": "Concepto StandAlone: Cero Gateways, 100% Celular",
            "img": img_dir / "01_portada_standalone.png",
            "que_se_ve": "Smartphone comunicándose de forma inalámbrica y directa con una etiqueta electrónica.",
            "que_decir": "«Hoy aprenderemos a utilizar el modo StandAlone. ¿Qué significa? Que no necesitamos antenas, cables ni servidores en el local. Su propio teléfono celular se convierte en la antena transmisora para cambiar el precio de cualquier etiqueta de forma individual mediante Bluetooth.»",
            "que_hacer": "Mostrar una etiqueta suelta en la mano y el celular para evidenciar que no hay cables involucrados.",
            "mensaje_clave": "Máxima portabilidad para ferreterías, tiendas de paso y salas de reunión."
        },
        {
            "num": 2, "tot": 9, "dur": "2.0 min", "time": "1:30",
            "title": "Descarga de Cloud Tag y Permisos de Radio BLE",
            "img": img_dir / "02_qr_descarga_cloudtag.png",
            "que_se_ve": "Código QR para descarga de la app oficial en App Store y Google Play Store.",
            "que_decir": "«Para instalar la app oficial, escaneamos este código QR o buscamos 'Cloud Tag' de Minew en la tienda de aplicaciones. Al abrirla por primera vez, es indispensable aceptar los permisos de Bluetooth y Ubicación, ya que los sistemas operativos móviles exigen este permiso para explorar balizas de radiofrecuencia.»",
            "que_hacer": "Asegurarse de que todos los operarios en la sala tengan la app abierta con el Bluetooth activo.",
            "mensaje_clave": "Permisos de Bluetooth y Ubicación son obligatorios para detectar las etiquetas."
        },
        {
            "num": 3, "tot": 9, "dur": "2.0 min", "time": "3:30",
            "title": "Inicio de Sesión y Módulo 'Actualización Independiente'",
            "img": img_dir / "03_login_cloudtag.png",
            "que_se_ve": "Pantalla de login con usuario, contraseña y el botón 'Actualización independiente' en el menú.",
            "que_decir": "«Ingresamos el usuario y contraseña asignados por MMCall y presionamos Iniciar sesión. En la pantalla principal veremos el nombre de nuestra sucursal. Ahora pulsamos el botón azul que dice 'Actualización independiente'. Este es el módulo que nos permite trabajar uno a uno con cada etiqueta.»",
            "que_hacer": "Verificar que el host esté configurado en cloud.minewesl.com.",
            "mensaje_clave": "Acceso seguro vinculado a la cuenta corporativa de la empresa."
        },
        {
            "num": 4, "tot": 9, "dur": "2.0 min", "time": "5:30",
            "title": "Identificación Óptica por Código de Barras Frontal",
            "img": img_dir / "06_escaner_codigo_barra.png",
            "que_se_ve": "Cámara del celular encuadrando el código de barras impreso en el marco de la etiqueta.",
            "que_decir": "«El método más rápido y recomendado en el pasillo es el escáner con la cámara. Pulsamos el icono de escáner y apuntamos al código de barras impreso en el marco de la etiqueta. En menos de un segundo, la app lee la dirección MAC exacta sin que tengamos que escribir ningún número.»",
            "que_hacer": "Hacer la demostración en vivo apuntando la cámara a la etiqueta hasta escuchar el pitido de confirmación.",
            "mensaje_clave": "El escaneo óptico frontal elimina cualquier margen de error al elegir la etiqueta."
        },
        {
            "num": 5, "tot": 9, "dur": "2.0 min", "time": "7:30",
            "title": "Identificación Alternativa por Baliza BLE y RSSI",
            "img": img_dir / "07_lista_bluetooth_mac.png",
            "que_se_ve": "Lista de dispositivos Bluetooth cercanos ordenados por nivel de señal RSSI.",
            "que_decir": "«Si el código de barras estuviera sucio o dañado, usamos el método alternativo: exploración Bluetooth. La app lista todos los dispositivos cercanos. La etiqueta que tenemos pegada al celular mostrará una señal muy fuerte, cerca de menos cuarenta dBm. Cotejamos los últimos cuatro dígitos de la MAC serigrafiada atrás y la seleccionamos.»",
            "que_hacer": "Mostrar cómo la señal RSSI sube al acercar el celular a centímetros de la etiqueta.",
            "mensaje_clave": "Doble método de detección para asegurar continuidad ante cualquier eventualidad física."
        },
        {
            "num": 6, "tot": 9, "dur": "2.0 min", "time": "9:30",
            "title": "Selección de Plantilla y Asociación de Producto",
            "img": img_dir / "08_unir_etiqueta_plantilla.png",
            "que_se_ve": "Menú de plantillas por tamaño de pantalla (2.9 pulg tricolor) y catálogo de productos.",
            "que_decir": "«En esta pantalla seleccionamos la plantilla que corresponda al tamaño físico del tag, por ejemplo dos coma nueve pulgadas tricolor. Luego elegimos el producto de la lista oficial de la tienda o escribimos el nombre en el buscador.»",
            "que_hacer": "Recalcar la regla de oro: nunca seleccionar una plantilla de cuatro pulgadas en una etiqueta de dos pulgadas para evitar que el texto se corte.",
            "mensaje_clave": "Coincidencia estricta entre la plantilla elegida y el modelo físico del tag."
        },
        {
            "num": 7, "tot": 9, "dur": "2.0 min", "time": "11:30",
            "title": "Modificación de Valores en Vivo (Teclado Móvil)",
            "img": img_dir / "10_modificar_valores_en_vivo.png",
            "que_se_ve": "Pantalla de edición con campos de Nombre, Precio normal, Precio oferta y Código QR.",
            "que_decir": "«Si necesitamos cambiar el precio en este instante sin pasar por el computador, presionamos 'Modificar'. Con el teclado del teléfono editamos el valor, añadimos una palabra de oferta o cambiamos el enlace del código QR. La vista previa muestra exactamente cómo lucirá la pantalla antes de enviarla.»",
            "que_hacer": "Escribir un precio de prueba en vivo para que los asistentes vean la reactividad del formulario.",
            "mensaje_clave": "Flexibilidad total para promociones de último minuto en sala de venta."
        },
        {
            "num": 8, "tot": 9, "dur": "2.0 min", "time": "13:30",
            "title": "Inyección por Radiofrecuencia y Refresco E-ink",
            "img": img_dir / "13_transferencia_bluetooth_rf.png",
            "que_se_ve": "Botón azul 'Seguro' y pantalla de la etiqueta parpadeando durante la reorientación de pigmentos.",
            "que_decir": "«Con los datos revisados, sostenemos el celular a menos de medio metro de la etiqueta y presionamos el botón azul 'Seguro'. Durante tres segundos el celular envía los paquetes de imagen por Bluetooth. Observen la etiqueta: parpadea brevemente y congela el nuevo precio de forma permanente. ¡Operación completada!»",
            "que_hacer": "Sostener el tag frente a la audiencia para que todos vean el parpadeo físico y la nitidez final del texto.",
            "mensaje_clave": "Actualización inalámbrica instantánea con confirmación visual en segundos."
        },
        {
            "num": 9, "tot": 9, "dur": "1.5 min", "time": "15:30",
            "title": "Auditoría en Góndola y Buenas Prácticas",
            "img": img_dir / "15_etiqueta_actualizada_final.png",
            "que_se_ve": "Etiqueta montada en el riel de góndola exhibiendo el nuevo precio con nitidez absoluta.",
            "que_decir": "«Para finalizar, revisamos que los números sean perfectamente legibles y que el código QR apunte a la ficha correcta. La app queda inmediatamente lista para escanear la siguiente etiqueta. Con este método, una persona puede actualizar decenas de precios en un solo pasillo en pocos minutos.»",
            "que_hacer": "Invitar a un operario o asistente a hacer una etiqueta completa por sí mismo para fijar el aprendizaje.",
            "mensaje_clave": "Proceso ágil, seguro y completamente dominado por el personal de tienda."
        },
    ]

    for p_info in paradas_03:
        add_parada_card(
            doc,
            parada_num=p_info["num"], total_paradas=p_info["tot"],
            title=p_info["title"], duration=p_info["dur"], time_accum=p_info["time"],
            img_path=p_info["img"], que_se_ve=p_info["que_se_ve"],
            que_decir=p_info["que_decir"], que_hacer=p_info["que_hacer"],
            mensaje_clave=p_info["mensaje_clave"]
        )

    add_heading_1(doc, "4. Preguntas Frecuentes del Personal de Tienda")
    add_bullet(doc, "¿A qué distancia debo sostener el celular al presionar 'Seguro'?", "A menos de 50 centímetros de la etiqueta para asegurar que la transferencia Bluetooth no sufra interferencias.")
    add_bullet(doc, "¿Qué hago si la pantalla de la etiqueta no cambia?", "Verificar que el Bluetooth del teléfono siga encendido, que la plantilla corresponda a las pulgadas del tag y volver a presionar 'Seguro'.")
    add_bullet(doc, "¿Se borra la información si apago el teléfono?", "No. Una vez que la pantalla E-ink parpadea y congela la imagen, los pigmentos físicos quedan fijados permanentemente sin depender del teléfono ni de batería activa.")

    doc.save(str(out_docx))
    print(f"[OK] Guion 03 generado en: {out_docx}")
    convert_docx_to_pdf(out_docx, out_pdf)


def main():
    print("=== Generando los 3 Guiones del Presentador ('Torpedos Internos') ===")
    build_guion_01()
    build_guion_02()
    build_guion_03()
    print("=== [COMPLETO] Los 3 Guiones fueron generados y convertidos a PDF exitosamente ===")


if __name__ == "__main__":
    main()
