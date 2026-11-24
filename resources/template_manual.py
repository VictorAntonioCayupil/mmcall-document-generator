"""
Template Oficial para la Generación de Manuales de MMCall Academy
Actualizado: Control de paginación natural (cero páginas vacías) y soporte para screenshots móviles.
"""

import os
import sys
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Paleta real MMCall Academy
C_RED = RGBColor(0xE3, 0x06, 0x13)          # Título principal, encabezado de tabla
C_LILA = RGBColor(0x9C, 0x89, 0xF5)         # "ACADEMY", eyebrow superior
C_GRAY_DARK = RGBColor(0x40, 0x40, 0x40)    # "MMCALL", encabezados de sección
C_GRAY_MED = RGBColor(0x7F, 0x8C, 0x8D)     # Pies de página / cabecera pages 2+
C_GRAY_LIGHT = RGBColor(0x80, 0x80, 0x80)   # Subtítulo de portada, sello
C_BODY = RGBColor(0x33, 0x33, 0x33)         # Texto de cuerpo
C_TABLE_ALT = "F2F2F2"                      # Fila alterna de tabla (hex)
C_TABLE_WHITE = "FFFFFF"

FONT_BODY = "Calibri"
FONT_EYEBROW = "Consolas"

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, color_hex):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))

def _brand_run(paragraph, text, color, size=None, bold=True):
    r = paragraph.add_run(text)
    r.font.name = FONT_BODY
    r.font.bold = bold
    r.font.color.rgb = color
    if size:
        r.font.size = size
    return r

def create_document_base(output_docx_name, device_name, subtitle_text, logo_path="imagenes/logo_mmcall_nobg.png"):
    doc = docx.Document()

    # 1. Configurar márgenes estándar (1.0 pulgada)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = True

        # Cabecera (páginas 2+)
        header = section.header
        header_table = header.add_table(1, 2, Inches(6.5))
        header_table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        cell_l = header_table.cell(0, 0)
        cell_r = header_table.cell(0, 1)

        p_head_l = cell_l.paragraphs[0]
        p_head_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_tit = p_head_l.add_run(device_name)
        r_tit.font.name = FONT_BODY
        r_tit.font.size = Pt(9)
        r_tit.font.color.rgb = C_GRAY_MED

        p_head_r = cell_r.paragraphs[0]
        p_head_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _brand_run(p_head_r, "MMCALL ", C_GRAY_DARK, Pt(9.5))
        _brand_run(p_head_r, "ACADEMY", C_LILA, Pt(9.5))

        # Pie de página (páginas 2+)
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_foot = p_foot.add_run("S&D MMCall Paging Services LTDA. · Manual Técnico")
        r_foot.font.name = FONT_BODY
        r_foot.font.size = Pt(8.5)
        r_foot.font.color.rgb = C_GRAY_MED

    # Fuente normal
    style_normal = doc.styles['Normal']
    style_normal.font.name = FONT_BODY
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = C_BODY

    # --- PORTADA OFICIAL MMCALL ACADEMY ---
    p_aca = doc.add_paragraph()
    p_aca.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_aca = p_aca.add_run("PROGRAMA DE CAPACITACIÓN ACADÉMICA")
    r_aca.font.name = FONT_EYEBROW
    r_aca.font.size = Pt(9)
    r_aca.font.bold = True
    r_aca.font.color.rgb = C_LILA
    p_aca.paragraph_format.space_before = Pt(28)
    p_aca.paragraph_format.space_after = Pt(4)

    p_mmc = doc.add_paragraph()
    p_mmc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _brand_run(p_mmc, "MMCALL ", C_GRAY_DARK, Pt(22))
    _brand_run(p_mmc, "ACADEMY", C_LILA, Pt(22))
    p_mmc.paragraph_format.space_after = Pt(18)

    if os.path.exists(logo_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_img = p_img.add_run()
        r_img.add_picture(logo_path, width=Inches(2.2))
        p_img.paragraph_format.space_after = Pt(18)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(device_name)
    r_title.font.name = FONT_BODY
    r_title.font.size = Pt(24)
    r_title.font.bold = True
    r_title.font.color.rgb = C_RED
    p_title.paragraph_format.space_after = Pt(6)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(subtitle_text)
    r_sub.font.name = FONT_BODY
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = C_GRAY_LIGHT
    p_sub.paragraph_format.space_after = Pt(36)

    p_sello = doc.add_paragraph()
    p_sello.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sello = p_sello.add_run("Sello de Calidad MMCall Academy")
    r_sello.font.name = FONT_BODY
    r_sello.font.size = Pt(10)
    r_sello.font.italic = True
    r_sello.font.bold = False
    r_sello.font.color.rgb = C_GRAY_LIGHT
    p_sello.paragraph_format.space_after = Pt(18)

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Soporte y Desarrollo (S&D) - MMCall Paging Services LTDA. - Chile")
    r_meta.font.name = FONT_BODY
    r_meta.font.size = Pt(9)
    r_meta.font.color.rgb = C_GRAY_MED

    # ÚNICO SALTO DE PÁGINA AUTORIZADO: Portada -> Página 2
    doc.add_page_break()

    # Repetición del título en rojo al iniciar la página 2
    p_rep = doc.add_paragraph()
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rep = p_rep.add_run(device_name)
    r_rep.font.name = FONT_BODY
    r_rep.font.size = Pt(16)
    r_rep.font.bold = True
    r_rep.font.color.rgb = C_RED
    p_rep.paragraph_format.space_after = Pt(12)

    return doc

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT_BODY
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = C_RED
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT_BODY
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = C_GRAY_DARK
    return p

def add_bullet(doc, bold_prefix, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.15
    r_bold = p.add_run(bold_prefix)
    r_bold.font.name = FONT_BODY
    r_bold.font.bold = True
    r_bold.font.color.rgb = C_BODY
    r_body = p.add_run(text)
    r_body.font.name = FONT_BODY
    r_body.font.color.rgb = C_BODY
    return p

def add_body(doc, text, bold_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = FONT_BODY
        r_b.font.bold = True
        r_b.font.color.rgb = C_BODY
    r = p.add_run(text)
    r.font.name = FONT_BODY
    r.font.color.rgb = C_BODY
    return p

def add_image_centered(doc, img_path, caption="", max_height_inches=3.2, max_width_inches=4.2):
    """Inserta una imagen centrada limitando la altura para evitar saltos y páginas vacías."""
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run()
        
        # Determinar si es captura vertical (móvil) u horizontal
        try:
            from PIL import Image
            with Image.open(img_path) as im:
                w, h = im.size
                if h > w * 1.2:
                    r.add_picture(img_path, height=Inches(max_height_inches))
                else:
                    r.add_picture(img_path, width=Inches(max_width_inches))
        except Exception:
            r.add_picture(img_path, height=Inches(max_height_inches))

        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(6)
            r_cap = p_cap.add_run(f"Figura: {caption}")
            r_cap.font.name = FONT_BODY
            r_cap.font.size = Pt(8.5)
            r_cap.font.italic = True
            r_cap.font.color.rgb = C_GRAY_MED
    else:
        print(f"Warning: Imagen no encontrada: {img_path}")

def add_step_card(doc, step_title, step_bullets, img_path, caption=""):
    """
    Inserta un paso técnico ilustrado en una tabla de 2 columnas sin bordes visibles:
    Columna izquierda: Título y viñetas explicativas.
    Columna derecha: Captura de pantalla móvil compacta (máx 3.0" de alto).
    Mantiene el paso y su imagen siempre juntos en la misma página.
    """
    table = doc.add_table(rows=1, cols=2)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    row = table.rows[0]
    prevent_row_split(row)
    cell_l, cell_r = row.cells[0], row.cells[1]
    cell_l.width = Inches(4.3)
    cell_r.width = Inches(2.2)
    
    set_cell_margins(cell_l, top=60, bottom=60, left=40, right=60)
    set_cell_margins(cell_r, top=60, bottom=60, left=60, right=40)
    
    # Texto a la izquierda
    p_t = cell_l.paragraphs[0]
    p_t.paragraph_format.space_before = Pt(2)
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run(step_title)
    r_t.font.name = FONT_BODY
    r_t.font.size = Pt(11)
    r_t.font.bold = True
    r_t.font.color.rgb = C_GRAY_DARK
    
    for b_prefix, b_text in step_bullets:
        p_b = cell_l.add_paragraph(style='List Bullet')
        p_b.paragraph_format.space_after = Pt(2)
        p_b.paragraph_format.line_spacing = 1.1
        if b_prefix:
            rb = p_b.add_run(b_prefix)
            rb.font.name = FONT_BODY
            rb.font.bold = True
            rb.font.size = Pt(9.5)
            rb.font.color.rgb = C_BODY
        rt = p_b.add_run(b_text)
        rt.font.name = FONT_BODY
        rt.font.size = Pt(9.5)
        rt.font.color.rgb = C_BODY

    # Imagen a la derecha
    if os.path.exists(img_path):
        p_img = cell_r.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(2)
        r_img = p_img.add_run()
        r_img.add_picture(img_path, height=Inches(2.8))
        
        if caption:
            p_cap = cell_r.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(2)
            rc = p_cap.add_run(caption)
            rc.font.name = FONT_BODY
            rc.font.size = Pt(8)
            rc.font.italic = True
            rc.font.color.rgb = C_GRAY_MED

    # Espacio tras la tabla
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(0)
    p_sep.paragraph_format.space_after = Pt(4)

def add_zebra_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Encabezado
    hdr_row = table.rows[0]
    prevent_row_split(hdr_row)
    hdr_cells = hdr_row.cells
    for i, h in enumerate(headers):
        set_cell_background(hdr_cells[i], "E30613")
        set_cell_margins(hdr_cells[i])
        p = hdr_cells[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(h)
        r.font.name = FONT_BODY
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if col_widths and i < len(col_widths):
            hdr_cells[i].width = Inches(col_widths[i])
            
    # Filas
    for ridx, row in enumerate(rows):
        row_obj = table.add_row()
        prevent_row_split(row_obj)
        cells = row_obj.cells
        fill = C_TABLE_ALT if ridx % 2 == 0 else C_TABLE_WHITE
        for c in cells:
            set_cell_background(c, fill)
        for i, val in enumerate(row):
            set_cell_margins(cells[i])
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(str(val))
            r.font.name = FONT_BODY
            r.font.size = Pt(9)
            r.font.color.rgb = C_BODY
            if col_widths and i < len(col_widths):
                cells[i].width = Inches(col_widths[i])
                
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(4)
    return table

def compile_to_pdf(docx_path, pdf_path):
    try:
        from docx2pdf import convert
        print(f"Compilando {docx_path} a PDF...")
        convert(str(docx_path), str(pdf_path))
        print("Compilacion completada con exito!")
    except Exception as e:
        print(f"Error al compilar PDF: {e}")
