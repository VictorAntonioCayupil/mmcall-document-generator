"""
Template para la Generación de Manuales de MMCall Academy
Este script sirve como base para estructurar nuevos manuales en formato Word y PDF
respetando el diseño estándar y la identidad visual de la empresa.
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import sys

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

def create_document_base(output_docx_name, device_name, subtitle_text):
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
        
        p_head = cell_r.paragraphs[0]
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_head = p_head.add_run(f"MMCALL ACADEMY  |  Manual Técnico - {device_name}")
        r_head.font.name = 'Arial'
        r_head.font.size = Pt(8.5)
        r_head.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        
        # Pie de página (páginas 2+)
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_foot = p_foot.add_run("S&D Mmcall Paging Services LTDA.")
        r_foot.font.name = 'Arial'
        r_foot.font.size = Pt(8.5)
        r_foot.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        
    # Establecer fuente normal en Arial
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
    
    # Paleta de colores oficial
    c_primary = RGBColor(0x1F, 0x3A, 0x8A)  # Azul corporativo
    c_accent = RGBColor(0x9C, 0x89, 0xF5)   # Lila de MMCall Academy
    
    # --- PORTADA ---
    p_aca = doc.add_paragraph()
    p_aca.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_aca = p_aca.add_run("PROGRAMA DE CAPACITACIÓN ACADÉMICA")
    r_aca.font.name = 'Arial'
    r_aca.font.size = Pt(13)
    r_aca.font.bold = True
    r_aca.font.color.rgb = c_primary
    p_aca.paragraph_format.space_before = Pt(36)
    p_aca.paragraph_format.space_after = Pt(6)
    
    p_mmc = doc.add_paragraph()
    p_mmc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_mmc = p_mmc.add_run("MMCALL ACADEMY")
    r_mmc.font.name = 'Arial'
    r_mmc.font.size = Pt(22)
    r_mmc.font.bold = True
    r_mmc.font.color.rgb = c_accent
    p_mmc.paragraph_format.space_after = Pt(24)
    
    # Intentar cargar logotipo si existe en el CWD
    logo_path = "imagenes/logo_mmcall_nobg.png"
    if os.path.exists(logo_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_img = p_img.add_run()
        r_img.add_picture(logo_path, width=Inches(2.5))
        p_img.paragraph_format.space_after = Pt(24)
        
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"Manual Técnico - {device_name}")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = c_primary
    p_title.paragraph_format.space_after = Pt(6)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(subtitle_text)
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
    p_sub.paragraph_format.space_after = Pt(48)
    
    p_sello = doc.add_paragraph()
    p_sello.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sello = p_sello.add_run("Sello de Calidad MMCall Academy")
    r_sello.font.name = 'Arial'
    r_sello.font.size = Pt(11)
    r_sello.font.italic = True
    r_sello.font.bold = True
    r_sello.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)
    p_sello.paragraph_format.space_after = Pt(24)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Soporte y Desarrollo (S&D) - MMCall Paging Services LTDA. - Chile")
    r_meta.font.name = 'Arial'
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
    
    doc.add_page_break()
    return doc

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x8A)
    return p
    
def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x9C, 0x89, 0xF5)
    return p

def add_bullet(doc, bold_prefix, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    r_bold = p.add_run(bold_prefix)
    r_bold.font.bold = True
    p.add_run(text)
    return p

def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    p.add_run(text)
    return p

def add_image_centered(doc, img_path, caption):
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run()
        r.add_picture(img_path, width=Inches(4.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run(f"Figura: {caption}")
        r_cap.font.name = 'Arial'
        r_cap.font.size = Pt(9)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
    else:
        print(f"Warning: Imagen no encontrada: {img_path}")

def compile_to_pdf(docx_path, pdf_path):
    try:
        from docx2pdf import convert
        print(f"Converting {docx_path} to PDF...")
        convert(docx_path, pdf_path)
        print("Compilation complete!")
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
