# -*- coding: utf-8 -*-
"""
Generador Maestro de Presentaciones PowerPoint (.pptx) — MMCall Academy (S&D)
Genera las 3 presentaciones oficiales con el estándar MMCall Dark Studio (16:9 Widescreen):
1. 01_General_y_Ventas_ESL -> Presentacion_Oficial_General_y_Ventas_ESL.pptx
2. 02_Gateway_Minew_ESL -> Presentacion_Oficial_Gateway_Minew_ESL.pptx
3. 03_ESL_Modo_Standalone_Celular -> Presentacion_Oficial_ESL_Modo_Standalone.pptx
"""

import os
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = Path(r"C:\Users\vcayu\Desktop\REPOSITORIO MMCALL ACADEMY\SyD\ESL")
DIR_01 = BASE_DIR / "01_General_y_Ventas_ESL"
DIR_02 = BASE_DIR / "02_Gateway_Minew_ESL"
DIR_03 = BASE_DIR / "03_ESL_Modo_Standalone_Celular"

# Colores Oficiales MMCall Academy
C_BG_RGB = RGBColor(7, 10, 16)
C_CARD_RGB = RGBColor(16, 22, 36)
C_CARD_HI_RGB = RGBColor(22, 30, 48)
C_RED_RGB = RGBColor(227, 6, 19)
C_PURPLE_RGB = RGBColor(156, 137, 245)
C_CYAN_RGB = RGBColor(0, 240, 255)
C_GREEN_RGB = RGBColor(16, 185, 129)
C_AMBER_RGB = RGBColor(245, 158, 11)
C_WHITE_RGB = RGBColor(248, 250, 252)
C_MUTED_RGB = RGBColor(148, 163, 184)
C_BORDER_RGB = RGBColor(52, 65, 92)

SLIDE_W = 13.333
SLIDE_H = 7.5


def create_deck(title_module: str):
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def add_slide_background(slide):
    # Fondo completo
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_RGB
    bg.line.color.rgb = C_BG_RGB
    return bg


def add_top_bar(slide):
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_W), Inches(0.08))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = C_RED_RGB
    top_bar.line.color.rgb = C_RED_RGB


def add_header(slide, logo_path: Path, eyebrow: str, title: str, subtitle: str, slide_num: int, total_slides: int):
    add_top_bar(slide)

    x_logo = 0.5
    y_logo = 0.35
    w_logo = 1.6
    h_logo = 0.45

    if logo_path.exists():
        try:
            slide.shapes.add_picture(str(logo_path), Inches(x_logo), Inches(y_logo), width=Inches(w_logo))
            x_text = 2.3
        except Exception:
            x_text = 0.5
    else:
        x_text = 0.5

    # Eyebrow + Title box
    tb = slide.shapes.add_textbox(Inches(x_text), Inches(0.28), Inches(8.5), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_eye = tf.paragraphs[0]
    p_eye.text = eyebrow.upper()
    p_eye.font.name = "Outfit"
    p_eye.font.size = Pt(11)
    p_eye.font.bold = True
    p_eye.font.color.rgb = C_PURPLE_RGB

    p_title = tf.add_paragraph()
    p_title.text = title
    p_title.font.name = "Outfit"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_WHITE_RGB

    # Subtitle
    tb_sub = slide.shapes.add_textbox(Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = subtitle
    p_sub.font.name = "Calibri"
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = C_MUTED_RGB

    # Línea divisoria
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.58), Inches(12.333), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = C_BORDER_RGB
    line.line.color.rgb = C_BORDER_RGB

    # Badge de Diapositiva
    badge_w = 1.6
    badge_h = 0.42
    badge_x = SLIDE_W - 0.5 - badge_w
    badge_y = 0.38
    b_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(badge_x), Inches(badge_y), Inches(badge_w), Inches(badge_h))
    b_shp.fill.solid()
    b_shp.fill.fore_color.rgb = C_CARD_HI_RGB
    b_shp.line.color.rgb = C_PURPLE_RGB
    b_shp.line.width = Pt(1.5)
    b_tf = b_shp.text_frame
    b_tf.margin_left = b_tf.margin_top = b_tf.margin_right = b_tf.margin_bottom = 0
    bp = b_tf.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    bp.text = f"{slide_num} / {total_slides}"
    bp.font.name = "Consolas"
    bp.font.size = Pt(13)
    bp.font.bold = True
    bp.font.color.rgb = C_PURPLE_RGB


def add_footer(slide, module_info: str):
    foot_h = 0.45
    foot_y = SLIDE_H - foot_h
    bg_foot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(foot_y), Inches(SLIDE_W), Inches(foot_h))
    bg_foot.fill.solid()
    bg_foot.fill.fore_color.rgb = RGBColor(4, 6, 11)
    bg_foot.line.color.rgb = C_BORDER_RGB
    bg_foot.line.width = Pt(0.5)

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(foot_y + 0.1), Inches(12.333), Inches(0.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"MMCall Academy — Programa de Capacitación Técnica Oficial (S&D)   |   {module_info}"
    p.font.name = "Calibri"
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_MUTED_RGB


def add_bullets_card(slide, x, y, w, h, card_title: str, bullets, accent_color=C_RED_RGB):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD_RGB
    card.line.color.rgb = C_BORDER_RGB
    card.line.width = Pt(1.5)

    # Accent top strip inside card
    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.2), Inches(y + 0.18), Inches(1.8), Inches(0.04))
    strip.fill.solid()
    strip.fill.fore_color.rgb = accent_color
    strip.line.color.rgb = accent_color

    tb = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.28), Inches(w - 0.4), Inches(h - 0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_ct = tf.paragraphs[0]
    p_ct.text = card_title.upper()
    p_ct.font.name = "Outfit"
    p_ct.font.size = Pt(13)
    p_ct.font.bold = True
    p_ct.font.color.rgb = accent_color
    p_ct.space_after = Pt(10)

    for prefix, body in bullets:
        p = tf.add_paragraph()
        p.space_after = Pt(8)
        
        # Bullet indicator
        r_pre = p.add_run()
        r_pre.text = f"• {prefix} "
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(12)
        r_pre.font.bold = True
        r_pre.font.color.rgb = C_WHITE_RGB

        r_bod = p.add_run()
        r_bod.text = body
        r_bod.font.name = "Calibri"
        r_bod.font.size = Pt(11.5)
        r_bod.font.color.rgb = C_MUTED_RGB


def add_image_card(slide, x, y, w, h, img_path: Path, caption: str = ""):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD_RGB
    card.line.color.rgb = C_BORDER_RGB
    card.line.width = Pt(1.5)

    pad = 0.15
    cap_h = 0.35 if caption else 0
    img_max_w = w - pad * 2
    img_max_h = h - pad * 2 - cap_h

    if img_path.exists():
        try:
            # Añadir imagen preservando aspecto dentro del contenedor
            slide.shapes.add_picture(
                str(img_path),
                Inches(x + pad),
                Inches(y + pad),
                width=Inches(img_max_w),
                height=Inches(img_max_h),
            )
        except Exception as e:
            tb = slide.shapes.add_textbox(Inches(x + pad), Inches(y + pad), Inches(img_max_w), Inches(img_max_h))
            tb.text_frame.text = f"Imagen no disponible: {img_path.name}"
    else:
        tb = slide.shapes.add_textbox(Inches(x + pad), Inches(y + pad), Inches(img_max_w), Inches(img_max_h))
        tb.text_frame.text = f"Archivo no encontrado: {img_path.name}"

    if caption:
        cap_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.05), Inches(y + h - cap_h - 0.05), Inches(w - 0.1), Inches(cap_h))
        cap_bar.fill.solid()
        cap_bar.fill.fore_color.rgb = RGBColor(11, 15, 25)
        cap_bar.line.color.rgb = C_BORDER_RGB
        cap_bar.line.width = Pt(0.5)

        tb_cap = slide.shapes.add_textbox(Inches(x + 0.1), Inches(y + h - cap_h - 0.02), Inches(w - 0.2), Inches(cap_h))
        tf_c = tb_cap.text_frame
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        pc = tf_c.paragraphs[0]
        pc.alignment = PP_ALIGN.CENTER
        pc.text = caption.upper()
        pc.font.name = "Consolas"
        pc.font.size = Pt(10.5)
        pc.font.bold = True
        pc.font.color.rgb = C_CYAN_RGB


def add_cover_slide(prs, title: str, subtitle: str, badge_text: str, logo_path: Path, hero_img_path: Path = None):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    add_slide_background(slide)
    add_top_bar(slide)

    if logo_path.exists():
        try:
            slide.shapes.add_picture(str(logo_path), Inches(1.0), Inches(0.9), width=Inches(2.8))
        except Exception:
            pass

    # Badge de portada
    b_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.2), Inches(3.6), Inches(0.48))
    b_shp.fill.solid()
    b_shp.fill.fore_color.rgb = C_CARD_HI_RGB
    b_shp.line.color.rgb = C_RED_RGB
    b_shp.line.width = Pt(1.5)
    bp = b_shp.text_frame.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    bp.text = badge_text.upper()
    bp.font.name = "Outfit"
    bp.font.size = Pt(12)
    bp.font.bold = True
    bp.font.color.rgb = C_RED_RGB

    # Título Principal
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.9), Inches(6.8), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Outfit"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = C_WHITE_RGB

    # Subtítulo
    tb_sub = slide.shapes.add_textbox(Inches(1.0), Inches(5.1), Inches(6.8), Inches(1.2))
    tf_s = tb_sub.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
    ps = tf_s.paragraphs[0]
    ps.text = subtitle
    ps.font.name = "Calibri"
    ps.font.size = Pt(15)
    ps.font.color.rgb = C_MUTED_RGB

    # Tarjeta lateral derecha con imagen o preview
    if hero_img_path and hero_img_path.exists():
        add_image_card(slide, 8.2, 1.2, 4.3, 5.3, hero_img_path, caption="MMCall Academy Hardware Oficial")

    add_footer(slide, "MMCall Chile · Soporte y Desarrollo (S&D)")
    return slide


# =============================================================================
# 1. PRESENTACIÓN GENERAL Y VENTAS (01)
# =============================================================================

def build_pptx_01():
    out_file = DIR_01 / "Presentacion_Oficial_General_y_Ventas_ESL.pptx"
    logo_file = DIR_01 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_01 / "imagenes"

    prs = create_deck("01 General y Ventas ESL")
    blank = prs.slide_layouts[6]

    # Portada
    add_cover_slide(
        prs,
        title="Ecosistema de Etiquetas Electrónicas (ESL)",
        subtitle="Solución integral para digitalización de precios, sincronización omnicanal y agilidad operativa en sala de venta.",
        badge_text="Capacitación Oficial de Producto",
        logo_path=logo_file,
        hero_img_path=img_dir / "image1.jpg",
    )

    slides_data = [
        {
            "eyebrow": "Concepto Fundamental",
            "title": "Tecnología de Tinta Electrónica E-ink Bi-estable",
            "subtitle": "Principios físicos y ventajas de legibilidad extrema en punto de venta retail.",
            "card_title": "Características del Panel E-ink",
            "bullets": [
                ("Cero Consumo en Reposo:", "La tecnología bi-estable solo consume energía al alterar los pigmentos. Congela precios por años sin batería activa."),
                ("Ángulo de Visión de 180°:", "Visibilidad perfecta desde cualquier perspectiva, sin los molestos reflejos de las pantallas LCD."),
                ("Legibilidad Bajo Luz Solar:", "Comportamiento reflectivo que mejora su contraste bajo iluminación directa o tubos fluorescentes de tienda."),
                ("Paleta Tricolor:", "Pigmentos negro, blanco y rojo brillante para destacar ofertas, promociones flash y códigos."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "image5.png",
            "caption": "Góndola Retail con Pantallas E-ink",
        },
        {
            "eyebrow": "Eficiencia Operativa",
            "title": "Precios Dinámicos y Cumplimiento Normativo",
            "subtitle": "Sincronización instantánea entre góndola y caja evitando discrepancias y multas.",
            "card_title": "Agilidad y Precisión en Sala",
            "bullets": [
                ("Actualización en Segundos:", "Modifique miles de etiquetas simultáneamente desde el ERP central o eCommerce."),
                ("Cero Discrepancias POS:", "Elimina el 100% de los errores humanos por cambio tardío de flejes de papel."),
                ("Cumplimiento SERNAC:", "Previene infracciones a la Ley del Consumidor por diferencias entre precio exhibido y cobrado."),
                ("Campañas Automatizadas:", "Active promociones 'Happy Hour' o descuentos nocturnos de perecibles programados."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "image6.png",
            "caption": "Sincronización Dinámica de Precios",
        },
        {
            "eyebrow": "Agilidad Logística",
            "title": "Pick-by-Light LED Multicolor y Picking Express",
            "subtitle": "Optimización del tiempo de preparación de pedidos online en salas de venta y dark stores.",
            "card_title": "Operación Guiada por Luz",
            "bullets": [
                ("LED Frontal de 7 Colores:", "Luz ultrabrillante programable (rojo, verde, azul, cian, amarillo, lila, blanco)."),
                ("Rutas de Picking Paralelas:", "Diferentes operarios recolectan en el mismo pasillo guiados por su color asignado."),
                ("Reducción de Tiempos del 40%:", "Elimina la búsqueda visual de códigos EAN en estanterías saturadas."),
                ("Reposición Nocturna:", "El personal identifica de inmediato góndolas vacías o lotes con fecha próxima a vencer."),
            ],
            "accent": C_CYAN_RGB,
            "img": img_dir / "image14.png",
            "caption": "Dark Stores y Supermercados con Pick-by-Light",
        },
        {
            "eyebrow": "Catálogo de Hardware",
            "title": "Series de Etiquetas: MTag, STag, DS y RS",
            "subtitle": "Diversidad de formatos desde 2.1 hasta 7.5 pulgadas para toda categoría de retail.",
            "card_title": "Gama Oficial de Dispositivos",
            "bullets": [
                ("Serie MTag (2.9 y 4.2 pulg):", "El estándar del mercado para abarrotes, vestuario y farmacia con radio BLE 5.0 y NFC."),
                ("Serie STag (Ultra-Slim):", "Perfil ultra delgado de diseño prémium para joyerías, cosmética y óptica."),
                ("Serie DS (Gran Formato):", "Pantallas de 5.8 y 7.5 pulgadas ideales para cabeceras de góndola, electro y ferretería."),
                ("Serie RS (Grado Industrial):", "Carcasa hermética IP67 resistente al agua, polvo y frío extremo (-20 °C en congeladores)."),
            ],
            "accent": C_GREEN_RGB,
            "img": img_dir / "image12.png",
            "caption": "Catálogo Oficial Series MTag / STag / DS",
        },
        {
            "eyebrow": "Instalación Física",
            "title": "Accesorios y Sistemas Universales de Montaje",
            "subtitle": "Fijaciones modulares compatibles con cualquier perfil de estantería comercial.",
            "card_title": "Soportes de Alta Resistencia",
            "bullets": [
                ("Rieles para Góndola:", "Perfiles de policarbonato y aluminio que se insertan en la ranura estándar de estantes."),
                ("Clips para Canastos:", "Grapas de plástico rígido para cajas plásticas, bins agrícolas y exhibidores de fruta."),
                ("Soportes Magnéticos:", "Imanes de neodimio de fijación instantánea en racks metálicos y congeladores industriales."),
                ("Atriles T-Stand:", "Bases acrílicas inclinadas para mostradores, vitrinas de vidrio y mesones de atención."),
            ],
            "accent": C_AMBER_RGB,
            "img": img_dir / "image13.png",
            "caption": "Rieles, Clips y Fijaciones Universales",
        },
        {
            "eyebrow": "Infraestructura",
            "title": "Gateway Minew G1-E y Enlace Nube Seguro",
            "subtitle": "Concentrador inteligente PoE/Wi-Fi con cobertura masiva y cifrado MQTT SSL.",
            "card_title": "Especificaciones del Concentrador",
            "bullets": [
                ("Cobertura Extensa:", "Radio de 30 a 50 metros en línea de vista, gestionando hasta 5.000 etiquetas por unidad."),
                ("Conectividad Dual:", "Puerto Ethernet RJ45 con Power over Ethernet (PoE 802.3af) y Wi-Fi 2.4 GHz."),
                ("Túnel MQTT SSL:", "Cifrado de grado bancario TLS hacia la nube Minew Cloud o servidores locales."),
                ("Watchdog Inteligente:", "Monitoreo de latencia y auto-reinicio preventivo para 99.9% de uptime operativo."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "G1-E.png",
            "caption": "Gateway Minew G1-E de Alta Capacidad",
        },
        {
            "eyebrow": "Gestión Centralizada",
            "title": "Plataforma Minew Cloud y Diseñador Visual",
            "subtitle": "Administración multinivel omnicanal y creación de plantillas de alta fidelidad.",
            "card_title": "Capacidades de la Plataforma",
            "bullets": [
                ("Editor Drag & Drop:", "Diseñe plantillas con fuentes corporativas, logos, íconos de promoción y códigos de barras."),
                ("Monitoreo de Telemetría:", "Visualice nivel de batería, calidad de señal RSSI y estado de actualización de cada tag."),
                ("Integración ERP / API REST:", "Conectores nativos para sincronizar stock y precios desde SAP, Oracle, Odoo y POS."),
                ("Gestión Multinivel:", "Control centralizado de cadenas completas de supermercados con permisos por sucursal."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "image22.png",
            "caption": "Panel de Control Minew ESL Cloud",
        },
        {
            "eyebrow": "Impacto Económico",
            "title": "Retorno de Inversión (ROI) y Sostenibilidad",
            "subtitle": "Amortización comprobada en menos de doce meses y eliminación del gasto en papel.",
            "card_title": "Métricas Cuantificables de Ahorro",
            "bullets": [
                ("Reducción del 95% en Mano de Obra:", "El personal libera cientos de horas al mes para enfocarse en la atención al cliente."),
                ("Cero Insumos Desechables:", "Ahorro total en resmas de papel, cartuchos de tóner, etiquetas autoadhesivas e impresoras."),
                ("Periodo de Amortización (ROI):", "Recuperación de la inversión estimada entre 8 y 14 meses según el tamaño de la tienda."),
                ("Sello Verde Ecológico:", "Disminución directa de la huella de carbono y fomento de un retail sustentable."),
            ],
            "accent": C_GREEN_RGB,
            "img": img_dir / "image18.png",
            "caption": "Topología Integral del Ecosistema ESL",
        },
    ]

    total_s = len(slides_data) + 1
    for idx, d in enumerate(slides_data, 2):
        s = prs.slides.add_slide(blank)
        add_slide_background(s)
        add_header(s, logo_file, d["eyebrow"], d["title"], d["subtitle"], idx, total_s)
        add_bullets_card(s, 0.5, 1.75, 6.8, 5.0, d["card_title"], d["bullets"], accent_color=d["accent"])
        add_image_card(s, 7.6, 1.75, 5.2, 5.0, d["img"], caption=d["caption"])
        add_footer(s, "Ecosistema ESL · MMCall Academy")

    prs.save(str(out_file))
    print(f"[OK] Presentación 01 guardada en: {out_file} ({out_file.stat().st_size / 1024:.1f} KB)")


# =============================================================================
# 2. PRESENTACIÓN GATEWAY MINEW ESL (02)
# =============================================================================

def build_pptx_02():
    out_file = DIR_02 / "Presentacion_Oficial_Gateway_Minew_ESL.pptx"
    logo_file = DIR_02 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_02 / "imagenes"

    prs = create_deck("02 Gateway Minew ESL")
    blank = prs.slide_layouts[6]

    add_cover_slide(
        prs,
        title="Configuración y Diagnóstico del Gateway Minew G1-E",
        subtitle="Manual interactivo de puesta en marcha, conectividad Ethernet/Wi-Fi, túnel MQTT SSL y resolución de problemas.",
        badge_text="Manual Técnico de Infraestructura",
        logo_path=logo_file,
        hero_img_path=img_dir / "02_status_top.png",
    )

    slides_data = [
        {
            "eyebrow": "Acceso Inicial",
            "title": "Conexión por Modo Access Point (Wi-Fi de Servicio)",
            "subtitle": "Ingreso a la consola web local para la configuración inicial del gateway en terreno.",
            "card_title": "Paso a Paso de Conexión Local",
            "bullets": [
                ("SSID de Servicio:", "Conectarse a la red Wi-Fi emitida por el equipo (ej. 'Gateway_Minew_XXXX')."),
                ("Dirección IP por Defecto:", "Abrir navegador web e ingresar a la URL http://192.168.1.1."),
                ("Credenciales de Administrador:", "Usuario: admin | Contraseña oficial asignada por MMCall S&D."),
                ("Panel General (Status):", "Revisar versión de firmware, tiempo activo (uptime) y MAC serigrafiada."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "01_login.png",
            "caption": "Login Web del Gateway Minew G1-E",
        },
        {
            "eyebrow": "Red Cableada",
            "title": "Configuración de Red: Modo Ethernet LAN",
            "subtitle": "Configuración óptima para despliegues comerciales mediante cable de red con PoE.",
            "card_title": "Opciones de Enlace Cableado",
            "bullets": [
                ("Protocolo DHCP (Recomendado):", "Asigna automáticamente IP, máscara y DNS desde el router de la tienda."),
                ("IP Estática para Servidores:", "Configure IP fija, Gateway predeterminado y DNS si la red corporativa lo exige."),
                ("Alimentación PoE 802.3af:", "Permite energizar el dispositivo a través del mismo cable de datos hasta 100 metros."),
                ("LED de Enlace (LAN):", "Verifique que el LED verde frontal parpadee indicando tráfico activo."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "04_network_ethernet.png",
            "caption": "Parámetros de Interfaz Ethernet",
        },
        {
            "eyebrow": "Red Inalámbrica",
            "title": "Configuración de Red: Modo Wi-Fi Cliente",
            "subtitle": "Conexión del gateway a la red inalámbrica existente del local comercial.",
            "card_title": "Asociación Inalámbrica",
            "bullets": [
                ("Escaneo de Redes (Scan):", "Presionar 'Scan' para listar las redes inalámbricas 2.4 GHz disponibles."),
                ("Selección de SSID:", "Elegir la red del cliente (ej. 'MMCALL' o red corporativa) y presionar 'Select'."),
                ("Seguridad y Clave:", "Ingresar la contraseña WPA2/WPA3 Personal con cifrado AES."),
                ("Aplicar y Reiniciar:", "Pulsar 'Save & Apply' para que el equipo se asocie a la red Wi-Fi."),
            ],
            "accent": C_CYAN_RGB,
            "img": img_dir / "05_network_wireless_scan.png",
            "caption": "Escaneo de Redes Wi-Fi Disponibles",
        },
        {
            "eyebrow": "Ingeniería de Redes",
            "title": "Solución al Fallo DHCP Option 50 en Router Wi-Fi",
            "subtitle": "Diagnóstico y resolución de rechazos de IP en routers con control estricto de lease.",
            "card_title": "Causa Raíz y Corrección Técnica",
            "bullets": [
                ("Causa Raíz (DHCP Option 50):", "El cliente DHCP solicita su última IP previa. Si el router no la tiene en pool, rechaza la conexión con DHCPNAK."),
                ("Liberación de Lease (udhcpc):", "Reiniciar el cliente DHCP sin enviar Option 50 forzando un DISCOVER limpio."),
                ("Asignación Estática Alternativa:", "Fijar IP estática libre en el segmento del router para eludir la negociación DHCP."),
                ("Respaldo de Configuración:", "Guardar los parámetros en flash para que persistan ante cortes eléctricos."),
            ],
            "accent": C_AMBER_RGB,
            "img": img_dir / "07_network_wireless_form.png",
            "caption": "Formulario de Conexión Inalámbrica",
        },
        {
            "eyebrow": "Enlace Nube",
            "title": "Configuración del Servicio MQTT SSL con Minew Cloud",
            "subtitle": "Canal de telemetría seguro para recepción de plantillas y despacho de comandos.",
            "card_title": "Parámetros del Broker MQTT",
            "bullets": [
                ("Host del Servidor:", "Configurar mqtt.minewesl.com (o el host asignado a la instancia privada)."),
                ("Puerto Seguro SSL:", "Utilizar puerto 8883 con verificación de certificados TLS de grado bancario."),
                ("Client ID y Credenciales:", "Ingresar el identificador único del gateway y token de autenticación."),
                ("Tópicos de Publicación/Suscripción:", "Garantizan la entrega en tiempo real de actualizaciones hacia las etiquetas."),
            ],
            "accent": C_GREEN_RGB,
            "img": img_dir / "14_service_mqtt_top.png",
            "caption": "Parámetros MQTT SSL hacia Minew Cloud",
        },
        {
            "eyebrow": "Alta Disponibilidad",
            "title": "Estrategia de Auto-Reboot y Failover de Enlace",
            "subtitle": "Mecanismos de vigilancia automática para garantizar 99.9% de disponibilidad en sala.",
            "card_title": "Vigilancia de Conectividad (Watchdog)",
            "bullets": [
                ("Detección de Caída de Red:", "El gateway realiza pings periódicos a la puerta de enlace y al broker MQTT."),
                ("Reinicio Automático Programado:", "Reinicio preventivo cada 24 o 48 horas en horario no comercial (madrugada)."),
                ("Failover Ethernet a Wi-Fi:", "Conmutación automática hacia la red secundaria en caso de corte del cable físico."),
                ("Memoria de Tareas:", "Las órdenes pendientes se resguardan y reanudan inmediatamente tras restablecer enlace."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "21_other_auto_reboot_on.png",
            "caption": "Activación del Temporizador Auto-Reboot",
        },
        {
            "eyebrow": "Plataforma Central",
            "title": "Vinculación del Gateway en la Plataforma Minew Cloud",
            "subtitle": "Asignación del dispositivo a la tienda activa y verificación del estado Online.",
            "card_title": "Registro en la Nube",
            "bullets": [
                ("Ingreso a la Consola Web:", "Acceder a cloud.minewesl.com con credenciales autorizadas."),
                ("Módulo de Gateways (Base Stations):", "Presionar 'Add Gateway' e ingresar la dirección MAC serigrafiada."),
                ("Asignación de Tienda:", "Vincular el equipo a la sucursal correspondiente para enrutar sus etiquetas."),
                ("Verificación Online:", "El indicador cambiará a color verde confirmando comunicación bidireccional."),
            ],
            "accent": C_CYAN_RGB,
            "img": img_dir / "26_cloud_gateway_online.png",
            "caption": "Gateway Reportando Online en Minew Cloud",
        },
        {
            "eyebrow": "Mantenimiento y Soporte",
            "title": "Diagnóstico, Logs del Sistema y Actualización OTA",
            "subtitle": "Herramientas de inspección técnica avanzada y actualización remota de firmware.",
            "card_title": "Rutinas de Diagnóstico S&D",
            "bullets": [
                ("Descarga de Logs (klog / syslog):", "Permite revisar el handshake DHCP, conexión MQTT y eventos de radiofrecuencia."),
                ("Actualización de Firmware OTA:", "Cargue paquetes .bin de firmware para añadir mejoras y parches de seguridad."),
                ("Botón de Reset Físico:", "Presione por 10 segundos para restaurar parámetros de fábrica en caso de emergencia."),
                ("Soporte MMCall:", "Acceso a diagnóstico remoto y reemplazo de hardware con respaldo oficial."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "24_other_logs_reboot.png",
            "caption": "Descarga de Logs y Registro de Eventos",
        },
    ]

    total_s = len(slides_data) + 1
    for idx, d in enumerate(slides_data, 2):
        s = prs.slides.add_slide(blank)
        add_slide_background(s)
        add_header(s, logo_file, d["eyebrow"], d["title"], d["subtitle"], idx, total_s)
        add_bullets_card(s, 0.5, 1.75, 6.8, 5.0, d["card_title"], d["bullets"], accent_color=d["accent"])
        add_image_card(s, 7.6, 1.75, 5.2, 5.0, d["img"], caption=d["caption"])
        add_footer(s, "Configuración Gateway G1-E · MMCall Academy")

    prs.save(str(out_file))
    print(f"[OK] Presentación 02 guardada en: {out_file} ({out_file.stat().st_size / 1024:.1f} KB)")


# =============================================================================
# 3. PRESENTACIÓN ESL MODO STANDALONE (03)
# =============================================================================

def build_pptx_03():
    out_file = DIR_03 / "Presentacion_Oficial_ESL_Modo_Standalone.pptx"
    logo_file = DIR_03 / "imagenes" / "logo_mmcall_nobg.png"
    img_dir = DIR_03 / "imagenes"

    prs = create_deck("03 ESL Modo StandAlone")
    blank = prs.slide_layouts[6]

    add_cover_slide(
        prs,
        title="Etiquetas Electrónicas Modo StandAlone",
        subtitle="Operación móvil 1 a 1 mediante Bluetooth 5.0 y la app oficial Cloud Tag sin requerir infraestructura de Gateways.",
        badge_text="Manual Operativo Móvil",
        logo_path=logo_file,
        hero_img_path=img_dir / "01_portada_standalone.png",
    )

    slides_data = [
        {
            "eyebrow": "Arquitectura Móvil",
            "title": "Concepto StandAlone: Gestión 1 a 1 por BLE 5.0",
            "subtitle": "Operación autónoma y flexible para puntos de venta independientes o salas de exhibición.",
            "card_title": "Pilares de la Modalidad StandAlone",
            "bullets": [
                ("Cero Infraestructura Fija:", "No requiere gateways, cables Ethernet ni red Wi-Fi de alta cobertura."),
                ("El Smartphone como Emisor:", "La radio Bluetooth del celular inyecta la información directamente al tag."),
                ("Ideal para Pop-Ups y Salas:", "Excelente para ferreterías, tiendas de conveniencia y salas de reuniones."),
                ("App Oficial Cloud Tag:", "Disponible para iOS y Android con escaneo óptico por cámara integrado."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "02_qr_descarga_cloudtag.png",
            "caption": "Código QR para Descarga de Cloud Tag",
        },
        {
            "eyebrow": "Autenticación",
            "title": "Inicio de Sesión y Módulo 'Actualización Independiente'",
            "subtitle": "Ingreso seguro con credenciales corporativas y selección de la tienda activa.",
            "card_title": "Acceso al Entorno de Trabajo",
            "bullets": [
                ("Credenciales de Usuario:", "Ingresar usuario y clave proporcionados por el administrador MMCall."),
                ("Servidor Host:", "Confirmar que apunte al host oficial cloud.minewesl.com."),
                ("Detección de Sucursal:", "La app identifica automáticamente la tienda (ej. ORAFTI CHILE S.A.)."),
                ("Botón 'Actualización Independiente':", "Acceso directo a la interfaz de vinculación de etiquetas 1 a 1."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "03_login_cloudtag.png",
            "caption": "Inicio de Sesión en la App Móvil",
        },
        {
            "eyebrow": "Detección Óptica",
            "title": "Identificación de la Etiqueta: Escáner de Código de Barras",
            "subtitle": "El método más ágil y seguro en pasillos concurridos usando la cámara del smartphone.",
            "card_title": "Procedimiento de Escaneo Frontal",
            "bullets": [
                ("Activación de Cámara:", "Presionar el botón de escáner en la pantalla principal de Cloud Tag."),
                ("Apuntar al Código Frontal:", "Encuadrar el código de barras impreso en el marco de la etiqueta."),
                ("Filtrado Instantáneo de MAC:", "La app reconoce la dirección MAC en menos de un segundo."),
                ("Acceso a Vinculación:", "El sistema abre automáticamente la vista de selección de plantilla."),
            ],
            "accent": C_CYAN_RGB,
            "img": img_dir / "06_escaner_codigo_barra.png",
            "caption": "Escaneo Óptico del Código de Barras Frontal",
        },
        {
            "eyebrow": "Detección Inalámbrica",
            "title": "Identificación Alternativa: Exploración Bluetooth BLE",
            "subtitle": "Búsqueda inalámbrica por dirección MAC serigrafiada y medición de señal RSSI.",
            "card_title": "Selección por Baliza Bluetooth",
            "bullets": [
                ("Lista de Balizas BLE:", "Exploración de dispositivos en el radio de alcance inmediato."),
                ("Filtro por Dirección MAC:", "Cotejar los últimos 4 dígitos con la serigrafía posterior del tag."),
                ("Indicador de Distancia RSSI:", "Valores de -35 dBm a -45 dBm confirman cercanía a menos de 50 cm."),
                ("Toque para Conectar:", "Seleccionar el dispositivo para abrir el editor de contenido."),
            ],
            "accent": C_GREEN_RGB,
            "img": img_dir / "07_lista_bluetooth_mac.png",
            "caption": "Listado de Dispositivos BLE y Nivel RSSI",
        },
        {
            "eyebrow": "Diseño y Datos",
            "title": "Selección de Plantilla y Asociación de Producto",
            "subtitle": "Asignación del formato visual correspondiente al tamaño de la pantalla E-ink.",
            "card_title": "Emparejamiento de Información",
            "bullets": [
                ("Plantilla por Pulgadas:", "Elegir la plantilla que coincida exactamente (ej. 2.9 pulgadas tricolor)."),
                ("Catálogo de Productos:", "Buscar el artículo por nombre, SKU o código de barras del producto."),
                ("Modo Individual:", "Garantiza que la orden se envíe de manera exclusiva a la etiqueta seleccionada."),
                ("Vista Previa en Tiempo Real:", "Compruebe en la pantalla del celular cómo se verá el tag final."),
            ],
            "accent": C_AMBER_RGB,
            "img": img_dir / "08_unir_etiqueta_plantilla.png",
            "caption": "Asociación de Plantilla y Producto",
        },
        {
            "eyebrow": "Edición Flexible",
            "title": "Modificación de Valores en Vivo desde el Celular",
            "subtitle": "Edición inmediata de precios, nombres y enlaces sin necesidad de ingresar al ERP.",
            "card_title": "Personalización Directa en Sala",
            "bullets": [
                ("Botón 'Modificar':", "Permite sobreescribir cualquier campo de la plantilla seleccionada."),
                ("Teclado Móvil Integrado:", "Edite precios con descuento, texto de promoción o código SAP."),
                ("Generación Dinámica de QR:", "Ingrese una URL para que el código QR de la pantalla se genere en vivo."),
                ("Validación Rápida:", "Verifique que los textos no sobrepasen las dimensiones de la pantalla."),
            ],
            "accent": C_PURPLE_RGB,
            "img": img_dir / "10_modificar_valores_en_vivo.png",
            "caption": "Edición de Campos con Teclado Móvil",
        },
        {
            "eyebrow": "Transmisión RF",
            "title": "Inyección por Radiofrecuencia BLE y Refresco E-ink",
            "subtitle": "Transmisión de los paquetes de imagen y confirmación del ciclo de refresco físico.",
            "card_title": "Proceso de Actualización Física",
            "bullets": [
                ("Botón 'Seguro':", "Pulsar el botón azul para iniciar la transferencia Bluetooth de la imagen renderizada."),
                ("Distancia Recomendada:", "Mantenga el celular a menos de 50 cm del tag durante los 3 segundos de envío."),
                ("Parpadeo Bi-estable:", "La pantalla parpadeará brevemente mientras los pigmentos se reordenan."),
                ("Congelación Permanente:", "La nueva imagen queda fijada con consumo eléctrico cero."),
            ],
            "accent": C_RED_RGB,
            "img": img_dir / "13_transferencia_bluetooth_rf.png",
            "caption": "Transferencia de Imagen por Bluetooth 5.0",
        },
        {
            "eyebrow": "Auditoría Final",
            "title": "Verificación en Góndola y Buenas Prácticas",
            "subtitle": "Aseguramiento de calidad visual y resolución rápida de inconvenientes en terreno.",
            "card_title": "Checklist de Entrega",
            "bullets": [
                ("Inspección Visual:", "Confirmar nitidez de los dígitos, alineación del código de barras y QR."),
                ("Coincidencia de Plantilla:", "Nunca aplique plantillas de 4.2 pulg en etiquetas de 2.9 pulg para evitar recortes."),
                ("Batería del Smartphone:", "Asegúrese de contar con Bluetooth encendido y permisos de ubicación activos."),
                ("Listo para el Siguiente Tag:", "La app regresa al módulo lista para procesar la siguiente etiqueta en segundos."),
            ],
            "accent": C_CYAN_RGB,
            "img": img_dir / "15_etiqueta_actualizada_final.png",
            "caption": "Etiqueta Electrónica Actualizada en Góndola",
        },
    ]

    total_s = len(slides_data) + 1
    for idx, d in enumerate(slides_data, 2):
        s = prs.slides.add_slide(blank)
        add_slide_background(s)
        add_header(s, logo_file, d["eyebrow"], d["title"], d["subtitle"], idx, total_s)
        add_bullets_card(s, 0.5, 1.75, 6.8, 5.0, d["card_title"], d["bullets"], accent_color=d["accent"])
        add_image_card(s, 7.6, 1.75, 5.2, 5.0, d["img"], caption=d["caption"])
        add_footer(s, "Modo StandAlone ESL · MMCall Academy")

    prs.save(str(out_file))
    print(f"[OK] Presentación 03 guardada en: {out_file} ({out_file.stat().st_size / 1024:.1f} KB)")


def main():
    print("=== Generando las 3 Presentaciones Oficiales PowerPoint (MMCall Academy) ===")
    build_pptx_01()
    build_pptx_02()
    build_pptx_03()
    print("=== [COMPLETO] Las 3 presentaciones fueron generadas con éxito ===")


if __name__ == "__main__":
    main()
