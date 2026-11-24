# -*- coding: utf-8 -*-
"""
Generador de Video Full HD 1920x1080 — MMCall Academy (Soporte y Desarrollo S&D)
Título: Capacitación General y Ventas — Ecosistema de Etiquetas Electrónicas (ESL)
Voz: Victor Cayupil (generada localmente mediante VoiceStudio / OmniVoice)
"""

import argparse
import os
import subprocess
import sys
import time
import wave
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageFilter

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
IMG_DIR = BASE_DIR / "imagenes"
FRAME_DIR = SCRIPT_DIR / "video_frames"
AUDIO_CACHE = SCRIPT_DIR / "video_audio_cache"
SEG_DIR = SCRIPT_DIR / "video_segments"
FRAME_DIR.mkdir(exist_ok=True)
AUDIO_CACHE.mkdir(exist_ok=True)
SEG_DIR.mkdir(exist_ok=True)

FFMPEG = r"C:\Users\vcayu\AppData\Roaming\OmniVoice\media_tools\ffbin-df95abcb0ce6\win32\ffmpeg.exe"
REF_WAV_SRC = r"C:\Users\vcayu\AppData\Roaming\OmniVoice\voices\917e1e26-54629cc7.wav"
REF_WAV_TRIM = str(AUDIO_CACHE / "_ref_victor_8_5s.wav")
REF_TEXT_CLEAN = (
    "Hola, soy Victor Cayupil. En este vídeo revisaremos el funcionamiento "
    "y configuración del ecosistema inalámbrico de MMCall."
)

W, H = 1920, 1080

# Paleta MMCall Academy Dark Studio Standard
C_BG = (7, 10, 16)
C_CARD = (16, 22, 36)
C_CARD_HI = (22, 30, 48)
C_RED = (227, 6, 19)
C_PURPLE = (156, 137, 245)
C_CYAN = (34, 211, 238)
C_GREEN = (16, 185, 129)
C_AMBER = (245, 158, 11)
C_WHITE = (245, 247, 252)
C_MUTED = (156, 170, 194)
C_BORDER = (52, 65, 92)


def load_font(size: int, bold: bool = False, mono: bool = False):
    if mono:
        candidates = ["consolab.ttf" if bold else "consola.ttf", "courbd.ttf" if bold else "cour.ttf"]
    else:
        candidates = [
            "segoeuib.ttf" if bold else "segoeui.ttf",
            "calibrib.ttf" if bold else "calibri.ttf",
            "arialbd.ttf" if bold else "arial.ttf",
        ]
    for c in candidates:
        p = os.path.join(r"C:\Windows\Fonts", c)
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def make_studio_bg() -> Image.Image:
    y, x = np.mgrid[0:H, 0:W].astype(np.float32)
    r = np.full((H, W), 7.0, dtype=np.float32)
    g = np.full((H, W), 10.0, dtype=np.float32)
    b = np.full((H, W), 16.0, dtype=np.float32)

    d1 = np.sqrt(((x - W * 0.14) / (W * 0.42)) ** 2 + ((y - H * 0.18) / (H * 0.45)) ** 2)
    w1 = np.clip(1.0 - d1, 0.0, 1.0) ** 2
    r += w1 * 24.0
    g += w1 * 4.0
    b += w1 * 10.0

    d2 = np.sqrt(((x - W * 0.86) / (W * 0.46)) ** 2 + ((y - H * 0.82) / (H * 0.45)) ** 2)
    w2 = np.clip(1.0 - d2, 0.0, 1.0) ** 2
    r += w2 * 14.0
    g += w2 * 12.0
    b += w2 * 36.0

    grid = ((x % 80 == 0) | (y % 80 == 0)).astype(np.float32) * 5.0
    r += grid
    g += grid
    b += grid

    arr = np.stack([np.clip(r, 0, 255), np.clip(g, 0, 255), np.clip(b, 0, 255)], axis=2).astype(np.uint8)
    return Image.fromarray(arr, mode="RGB")


STUDIO_BG = None


def get_bg() -> Image.Image:
    global STUDIO_BG
    if STUDIO_BG is None:
        STUDIO_BG = make_studio_bg()
    return STUDIO_BG.copy()


def add_chrome(img: Image.Image, eyebrow: str, title: str, subtitle: str, step_num: int, total_steps: int):
    d = ImageDraw.Draw(img)
    d.rectangle([(0, 0), (W, 8)], fill=C_RED)

    logo_path = IMG_DIR / "logo_mmcall_nobg.png"
    if logo_path.exists():
        try:
            logo = Image.open(str(logo_path)).convert("RGBA")
            h_tgt = 46
            w_tgt = int(logo.width * (h_tgt / logo.height))
            logo = logo.resize((w_tgt, h_tgt), Image.Resampling.LANCZOS)
            img.paste(logo, (64, 42), logo)
            x_title = 64 + w_tgt + 24
        except Exception:
            x_title = 64
    else:
        x_title = 64

    f_eye = load_font(18, bold=True, mono=True)
    d.text((x_title, 36), eyebrow.upper(), fill=C_PURPLE, font=f_eye)

    f_h1 = load_font(34, bold=True)
    d.text((x_title, 64), title, fill=C_WHITE, font=f_h1)

    f_sub = load_font(21)
    d.text((64, 134), subtitle, fill=C_MUTED, font=f_sub)

    d.line([(64, 180), (W - 64, 180)], fill=C_BORDER, width=2)

    bw, bh = 190, 48
    bx = W - 64 - bw
    by = 44
    d.rounded_rectangle([(bx, by), (bx + bw, by + bh)], radius=10, fill=C_CARD_HI, outline=C_PURPLE, width=2)
    f_badge = load_font(19, bold=True, mono=True)
    badge_txt = f"MÓDULO {step_num} / {total_steps}"
    bbox = d.textbbox((0, 0), badge_txt, font=f_badge)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text((bx + (bw - tw) // 2, by + (bh - th) // 2 - 2), badge_txt, fill=C_PURPLE, font=f_badge)

    d.rectangle([(0, H - 56), (W, H)], fill=(4, 6, 11))
    d.line([(0, H - 56), (W, H - 56)], fill=C_BORDER, width=1)
    f_foot = load_font(16)
    foot_left = "MMCall Academy — Programa de Capacitación Técnica Oficial (S&D)"
    foot_right = "Solución Integral ESL · Ecosistema de Etiquetas Electrónicas"
    d.text((64, H - 38), foot_left, fill=C_MUTED, font=f_foot)
    bbox_r = d.textbbox((0, 0), foot_right, font=f_foot)
    d.text((W - 64 - (bbox_r[2] - bbox_r[0]), H - 38), foot_right, fill=C_MUTED, font=f_foot)


def draw_bullets(img: Image.Image, box, section_label: str, bullets, accent_color=C_RED):
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1 = box
    d.rounded_rectangle([(x0, y0), (x1, y1)], radius=16, fill=C_CARD, outline=C_BORDER, width=2)

    bar_len = 180
    d.line([(x0 + 28, y0 + 16), (x0 + 28 + bar_len, y0 + 16)], fill=accent_color, width=4)

    f_sec = load_font(21, bold=True)
    d.text((x0 + 28, y0 + 28), section_label.upper(), fill=accent_color, font=f_sec)

    cy = y0 + 76
    f_b = load_font(20, bold=True)
    f_t = load_font(19)

    for prefix, body in bullets:
        d.ellipse([(x0 + 32, cy + 6), (x0 + 44, cy + 18)], fill=accent_color)
        d.text((x0 + 56, cy), prefix, fill=C_WHITE, font=f_b)
        bbox_p = d.textbbox((x0 + 56, cy), prefix, font=f_b)
        tx = bbox_p[2] + 10
        rem_w = x1 - tx - 28

        words = body.split()
        lines = []
        cur_line = []
        for w_item in words:
            test = " ".join(cur_line + [w_item])
            bb = d.textbbox((0, 0), test, font=f_t)
            if bb[2] - bb[0] <= rem_w:
                cur_line.append(w_item)
            else:
                if cur_line:
                    lines.append(" ".join(cur_line))
                    cur_line = [w_item]
                else:
                    lines.append(w_item)
                    cur_line = []
        if cur_line:
            lines.append(" ".join(cur_line))

        for li, l_text in enumerate(lines):
            d.text((tx if li == 0 else x0 + 56, cy + li * 28), l_text, fill=C_MUTED, font=f_t)
        cy += max(len(lines) * 28, 28) + 16


def paste_contain(img: Image.Image, img_path: str, box, pad=12, caption: str = ""):
    x0, y0, x1, y1 = box
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([(x0, y0), (x1, y1)], radius=14, fill=C_CARD, outline=C_BORDER, width=2)

    avail_w = (x1 - x0) - pad * 2
    avail_h = (y1 - y0) - pad * 2 - (34 if caption else 0)

    p = Path(img_path)
    if not p.exists():
        f = load_font(18)
        d.text((x0 + 24, y0 + 24), f"Falta: {p.name}", fill=C_AMBER, font=f)
        return

    im = Image.open(str(p)).convert("RGB")
    im.thumbnail((avail_w, avail_h), Image.Resampling.LANCZOS)

    px = x0 + pad + (avail_w - im.width) // 2
    py = y0 + pad + (avail_h - im.height) // 2
    img.paste(im, (px, py))

    if caption:
        cap_y = y1 - 32
        d.rectangle([(x0 + 2, cap_y - 4), (x1 - 2, y1 - 2)], fill=(11, 15, 25))
        d.line([(x0 + 2, cap_y - 4), (x1 - 2, cap_y - 4)], fill=C_BORDER, width=1)
        f_cap = load_font(15, bold=True, mono=True)
        bb = d.textbbox((0, 0), caption, font=f_cap)
        tw = bb[2] - bb[0]
        d.text((x0 + ((x1 - x0) - tw) // 2, cap_y + 2), caption, fill=C_CYAN, font=f_cap)


# -----------------------------------------------------------------------------
# ESCENAS VISUALES GENERAL (1920x1080)
# -----------------------------------------------------------------------------

def render_s1():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Solución ESL",
        "1. Ecosistema de Etiquetas Electrónicas (ESL)",
        "Transformación digital en punto de venta con pantallas de tinta electrónica E-ink bi-estables.",
        1, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Fundamentos y Pantallas E-ink",
        [
            ("• Tinta Electrónica Bi-estable:", "Cero consumo de energía en reposo. El precio e imagen permanecen congelados indefinidamente."),
            ("• Legibilidad Extrema 180°:", "Ángulo de visión total sin reflejos, legible con luz solar o iluminación de tienda."),
            ("• Eliminación de Errores:", "Sincronización instantánea entre góndola y caja evitando sanciones de SERNAC."),
            ("• Autonomía de 5 a 10 Años:", "Alimentación por pilas botón reemplazables de máxima durabilidad."),
        ],
        accent_color=C_RED
    )
    paste_contain(c, str(IMG_DIR / "image1.jpg"), (1020, 258, 1420, 968), pad=12, caption="SISTEMA ESL EN GÓNDOLA")
    paste_contain(c, str(IMG_DIR / "image5.png"), (1440, 258, 1856, 968), pad=12, caption="VISIBILIDAD E-INK BI-ESTABLE")
    return c


def render_s2():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Eficiencia y Omnicanalidad",
        "2. Precios Dinámicos y Pick-by-Light LED",
        "Agilidad operativa en sala, picking para e-commerce y promociones flash sincronizadas.",
        2, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Capacidades Operativas Avanzadas",
        [
            ("• Actualización Masiva:", "Cambie miles de precios en segundos desde el ERP o eCommerce central."),
            ("• Pick-by-Light LED:", "Luz indicadora en 7 colores para picking express en dark stores y reposición nocturna."),
            ("• Interacción NFC y QR:", "Permite al cliente escanear con su smartphone para ver stock, recetas o pagar."),
            ("• Cero Desperdicio de Papel:", "Elimina costos de impresión, tintas térmicas y horas de reemplazo manual."),
        ],
        accent_color=C_PURPLE
    )
    paste_contain(c, str(IMG_DIR / "image6.png"), (1020, 258, 1420, 968), pad=12, caption="PICK-BY-LIGHT & PROMOCIONES")
    paste_contain(c, str(IMG_DIR / "image14.png"), (1440, 258, 1856, 968), pad=12, caption="SUPERMERCADOS Y RETAIL")
    return c


def render_s3():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Catálogo de Dispositivos",
        "3. Series de Hardware y Accesorios de Montaje",
        "Formatos versátiles desde 2.1 hasta 7.5 pulgadas diseñados para cada categoría de producto.",
        3, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Gama de Etiquetas y Accesorios",
        [
            ("• Serie MTag (2.9 y 4.2 pulg):", "Estándar de retail con pantalla tricolor (negro, blanco, rojo), BLE 5.0 y NFC."),
            ("• Serie STag (Ultra Slim):", "Perfil ultra delgado para vitrinas de joyería, cosmética y farmacias."),
            ("• Serie RS (Grado Industrial):", "Carcasa IP67 hermética para congeladores de frío extremo y áreas húmedas."),
            ("• Fijaciones Universales:", "Rieles de góndola, clips para canastos, soportes magnéticos y atriles T-stand."),
        ],
        accent_color=C_CYAN
    )
    paste_contain(c, str(IMG_DIR / "image12.png"), (1020, 258, 1420, 968), pad=12, caption="CATÁLOGO SERIES MTAG / STAG")
    paste_contain(c, str(IMG_DIR / "image13.png"), (1440, 258, 1856, 968), pad=12, caption="SISTEMAS DE FIJACIÓN Y RIELES")
    return c


def render_s4():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Infraestructura Inalámbrica",
        "4. Gateway Minew G1-E y Enlace Nube Seguro",
        "Concentrador inteligente PoE y Wi-Fi con túnel MQTT cifrado de alta capacidad.",
        4, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Topología y Conectividad",
        [
            ("• Cobertura RF BLE 5.0:", "Radio de 30 a 50 metros en línea de vista, administrando hasta 5.000 tags."),
            ("• Alimentación Dual:", "Soporte Power-over-Ethernet (PoE 802.3af) o alimentación local DC 5V."),
            ("• Seguridad MQTT SSL:", "Cifrado TLS de grado bancario hacia Minew Cloud o servidores locales on-premise."),
            ("• Watchdog Automático:", "Failover de enlace y reinicio preventivo para 99.9% de disponibilidad en sala."),
        ],
        accent_color=C_GREEN
    )
    paste_contain(c, str(IMG_DIR / "G1-E.png"), (1020, 258, 1420, 968), pad=12, caption="GATEWAY MINEW G1-E")
    paste_contain(c, str(IMG_DIR / "image18.png"), (1440, 258, 1856, 968), pad=12, caption="TOPOLOGÍA DE RED Y NUBE")
    return c


def render_s5():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Software y Gestión",
        "5. Plataforma Cloud Minew y App Cloud Tag",
        "Gestión omnicanal centralizada, editor visual de plantillas y app móvil para terreno.",
        5, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Herramientas de Administración",
        [
            ("• Portal Web Central:", "Supervisión en vivo de señal RF, estado de batería y tareas de actualización pendientes."),
            ("• Editor Visual de Plantillas:", "Diseño drag-and-drop con fuentes, logotipos corporativos y códigos de barra dinámicos."),
            ("• Conexión ERP / API REST:", "Sincronización automatizada con SAP, Oracle, Odoo y puntos de venta POS."),
            ("• App Móvil Cloud Tag:", "Asociación rápida con cámara del smartphone y modo StandAlone autónomo."),
        ],
        accent_color=C_AMBER
    )
    paste_contain(c, str(IMG_DIR / "image22.png"), (1020, 258, 1420, 968), pad=12, caption="PLATAFORMA WEB MINEW CLOUD")
    paste_contain(c, str(IMG_DIR / "image23.png"), (1440, 258, 1856, 968), pad=12, caption="APP MÓVIL CLOUD TAG")
    return c


def render_s6():
    c = get_bg()
    add_chrome(
        c, "MMCall Academy · Impacto Financiero",
        "6. Retorno de Inversión (ROI) y Respaldo Oficial",
        "Recuperación económica en menos de doce meses y soporte técnico especializado en Chile.",
        6, 6
    )
    draw_bullets(
        c, (64, 258, 980, 968),
        "Beneficios Cuantificables del Negocio",
        [
            ("• Retorno Rápido (ROI):", "Amortización de 8 a 14 meses por reducción drástica de insumos y mano de obra."),
            ("• Cumplimiento Normativo:", "Cero multas por diferencias de precio en caja y total transparencia al consumidor."),
            ("• Imagen de Innovación:", "Moderniza la sala de venta transmitiendo tecnología de vanguardia y sustentabilidad."),
            ("• Respaldo MMCall Chile:", "Garantía oficial, repuestos locales, soporte técnico y capacitación certificada."),
        ],
        accent_color=C_RED
    )
    paste_contain(c, str(IMG_DIR / "image14.png"), (1020, 258, 1420, 968), pad=12, caption="RETAIL MODERNO MMCall")
    paste_contain(c, str(IMG_DIR / "image20.jpeg"), (1440, 258, 1856, 968), pad=12, caption="NITIDEZ E-INK DE ALTO IMPACTO")
    return c


SCENES = [
    {
        "name": "s1_intro_ecosistema",
        "render": render_s1,
        "chunks": [
            "Bienvenidos al ecosistema de etiquetas electrónicas de MMCall Academy. Nuestras pantallas de tinta electrónica ofrecen legibilidad perfecta, cero consumo en reposo y sincronización automática entre la góndola y el sistema de caja.",
        ],
    },
    {
        "name": "s2_ventajas_picking",
        "render": render_s2,
        "chunks": [
            "La tecnología ESL permite actualizar miles de precios en segundos desde su ERP corporativo. Gracias al LED multicolor Pick-by-Light, el personal optimiza la preparación de pedidos online y la reposición nocturna con total precisión.",
        ],
    },
    {
        "name": "s3_catalogo_hardware",
        "render": render_s3,
        "chunks": [
            "Disponemos de un catálogo integral adaptado a cada necesidad: la serie MTag para retail general, la serie STag de diseño ultra delgado, y la serie industrial RS para zonas húmedas y cámaras de congelación.",
        ],
    },
    {
        "name": "s4_red_gateway_g1e",
        "render": render_s4,
        "chunks": [
            "El Gateway Minew G1-E es el corazón de la infraestructura inalámbrica. Con alimentación PoE y conexión MQTT cifrada mediante SSL, gestiona miles de etiquetas con un alcance de hasta cincuenta metros en línea de vista.",
        ],
    },
    {
        "name": "s5_software_cloud_app",
        "render": render_s5,
        "chunks": [
            "Nuestra plataforma Cloud centraliza la gestión de plantillas visuales, monitorea el nivel de batería de cada etiqueta e integra directamente su sistema ERP. Además, la aplicación móvil Cloud Tag facilita la auditoría y vinculación rápida en tienda.",
        ],
    },
    {
        "name": "s6_roi_conclusion",
        "render": render_s6,
        "chunks": [
            "Implementar etiquetas electrónicas MMCall garantiza un retorno de inversión en menos de un año, optimiza las horas hombre y moderniza la imagen de su tienda. Los invitamos a continuar con los módulos especializados de Gateway y modo StandAlone.",
        ],
    },
]


def trim_reference_audio():
    if Path(REF_WAV_TRIM).exists():
        return REF_WAV_TRIM
    print(f"Recortando audio de referencia de Victor Cayupil a 8.5s (calibración fonética exacta): {REF_WAV_SRC} -> {REF_WAV_TRIM}")
    cmd = [
        FFMPEG,
        "-y",
        "-i",
        str(REF_WAV_SRC),
        "-ss",
        "0.0",
        "-t",
        "8.5",
        "-ar",
        "24000",
        "-ac",
        "1",
        REF_WAV_TRIM,
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return REF_WAV_TRIM


def get_audio_duration(wav_path: str) -> float:
    with wave.open(wav_path, "rb") as wf:
        return wf.getnframes() / float(wf.getframerate())


def render_segment_mp4(frame_png: str, audio_wav: str, dur: float, out_mp4: str):
    dur_str = f"{dur:.2f}"
    cmd = [
        FFMPEG,
        "-y",
        "-loop",
        "1",
        "-i",
        frame_png,
        "-i",
        audio_wav,
        "-af",
        "apad=pad_dur=0.4",
        "-c:v",
        "libx264",
        "-tune",
        "stillimage",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-pix_fmt",
        "yuv420p",
        "-t",
        dur_str,
        out_mp4,
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def concat_segments(seg_list, final_mp4):
    list_txt = SEG_DIR / "concat_list.txt"
    with open(list_txt, "w", encoding="utf-8") as f:
        for s in seg_list:
            s_norm = s.replace("\\", "/")
            f.write(f"file '{s_norm}'\n")

    cmd = [
        FFMPEG,
        "-y",
        "-f",
        "concat",
        "-safe",
        "0",
        "-i",
        str(list_txt),
        "-c",
        "copy",
        str(final_mp4),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--frames-only", action="store_true", help="Solo generar los frames PNG")
    parser.add_argument("--skip-tts", action="store_true", help="Usar audios en cache")
    args = parser.parse_args()

    os.environ["PATH"] = os.path.dirname(FFMPEG) + os.pathsep + os.environ.get("PATH", "")

    print("=== 1. Generando Frames Gráficos Full HD (1920x1080) ===", flush=True)
    for sc in SCENES:
        frame_path = FRAME_DIR / f"{sc['name']}.png"
        img = sc["render"]()
        img.save(str(frame_path), quality=95)
        print(f"  [OK] Frame guardado: {frame_path.name}", flush=True)

    if args.frames_only:
        print("[FIN] Modo --frames-only completado con éxito.")
        return

    print("=== 2. Cargando Motor Local VoiceStudio (OmniVoice) con voz 'Victor Cayupil' ===", flush=True)
    import torch
    from omnivoice.models.omnivoice import OmniVoice
    from omnivoice.utils.audio import save_audio
    from omnivoice.utils.dtype import tts_dtype_name

    ref_wav = trim_reference_audio()

    all_audios_cached = True
    for sc in SCENES:
        for ci in range(1, len(sc["chunks"]) + 1):
            cw = AUDIO_CACHE / f"{sc['name']}_c{ci}.wav"
            if not cw.exists():
                all_audios_cached = False
                break

    if not all_audios_cached and not args.skip_tts:
        device = "cuda:0" if torch.cuda.is_available() else "cpu"
        print(f"Dispositivo de inferencia TTS: {device}", flush=True)
        t_load = time.time()
        model = OmniVoice.from_pretrained(
            "k2-fsa/OmniVoice",
            device_map=device,
            dtype=getattr(torch, tts_dtype_name(device)),
        )
        print(f"  [OK] Modelo OmniVoice cargado en {time.time() - t_load:.1f}s. Pre-codificando voz...", flush=True)
        clone_prompt = model.create_voice_clone_prompt(
            ref_audio=ref_wav,
            ref_text=REF_TEXT_CLEAN,
        )

        for sc in SCENES:
            for ci, text_chunk in enumerate(sc["chunks"], 1):
                cw = AUDIO_CACHE / f"{sc['name']}_c{ci}.wav"
                if cw.exists() and cw.stat().st_size > 1000:
                    print(f"  [CACHE] {cw.name} ya existe.", flush=True)
                    continue
                print(f"  [TTS] Sintetizando {sc['name']} trozo {ci}/{len(sc['chunks'])}: \"{text_chunk[:45]}...\"", flush=True)
                t0 = time.time()
                audios = model.generate(
                    text=text_chunk,
                    voice_clone_prompt=clone_prompt,
                    language="es",
                    speed=1.03,
                )
                save_audio(str(cw), audios[0], model.sampling_rate, bits_per_sample=16)
                print(f"    -> Guardado {cw.name} en {time.time() - t0:.1f}s", flush=True)
    else:
        print("  [INFO] Audios en cache disponibles. Saltando inferencia TTS.", flush=True)

    print("=== 3. Ensamblando Segmentos MP4 con FFmpeg ===", flush=True)
    seg_mp4s = []
    for sc in SCENES:
        sc_name = sc["name"]
        frame_png = str(FRAME_DIR / f"{sc_name}.png")
        chunk_wavs = [str(AUDIO_CACHE / f"{sc_name}_c{ci}.wav") for ci in range(1, len(sc["chunks"]) + 1)]
        scene_wav = str(AUDIO_CACHE / f"{sc_name}_FULL.wav")

        list_f = AUDIO_CACHE / f"{sc_name}_list.txt"
        with open(list_f, "w", encoding="utf-8") as lf:
            for cw in chunk_wavs:
                cw_norm = cw.replace("\\", "/")
                lf.write(f"file '{cw_norm}'\n")

        subprocess.run(
            [FFMPEG, "-y", "-f", "concat", "-safe", "0", "-i", str(list_f), "-c", "copy", scene_wav],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        dur = get_audio_duration(scene_wav) + 0.6
        out_mp4 = str(SEG_DIR / f"{sc_name}.mp4")
        print(f"  [SEG] Renderizando {sc_name}.mp4 ({dur:.1f}s)...", flush=True)
        render_segment_mp4(frame_png, scene_wav, dur, out_mp4)
        seg_mp4s.append(out_mp4)

    final_video = BASE_DIR / "Video_Capacitacion_General_ESL.mp4"
    print(f"=== 4. Concatenando Video Final Full HD: {final_video.name} ===", flush=True)
    concat_segments(seg_mp4s, str(final_video))
    sz_mb = final_video.stat().st_size / (1024 * 1024)
    print(f"[OK COMPLETO] Video Oficial generado en: {final_video} ({sz_mb:.2f} MB)", flush=True)


if __name__ == "__main__":
    main()
