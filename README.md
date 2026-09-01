# 🎓 MMCall Academy - Document Generator Skill

> Skill oficial e instalable de automatización y estandarización visual para **Claude Code**, **Google Antigravity**, **Cursor** y asistentes de IA. Genera manuales técnicos profesionales (Word `.docx`, PDF) y presentaciones interactivas (HTML5) con la identidad de **MMCall Academy & Soporte y Desarrollo (S&D)**.

---

## ⚡ Instalación Rápida para Claude / Gemini

### 🤖 Método 1: Pedírselo directamente a Claude Code
Si estás usando **Claude Code** en la terminal, simplemente indícale:
```text
Claude, clona e instala la skill de MMCall Academy desde este repositorio:
git clone https://github.com/TU_USUARIO/mmcall-document-generator.git ~/.claude/skills/mmcall-document-generator
```

### 💻 Método 2: Instalación por Script (PowerShell / Bash)
Clona el repositorio y ejecuta el instalador:

**En Windows (PowerShell):**
```powershell
git clone https://github.com/TU_USUARIO/mmcall-document-generator.git
cd mmcall-document-generator
.\install.ps1
```

**En Linux / Mac / WSL:**
```bash
git clone https://github.com/TU_USUARIO/mmcall-document-generator.git
cd mmcall-document-generator
bash install.sh
```

---

## 🚀 ¿Qué hace esta Skill automáticamente?

Cuando la Skill está instalada en tu IA, cada vez que le pidas crear un manual o presentación:

```text
"Genera el manual técnico y la presentación interactiva para el equipo [Nombre del Dispositivo]"
```

La IA aplicará automáticamente:

### 1. 📄 En Manuales Word (`.docx`) y PDF:
* **Portada Institucional:** Logotipo oficial centrado, títulos en Azul Corporativo (`#1F3A8A`), subtítulos en Lila (`#9C89F5`), metadatos de S&D y el *Sello de Calidad MMCall Academy* en verde (`#27AE60`).
* **Encabezados Académicos:** Tabla de cabecera superior en páginas 2+ (`MMCALL ACADEMY | Manual Técnico - [Equipo]`) con línea sutil de división.
* **Pie de Página Oficial:** Leyenda `S&D Mmcall Paging Services LTDA.` en gris 8.5pt.
* **Geometría y Tipografía:** Arial 10.5pt, interlineado 1.15, espacio de 6pt entre párrafos y márgenes de 1 pulgada (2.54 cm).
* **Tablas Cebra:** Celdas con padding y fondo alterno `#F2F4F7`, encabezado azul con texto blanco.
* **Estructura Visual de Imágenes:** Capturas de pantalla del software insertadas al inicio de cada sección explicativa para mostrar los parámetros exactos a escribir antes del procedimiento.
* **Compilación a PDF:** Conversión automática de alta fidelidad vía `docx2pdf`.

### 2. 🖥️ En Presentaciones Interactivas (HTML5):
* **Glassmorphic UI:** Fondo oscuro (`#0f172a`) con tarjetas de vidrio translúcidas (`backdrop-filter: blur(16px)`).
* **Escalado Adaptativo 16:9:** Optimizado para proyectores y pantallas de 32 pulgadas en terreno.
* **Paneles con Scroll:** Scrollbars suaves (`overflow-y: auto; max-height: 480px;`) para no cortar textos explicativos largos.
* **Zoom Lightbox Modal:** Ampliación accesible de capturas mediante click y atajos de teclado (`Tab`, `Enter`, `Espacio`, `Escape`).
* **Cuestionario Evaluativo:** Módulo de evaluación de 5 preguntas técnicas con retroalimentación en tiempo real y resumen de puntaje con el lema de MMCall Academy.

---

## 📁 Estructura del Repositorio

```text
mmcall-document-generator/
├── SKILL.md                          # Reglas y directrices completas que lee la IA
├── CLAUDE.md                         # Directiva de auto-activación para Claude Code
├── README.md                         # Documentación y guía de instalación
├── install.ps1                       # Instalador automático para Windows
├── install.sh                        # Instalador automático para Linux/Mac
├── .gitignore                        # Archivos excluidos del control de versiones
└── resources/
    ├── template_manual.py            # Generador de Word y PDF con python-docx
    └── template_presentation.html    # Boilerplate interactivo HTML/CSS/JS
```

---

## 🎨 Paleta de Colores Oficial

| Elemento | Código HEX | Rol |
| :--- | :--- | :--- |
| **Azul MMCall** | `#1F3A8A` / `#1E3A8A` | Títulos principales, cabeceras de tablas |
| **Lila Academy** | `#9C89F5` | Subtítulos, destaques interactivos |
| **Rojo MMCall** | `#E53935` | Alertas, botones y acciones |
| **Verde Calidad** | `#27AE60` | Sello de Certificación de Calidad |
| **Fondo Oscuro UI**| `#0F172A` | Background para presentaciones interactivas |
| **Gris Neutro** | `#7F8C8D` / `#94A3B8` | Pies de figura, textos de cabecera y pie |

---

## 🏢 Créditos
Desarrollado y mantenido por **Soporte y Desarrollo (S&D)**  
**MMCall Paging Services LTDA. - Chile**