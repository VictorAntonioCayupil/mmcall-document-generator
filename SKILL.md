---
name: mmcall-document-generator
description: >-
  Generates and formats documents (HTML presentations, Word manuals, PDF guides) following the official MMCall Academy style guidelines.
  Use this skill when the user requests a new manual, PDF, Word document, or interactive presentation for MMCall Academy, or wants to format an existing document to fit the standard styling.
---

# MMCall Academy Document Generator Skill

Use this skill to create, structure, and format educational and technical resources for MMCall Academy (including interactive HTML slideshows, Word `.docx` user manuals, and PDF compilations) following a strict, unified visual and structural standard.

--------------------------------------------------------------------------------

## 1. Word / PDF Document Design Guidelines

All printable manuals and guides must adhere to the following formatting specifications. These were verified field-by-field against `resources/template_manual.py`'s output *and* two independently published documents in `REPOSITORIO MMCALL ACADEMY\SyD\` from unrelated product families (`SENSORES\Sensor portillon\S_AlertaPermitetral.docx` and `T02\MANUAL T02 PAGER.docx`) — both use identical fonts, colors and structure, so this is the real company-wide standard, not one document's choice. An earlier version of this guide said `Arial` and a navy-blue/lila-purple/green palette; that was never correct — don't revert to it.

### Page Layout and Geometry
*   **Margins:** Exact 1.0 inch (2.54 cm) on all four sides.
*   **Page Setup:** First page header/footer must be different (disabled on the cover page, enabled on pages 2+).

### Fonts and Colors (exact values, verified)
*   **Primary Font:** `Calibri` for everything except the portada eyebrow line (below). Body text color `#333333`, no explicit size override needed beyond `Normal` style at 11pt.
*   **Color tokens:**
    ```
    Red        #E30613   main title, page-2 title repeat, table header fill
    Lila       #9C89F5   "ACADEMY" half of the brand text, portada eyebrow
    Dark gray  #404040   "MMCALL" half of the brand text, all section headings
    Med gray   #7F8C8D   header/footer running text (pages 2+)
    Light gray #808080   portada subtitle, "Sello de Calidad" line
    Body text  #333333
    ```
    There is no green anywhere and no navy blue anywhere — those were an invented palette from an earlier version of this guide.

### Headers and Footers (Pages 2+)
*   **Header:** A 1x2 grid table.
    *   Left cell: Empty.
    *   Right cell: Aligned right, three runs in sequence — `MMCALL ` (10pt, bold, `#404040`) + `ACADEMY` (10pt, bold, `#9C89F5`) + `  |  [Nombre del Documento]` (9.5pt, `#7F8C8D`). Not one flat-colored string.
*   **Footer:** Aligned left, 9pt Calibri, `#7F8C8D`: `S&D Mmcall Paging Services LTDA.`

### Document Cover Page Structure
1.  **Header Title:** `PROGRAMA DE CAPACITACIÓN ACADÉMICA` — **`Consolas`** (monospace, the one deliberate font departure), 9pt, bold, Lila `#9C89F5`, centered.
2.  **Brand line:** `MMCALL ` (Calibri 24pt bold `#404040`) + `ACADEMY` (Calibri 24pt bold `#9C89F5`) — two runs, two colors, centered.
3.  **Logo:** Center-aligned image `logo_mmcall_nobg.png` (width 2.5 inches) — the real file, never a drawn/placeholder badge.
4.  **Main Title:** the document's own name (e.g. `Manual Técnico - [Dispositivo]`), Calibri 26pt bold, **Red `#E30613`**, centered.
5.  **Subtitle:** one line of context, Calibri 13pt, Light gray `#808080`, centered.
6.  **Quality Certification:** `Sello de Calidad MMCall Academy`, Calibri 10pt, **not bold**, italic, Light gray `#808080` (no green anywhere).
7.  **Metadata Footer:** `Soporte y Desarrollo (S&D) - MMCall Paging Services LTDA. - Chile`, Calibri 9.5pt, Med gray `#7F8C8D`, centered.
8.  **Page break, then a title repeat:** the same main title text again at the top of page 2 (Calibri 18pt bold, Red `#E30613`, centered) before the first `add_heading_1` — present in both reference documents, easy to miss if copying only the portada block.

### Body Styling
*   **Headings:**
    *   **Heading 1:** 14pt, bold, Dark gray `#404040` (not navy, not lila), space before 18pt, space after 6pt.
    *   **Heading 2:** 12pt, bold, Dark gray `#404040` — same color as Heading 1, only the size differs.
*   **Bullet Points:** Use standard list bullets with a bold prefix run for clarity (e.g., `• **1. Paso:** Explicación.`).
*   **Zebra Tables:** Centered tables, white text on **`#E30613`** (red) background for the header row. Alternate data rows `#F2F2F2` / `#FFFFFF` (not `#F2F4F7`). Use `add_zebra_table()` from the template rather than rebuilding this per-document — it already implements the padding and shading.
*   **Images:** Centered, width 4.5 inches. Include a caption below in Calibri 9pt italic, Med gray `#7F8C8D`: `Figura: [Descripción]`.
*   **Image Placement:** Always place configuration screenshots of the software at the **beginning** of their respective section or step so the user knows what values to enter beforehand.

--------------------------------------------------------------------------------

## 2. Interactive Presentation Design Guidelines

`resources/template_presentation.html` is not an interpretation of "glassmorphism" — its tokens, class names and structure were extracted directly from a real published presentation (`REPOSITORIO MMCALL ACADEMY\SyD\SENSORES\Sensor portillon\presentacion_autocontenida.html`). Start every new presentation from that file and keep its class names; inventing new ones (or a different palette) is how a generated deck ends up not resembling the rest of the company's material — this happened once already and had to be redone.

### Fonts and Colors (exact tokens, don't approximate)
*   **Fonts:** `Inter` (body text) + `Outfit` (headings, brand text, buttons) + `Share Tech Mono` (badges, code/command snippets, the slide counter) — all three loaded via one Google Fonts link, see the template's `<head>`.
*   **Color tokens** (CSS custom properties, copy verbatim):
    ```css
    --bg-primary: #07090e;      --bg-secondary: #0f131f;
    --text-main: #f3f4f6;       --text-muted: #9ca3af;
    --mmcall-red: #e30613;      --mmcall-red-hover: #b8050f;
    --accent-blue: #2979ff;     --accent-green: #00e676;
    --accent-orange: #f59e0b;   --accent-gold: #ffb300;
    --accent-purple: #d500f9;
    --glass-bg: rgba(15, 20, 32, 0.75);
    --glass-border: rgba(255, 255, 255, 0.06);
    ```
    `#e30613` is the real MMCall red — not `#e53935`, which is close but wrong. `--accent-gold` (`#ffb300`) is the card-title / active-dot / badge color, not the lila purple — lila (`--accent-purple`, `#d500f9`) is reserved for the "ACADEMY" half of the brand text.

### Layout and Scaling
*   **Fixed 16:9 stage:** a `1280×720` container (`.slides-scale-container`), scaled to fit the viewport with a JS `transform: scale()` — not a fluid/responsive layout. See the template's `adjustScale()`.
*   **Ambient glow:** two blurred radial-gradient circles behind the content (`.glow-1` red top-right, `.glow-2` blue bottom-left) — subtle, not decoration to overdo.
*   **`h2` signature style:** every section heading gets a solid `4px` red left border (`border-left: 4px solid var(--mmcall-red)`) — this is the single most recognizable visual signature across real documents; don't drop it.
*   **Cards:** `.glass-card` (`background: var(--glass-bg)`, `border-radius: 12px`, `backdrop-filter: blur(10px)`) with an emoji-prefixed `<h3>` in gold and a muted `<p>` — this is the primary content unit, used far more than plain paragraphs.
*   **Screenshots:** `.shot-frame` (not a plain `<img>`) — bordered, black background, `object-fit: contain`, with a small "🔍 ampliar" hint and `cursor: zoom-in`.
*   **Dense text panels** (command lists, long bullet lists) get `overflow-y: auto; max-height: 480px;` so they scroll instead of truncating.

### Header, Logo and Branding
*   **Real logo, not a placeholder:** the header and the cover slide's `.cover-seal` both embed the actual `logo_mmcall_nobg.png` (Base64 data URI) as an `<img class="logo-img">` — never a CSS-drawn badge or initials standing in for the logo.
*   **Header layout:** logo image → `.brand-divider` → `.brand-text` (`MMCALL <span>ACADEMY</span>`, the span in `--accent-purple`) → `.brand-divider` → `.header-title` (the specific document's name/subtitle). Header has a `2px solid var(--mmcall-red)` bottom border.
*   **Cover slide:** `.badge-gold` eyebrow pill → `.main-title` (gradient-clipped white-to-gray text) → `.main-subtitle` → `.cover-seal` (small logo + "Sello de Calidad MMCall Academy").

### Navigation and Interactions
*   **Dots, not a session badge:** `.dots-navigation` in the header, one `.dot-nav` per slide, the active one gold and glowing (`box-shadow: 0 0 8px var(--accent-gold)`).
*   **Footer controls:** `.nav-btn` Anterior/Siguiente buttons plus a `Share Tech Mono` slide counter (`01 / 09` style), not just arrows.
*   **Keyboard Controls:** Map arrow keys `ArrowLeft` / `ArrowRight` to change slides.
*   **Image Lightbox:** Expandable screenshot overlay when clicking a `.shot-frame`. Keyboard focusable (`tabindex="0"`) and triggerable via `Enter` / `Space`, closeable via `Escape` or clicking outside the image.
*   **Interactive Quiz:** `.quiz-opt-btn` options that lock and color on click (`.correct` green, `.incorrect` red, revealing the right answer), a final score screen, and a retry button. 5 questions is the usual count but scale it to how much the document actually covers — don't pad or cut content to force exactly 5.

--------------------------------------------------------------------------------

## 3. Templates and Automation Scripts

Use the template scripts stored in the skill resources to build new deliverables quickly:
*   [template_manual.py](./resources/template_manual.py) - Python script using `docx` and `docx2pdf` to compile manuals.
*   [template_presentation.html](./resources/template_presentation.html) - Base boilerplate for interactive slideshows.

### Reference examples (ground truth)

`C:\Users\vcayu\Desktop\REPOSITORIO MMCALL ACADEMY\SyD\` holds the manuals and presentations already published for other MMCall products (ESL, RELOJES NS818, SENSORES — including the portillón sensor's own technical/command manuals, T02, PANTALLA M4, etc.). When this guide and the templates don't settle a styling question, open a comparable finished document there instead of guessing — it is the more authoritative source since it reflects what has actually shipped. Note that the sensor-side "Manual_Tecnico_Sensor_Portillon" / "Manual_Tecnico_Comandos_Portillon" / "S_AlertaPermitetral" documents in `SENSORES/Sensor portillon/` cover the MMCall V.133 hardware and its raw serial command set — they are a different document from an operating manual for the Diag-TOOL app itself; don't conflate the two when generating either one.
