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

All printable manuals and guides must adhere to the following formatting specifications:

### Page Layout and Geometry
*   **Margins:** Exact 1.0 inch (2.54 cm) on all four sides.
*   **Page Setup:** First page header/footer must be different (disabled on the cover page, enabled on pages 2+).

### Fonts and Colors
*   **Primary Font:** `Arial` (Charcoal `#2C3E50` for general text, 10.5pt, line spacing 1.15, space after paragraph 6pt).
*   **Primary Accent Color:** Navy Blue `#1F3A8A` (used for main titles and Heading 1).
*   **Secondary Accent Color:** Lila Purple `#9C89F5` (used for "MMCALL ACADEMY" and Heading 2).
*   **Success Color:** Green `#27AE60` (used for seal certifications).

### Headers and Footers (Pages 2+)
*   **Header:** A 1x2 grid table. 
    *   Left cell: Empty.
    *   Right cell: Aligned right, 8.5pt Arial, Gray `#7F8C8D` text: `MMCALL ACADEMY  |  Manual Técnico - [Nombre del Equipo]`
*   **Footer:** Aligned left, 8.5pt Arial, Gray `#7F8C8D` text: `S&D Mmcall Paging Services LTDA.`

### Document Cover Page Structure
1.  **Header Title:** `PROGRAMA DE CAPACITACIÓN ACADÉMICA` (Arial 13pt, bold, Navy Blue `#1F3A8A`, centered).
2.  **Subtitle:** `MMCALL ACADEMY` (Arial 22pt, bold, Lila `#9C89F5`, centered).
3.  **Logo:** Center-aligned image `logo_mmcall_nobg.png` (width 2.5 inches).
4.  **Main Title:** `Manual Técnico - [Nombre del Dispositivo]` (Arial 18pt, bold, Navy Blue `#1F3A8A`, centered).
5.  **Subtopic:** `[Detalle de Configuración o Método]` (Arial 12pt, Gray `#7F8C8D`, centered).
6.  **Quality Certification:** `Sello de Calidad MMCall Academy` (Arial 11pt, bold, italic, Green `#27AE60`, centered).
7.  **Metadata Footer:** `Soporte y Desarrollo (S&D) - MMCall Paging Services LTDA. - Chile` (Arial 9.5pt, Gray `#7F8C8D`, centered).

### Body Styling
*   **Headings:**
    *   **Heading 1:** 15pt, bold, Navy Blue, space before 18pt, space after 6pt.
    *   **Heading 2:** 12.5pt, bold, Lila Purple, space before 12pt, space after 4pt.
*   **Bullet Points:** Use standard list bullets with a bold prefix run for clarity (e.g., `• **1. Paso:** Explicación.`).
*   **Zebra Tables:** Centered tables, white text on `#1F3A8B` background for header row. Alternate data rows with `#F2F4F7` background. Custom cell padding: 100 dxa top/bottom, 120 dxa left/right.
*   **Images:** Centered, width 4.5 inches. Include a caption below in Arial 9pt italic gray: `Figura: [Descripción]`.
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
