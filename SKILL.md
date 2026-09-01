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

HTML slides must match the custom glassmorphism style sheet:

### Layout and Scaling
*   **Base Container:** Fit to screen with adaptive scaling without forcing limits. Perfect for 32" TV outputs in the field.
*   **Card Container:** Glassmorphic layout (`backdrop-filter: blur(16px)`).
*   **Card Text Scroll:** Any list description panel containing dense text (e.g., button mapping or list of codes) must have:
    ```css
    overflow-y: auto;
    max-height: 480px;
    ```
    This prevents truncation when text size is scaled up for short-sighted accessibility.

### Navigation and Interactions
*   **Header:** Standard header with logos and a right-aligned compact session badge (`0.85rem`).
*   **Keyboard Controls:** Map arrow keys `ArrowLeft` / `ArrowRight` to change slides.
*   **Image Lightbox:** Expandable screenshot overlay when clicking images. Keyboard focusable (`tabindex="0"`) and triggerable via `Enter` / `Space`, closeable via `Escape`.
*   **Interactive Quiz:** Include a 5-question evaluation quiz on the final slide with instant option validation and a final score summary container.

--------------------------------------------------------------------------------

## 3. Templates and Automation Scripts

Use the template scripts stored in the skill resources to build new deliverables quickly:
*   [template_manual.py](./resources/template_manual.py) - Python script using `docx` and `docx2pdf` to compile manuals.
*   [template_presentation.html](./resources/template_presentation.html) - Base boilerplate for interactive slideshows.
