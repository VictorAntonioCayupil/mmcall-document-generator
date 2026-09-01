# MMCall Academy Document Generator - Claude Instructions

Este repositorio contiene la **Skill Oficial de MMCall Academy** para la generación automática de manuales técnicos (Word `.docx` y PDF) y presentaciones interactivas (HTML5).

## Directiva Principal para Claude

Siempre que el usuario solicite crear, actualizar o formatear un manual, documento técnico o presentación interactiva para MMCall:
1. **Consulta obligatoria:** Lee el archivo `SKILL.md` de este repositorio para cargar todas las directrices de diseño, paleta de colores corporativos, tipografías, márgenes y estructura visual.
2. **Generación de Manuales (Word / PDF):** Utiliza la lógica y las funciones de `resources/template_manual.py` para programar el archivo Word con `python-docx`, asegurando la portada con el logo oficial, encabezados de dos columnas en páginas 2+, pie de página institucional, tablas cebra e inserción de capturas de pantalla de parámetros al inicio de cada sección. Luego compila automáticamente a PDF mediante `docx2pdf`.
3. **Presentaciones Interactivas (HTML5):** Utiliza la plantilla de `resources/template_presentation.html` para crear presentaciones con diseño Glassmorphism, paneles con scroll suave (`overflow-y: auto; max-height: 480px;`), zoom lightbox para capturas y un cuestionario evaluativo de 5 preguntas al final.
