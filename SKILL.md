---
name: mmcall-document-generator
description: >-
  Generates, structures, and standardizes MMCall Academy deliverables: Word (.docx) & PDF manuals, Speaker Guides ("Torpedos Internos"), PowerPoint (.pptx) presentations, interactive 3D HTML slideshows with emulators/simulators, and Full HD videos narrated with Victor Cayupil's official cloned voice.
---

# MMCall Academy — Sistema Integral de Documentación, Capacitación y Medios

Guía oficial para el diseño, desarrollo, estructuración y despliegue de los entregables de **MMCall Academy (Soporte y Desarrollo S&D)**. Este estándar garantiza una identidad corporativa unificada, rigurosa y de vanguardia técnica a lo largo de las 6 disciplinas del ecosistema.

---

## 1. Manuales Técnicos y de Operación (Word `.docx` y `.pdf`)

### Geometría y Paginación
*   **Márgenes:** Exactamente 1.0 pulgada (2.54 cm) en los cuatro bordes.
*   **Encabezados/Pies Diferenciados:** Portada sin encabezado ni pie de página; activos desde la página 2 en adelante.
*   **REGLA CRÍTICA — Un Solo Salto de Página:** En todo el documento debe existir **EXACTAMENTE UN** `doc.add_page_break()`, ubicado inmediatamente al final de la portada para pasar a la página 2.
*   **Cero Páginas Vacías o Huérfanas:** PROHIBIDO insertar saltos de página antes de secciones, tablas o imágenes. Word maneja el flujo automáticamente configurando `keep_with_next = True` en todos los títulos.
*   **Espaciados Compactos:**
    *   Heading 1: `space_before = Pt(14)`, `space_after = Pt(4)`, `keep_with_next = True`.
    *   Heading 2: `space_before = Pt(8)`, `space_after = Pt(3)`, `keep_with_next = True`.
    *   Párrafos y Viñetas: `space_after = Pt(3)` a `Pt(4)`, interlineado `1.15`.
    *   Tablas: `space_after = Pt(4)` a `Pt(6)`.

### Tratamiento de Imágenes y Screenshots
*   **Capturas Horizontales (16:9 / Escritorio):** Ancho máximo de `Inches(4.2)` a `Inches(4.5)`.
*   **Capturas Verticales de Móvil / Smartphone:**
    *   **NUNCA usar ancho 4.5 pulgadas:** Genera alturas mayores a 9" forzando páginas en blanco.
    *   **Diseño Paso a Paso en 2 Columnas (Estándar Oficial):** Crear una tabla de 2 columnas sin bordes. En la celda izquierda (65% del ancho) se ubica la explicación del paso, y en la celda derecha (35% del ancho) se inserta la captura del celular con altura controlada (`Inches(3.2)` a `Inches(3.5)`).

### Paleta de Colores y Tipografías Verificadas
```
Rojo Primario      #E30613   (RGB: 227, 6, 19)      Título principal, encabezados H1, cabeceras de tabla
Lila MMCall        #9C89F5   (RGB: 156, 137, 245)   Palabra "ACADEMY", eyebrow de portada
Gris Oscuro        #404040   (RGB: 64, 64, 64)      Palabra "MMCALL", títulos H2, prefijos de viñetas
Gris Medio         #7F8C8D   (RGB: 127, 140, 141)   Encabezado/pie de página recurrente (pág 2+)
Gris Claro         #808080   (RGB: 128, 128, 128)   Subtítulo de portada, Sello de Calidad
Texto del Cuerpo   #333333   (RGB: 51, 51, 51)      Fuente Calibri 11pt, interlineado 1.15
```

---

## 2. Guiones del Presentador / "Torpedos Internos" (Word `.docx` y `.pdf`)

El **Guion del Presentador ("Torpedo")** es el instrumento táctico para que el relator o técnico comercial ejecute una demostración impecable, dominando el discurso, las acciones en vivo y las respuestas ante objeciones.

### Estructura Mandatoria del Documento
1.  **Portada Interna:**
    *   Eyebrow Monospace: `GUION DEL PRESENTADOR · USO INTERNO` (Consolas 10pt bold, Lila `#9C89F5`).
    *   Brand Line: `MMCALL ` (Calibri 26pt bold `#404040`) + `ACADEMY` (Calibri 26pt bold `#9C89F5`).
    *   Logo institucional centrado.
    *   Título Principal: `Guion del Presentador «[Nombre del Módulo]»` (Calibri 24pt bold, Rojo `#E30613`).
    *   Aviso de Confidencialidad: *"Documento interno para relatores y técnicos comerciales. Úsalo en tu celular o impreso. NO proyectar ni compartir con la audiencia."*
    *   Un único salto de página hacia la página 2.
2.  **Sección 1 — Antes de la Demostración:**
    *   Checklist técnico: Archivos desconectados (100% offline), navegador recomendado (Chrome/Edge), modo pantalla completa (`F`), prueba previa de simuladores y visualización en móvil.
3.  **Sección 2 — Cronología del Recorrido:**
    *   Tabla resumen con: Parada, Módulo, Tiempo Sugerido y Minuto Acumulado.
    *   Especificación de versión completa (~25-30 min) y versión express (~15 min).
4.  **Sección 3 — Guion Parada por Parada (Tarjetas de 2 Columnas):**
    *   Cada parada se estructura dentro de una tarjeta con borde rojo izquierdo de 2 columnas:
        *   **Columna Izquierda (35%):** Badge de parada (`PARADA X / Y`, tiempo y minuto) y miniatura de la diapositiva/interfaz.
        *   **Columna Derecha (65%):**
            *   **Qué se ve:** Descripción concreta de los elementos desplegados en pantalla.
            *   **Qué decir:** Discurso sugerido en primera persona, fluido, convincente y con datos duros.
            *   **Qué hacer:** Acciones en vivo (rotar visor 3D, cambiar precio en simulador, activar Pick-by-Light LED, mover sliders de ROI, etc.).
            *   **Mensaje clave:** Concepto comercial o técnico de alto impacto que debe fijarse en la audiencia.
5.  **Sección 4 — Si Preguntan… (Manejo de Objeciones):**
    *   Preguntas complejas del cliente (autonomía de baterías, caídas de internet, costos de licencias, marco legal SERNAC, tiempos de recambio) con respuestas técnicas y comerciales directas.
6.  **Sección 5 — Si Algo Falla en Vivo (Plan de Contingencia):**
    *   Instrucciones rápidas ante proyector descalibrado, recarga accidental de página, lentitud de red o congelamiento de animación.

---

## 3. Presentaciones Oficiales PowerPoint (`.pptx` 16:9 Widescreen)

Formato oficial para reuniones corporativas, comités de compras y capacitaciones presenciales bajo el estándar **MMCall Dark Studio**.

### Dimensiones y Parámetros
*   **Relación de Aspecto:** 16:9 Widescreen (`slide_width = Inches(13.333)`, `slide_height = Inches(7.5)`).
*   **Fondo:** Sólido Deep Dark `#070A10` (RGB: 7, 10, 16).
*   **Cabecera:**
    *   Línea superior de acento Rojo `#E30613` de 8px (0.08").
    *   Logo corporativo `logo_mmcall_nobg.png` a la izquierda.
    *   Eyebrow en Lila `#9C89F5` (Outfit 11pt bold).
    *   Título de diapositiva en Blanco `#F8FAFC` (Outfit 22pt bold).
    *   Subtítulo en Gris Muted `#94A3B8` (Calibri 13pt).
    *   Línea divisoria horizontal sutil `#34415C`.
    *   Badge de Diapositiva: Rectángulo redondeado `#161E30` con borde Lila `#9C89F5` y contador en Consolas bold (`X / Y`).
*   **Layout de Contenido en 2 Columnas:**
    *   **Columna Izquierda (Tarjeta de Texto):** Ancho 6.8", fondo `#101624`, borde `#34415C`, franja de acento superior de color temático, título de sección en mayúsculas, y lista de viñetas con prefijo destacado en negrita blanca.
    *   **Columna Derecha (Tarjeta Gráfica):** Ancho 5.2", contenedor centrado para imágenes de hardware o capturas de interfaz en alta resolución, con barra inferior oscura y caption técnico en Cyan `#00F0FF` (Consolas).
*   **Pie de Página:**
    *   Franja inferior `#04060B` con borde superior.
    *   Texto institucional: `MMCall Academy — Programa de Capacitación Técnica Oficial (S&D)   |   [Módulo]`.

---

## 4. Presentaciones Web Interactivas 3D (`.html` Autocontenidas)

Presentaciones dinámicas ejecutables en cualquier navegador sin dependencias de internet ni servidores locales.

> [!CRITICAL] REGLA MANDATORIA DE FORMATO HTML (SLIDES 16:9):
> 1. **Estructura Estricta en Diapositivas:** Todo archivo HTML de presentación DEBE SER SIEMPRE una presentación por diapositivas (Slides 16:9 con `.slide.active`), con navegación secuencial paso a paso mediante botones Anterior/Siguiente, Dots indicadores en cabecera y atajos de teclado (flechas izquierda/derecha, espacio).
> 2. **PROHIBIDO EL FORMATO VERTICAL (SCROLL / LISTA):** Está ESTRICTAMENTE PROHIBIDO renderizar la presentación como una página web vertical continua (scroll page, landing page o lista de secciones). Las capacitaciones y demostraciones se presentan diapositiva a diapositiva.
> 3. **Confidencialidad Absoluta del Torpedo (Prohibido a la Vista del Cliente):**
>    * El **cliente o la audiencia NUNCA debe ver el torpedo ni saber que existe un guion**.
>    * En la pestaña de proyección/visor (`presentacion_autocontenida.html`), **NO debe existir ningún botón visible que diga "Torpedo", "Guion" ni cajones de notas desplegados a la vista de la audiencia**. La interfaz visible debe ser 100% limpia, pulcra y comercial.
>    * El torpedo pertenece **exclusivamente a la segunda pantalla privada del relator** (`visor_control_torpedo.html` o `torpedo_presentador.html`), sincronizado en segundo plano vía BroadcastChannel / postMessage. Si se requiere un botón para abrir la consola, debe ser neutro y discreto (ej: `[ 🎛️ Consola Privada ]` o atajo silencioso) y jamás mostrar textos internos, "Qué decir" ni respuestas a objeciones en la pantalla principal.
> 4. **Nombre Oficial Obligatorio:** En la raíz de cada módulo de producto, el archivo debe llamarse exactamente: `presentacion_autocontenida.html`.
> 5. **Arquitectura de Doble Pantalla / Presenter View (Proyección en Sala vs. Visor de Control Remoto):**
>    * Toda presentación oficial de capacitación debe operar bajo el sistema desacoplado de dos pantallas/pestañas:
>      * **Pestaña de Proyección (`presentacion_autocontenida.html`):** Vista limpia 16:9 proyectable para la audiencia en proyector o pantalla extendida, con botón `[ 🎛️ Visor Control [P] ]`, capa de Blackout (`B`) y enlace de telemetría.
>      * **Visor de Control y Torpedo (`visor_control_torpedo.html`):** Consola privada para el relator (en su laptop o pantalla secundaria). Incluye cronómetro digital, reloj real, estado de sincronización (`🟢 Enlace Activo`), monitores duales (slide actual + preview de próxima parada), matriz de salto directo a slides (1 al 12), botón de blackout y el guion técnico completo ("Qué ver", "Qué decir", "Qué hacer", "Mensaje clave", "Si Preguntan...").
>      * **Sincronización Bidireccional Redundante:** Enlace instantáneo mediante `BroadcastChannel` nativo + evento `storage` de `localStorage` + `postMessage`. El avance con clicker en el proyector o clic en la consola actualiza ambas vistas al unísono sin servidores locales ni internet.

### Componentes Clave y Estándar Visual

1.  **Portada Hero Institucional (Slide 1 de Alto Impacto):**
    *   **PROHIBIDO** iniciar la presentación con una lista plana de viñetas genéricas.
    *   **Estructura del Hero Stage:**
        *   `eyebrow`: Píldora translúcida con borde acento lila (`PROGRAMA DE CAPACITACIÓN OFICIAL · SyD`).
        *   `h1`: Tipografía `Outfit` en gradiente de alto contraste (`linear-gradient(135deg, #ffffff 0%, #cbd5e1 60%, #9c89f5 100%)`).
        *   `kpi-row` / `stats-ribbon`: Fila de tarjetas KPI glassmorphism destacando 3 a 4 métricas duras (ej. `220V AC`, `433.500 MHz`, `16 Audios Máx`, `EEPROM No Volátil`).
        *   `hero-stage`: Escenario 3D con anillos orbitales luminiscentes animados (`.hero-ring`, `.hero-ring-2`), chasis del equipo con animación de levitación suave (`@keyframes float`) y sombra de profundidad.
        *   Botón de llamada a la acción: `[ 🚀 Iniciar Capacitación Técnica ]` con síntesis sonora de acorde armónico Web Audio API.

2.  **Showroom / Visor 3D de Hardware y Tilt Giroscópico:**
    *   **Contenedor Perspectiva:** `.viewer-3d-wrap` con `perspective: 1200px;` y alineación centrada.
    *   **Chasis Tridimensional:** `.device-3d-card` con `transform-style: preserve-3d; transition: transform 0.15s ease-out;` y reflejo dinámico (*glare*) mediante pseudoelemento `::after` con gradiente lineal reflectivo.
    *   **Script de Tilt Giroscópico:**
        ```javascript
        window.addEventListener('mousemove', e => {
          const card = document.getElementById('device3D');
          if (!card || !card.offsetParent) return;
          const rect = card.getBoundingClientRect();
          const cx = rect.left + rect.width / 2;
          const cy = rect.top + rect.height / 2;
          const dx = (e.clientX - cx) / (window.innerWidth / 2);
          const dy = (e.clientY - cy) / (window.innerHeight / 2);
          card.style.transform = `rotateY(${dx * 22}deg) rotateX(${-dy * 22}deg)`;
        });
        ```
    *   **Efectos 3D Flip (Tarjetas de Volteo):** En secciones de preguntas u objeciones, utilizar `.flip-card` con `transform-style: preserve-3d;` y `.flip-card-inner` con `transform: rotateY(180deg);` y `backface-visibility: hidden;` para revelar respuestas interactivas en 3D.

3.  **Fondo Holográfico 3D y Luces Atmosféricas:**
    *   Orbes difusos de resplandor luminiscente en segundo plano (`.glow-orb`, `.glow-1`, `.glow-2` con desenfoque de 90px en rojo y lila).
    *   Canvas animado (`#bgCanvas`) con partículas que flotan en el espacio tridimensional y se vinculan estelarmente al pasar el ratón.

4.  **Simuladores de Ecosistema en Vivo:**
    *   Simulador de hardware: Comportamiento dinámico, halo volumétrico y activación de domo/LED multicolor (7 colores físicos).
    *   Reproducción acústica en vivo con archivos WAV reales embebidos en Base64.

5.  **Calculadora Interactiva de Retorno de Inversión (ROI):**
    *   Sliders reactivos que permiten ajustar número de equipos, horas/hombre y frecuencia de eventos, calculando en tiempo real el ahorro en CLP$ y el periodo de amortización.

6.  **Cuestionario Evaluativo (Quiz):**
    *   Preguntas técnicas interactivas con alternativas dinámicas.
    *   Efectos sonoros generados dinámicamente mediante la API nativa de audio del navegador (`AudioContext` / Web Audio API) sin archivos de audio externos.
    *   Retroalimentación visual instantánea (verde/rojo) y pantalla de diploma/puntaje acumulado.

7.  **Empaquetado Autocontenido:** Todas las imágenes, íconos y logos deben estar codificados en cadenas Base64 en línea (`data:image/png;base64,...`), logrando un único archivo `.html` portable.

---

## 5. Videos de Capacitación Full HD 1080p con Voz Oficial Clonada

Producción de videos educativos de alta fidelidad narrados con la voz oficial de **Victor Cayupil** generada localmente mediante **VoiceStudio / OmniVoice**.

### Diferenciación Mandatoria de Videos por Audiencia
*   **Video Técnico (Soporte & Desarrollo):** Enfocado en arquitectura de conexionado, diagramas de bloques, bornas (220V/12V), asignación de microswitchs / RF, diagnósticos de fallas y configuración.
*   **Video de Ventas / Marketing:** Enfocado en la propuesta de valor comercial, beneficios operativos, retorno de inversión (ROI) y casos de uso en salud, retail o industria. **PROHIBIDO** incluir pantallas complejas de software o jerga técnica excesiva; el cliente busca resolver un problema de negocio.
*   **Nombres de Clientes Genéricos:** PROHIBIDO utilizar nombres reales de clientes en videos, presentaciones o manuales públicos. Utilizar siempre nombres genéricos de fantasía ("Clínica Metropolitana", "Hospital Regional", "Centro de Salud Norte").

### Pipeline de Producción
1.  **Voz de Referencia:** `voices/917e1e26-54629cc7.wav` recortada a exactamente 8.0 segundos (`_ref_victor_8s.wav`), mono a 24000 Hz.
2.  **Motor de Inferencia:** `k2-fsa/OmniVoice` cargado en el entorno local de VoiceStudio:
    ```python
    model = OmniVoice.from_pretrained("k2-fsa/OmniVoice", device_map=device, dtype=getattr(torch, tts_dtype_name(device)))
    prompt = model.create_voice_clone_prompt(ref_audio=ref_wav, ref_text=REF_TEXT_CLEAN)
    ```
3.  **Caché Modular de Audio:** Los textos se dividen en trozos (chunks) de 25-30 palabras por escena. Los audios sintetizados se almacenan en `video_audio_cache/` (`s1_intro_c1.wav`, etc.). Si el archivo ya existe en caché, el pipeline salta la inferencia reduciendo el tiempo de re-ejecución a cero.
4.  **Generación de Frames Gráficos:** Imágenes maestras de 1920x1080 píxeles renderizadas con Pillow (`PIL`), aplicando el fondo degradado de estudio, tarjetas oscuras, tipografías del sistema y capturas en alta fidelidad.
5.  **Ensamblado y Concatenación con FFmpeg:**
    *   Cada escena se renderiza como un segmento MP4 sincronizado con la duración exacta del audio:
        ```bash
        ffmpeg -y -loop 1 -i frame.png -i scene.wav -c:v libx264 -tune stillimage -c:a aac -b:a 192k -pix_fmt yuv420p -t [dur] -shortest seg.mp4
        ```
    *   Los segmentos se unen sin re-codificación (`-c copy`) mediante el demuxer `concat` de FFmpeg, produciendo el archivo master `Video_Capacitacion_[Tema].mp4`.

---

## 6. Normas Corporativas de Integridad, Datos y Gráfica

1.  **PROHIBICIÓN TOTAL DE DATOS INVENTADOS O FICTICIOS:**
    *   ESTÁ ESTRICTAMENTE PROHIBIDO inventar o asumir afirmaciones comerciales de inventario no confirmadas (ej. "STOCK PERMANENTE EN SANTIAGO DE CHILE", "LABORATORIO DE INTEGRACIÓN LOCAL"), correos electrónicos o números telefónicos no provistos explícitamente por el usuario.
2.  **PARÁMETROS DE HARDWARE Y SIRENAS:**
    *   En balizas, luces y timbres de alarma: PROHIBIDO fijar decibeles (dB) arbitrarios o inventados. Especificar siempre de forma fidedigna: *"Sonidos predefinidos y sonidos personalizados (14 máximo)"*.
3.  **TIPOGRAFÍA CORPORATIVA EN ESQUEMAS Y DIAGRAMAS:**
    *   En diagramas arquitectónicos de conexión y esquemas CAD (formato PNG 1920x1080 o visualizadores): utilizar siempre la tipografía oficial **Segoe UI / Segoe UI Bold** (y Consolas para terminales/datos técnicos). Prohibido usar fuentes desalineadas como Calibri en gráficos técnicos de estudio oscuro.

---

## 7. Organización de Repositorios y Norma Zoho WorkDrive

Para garantizar una navegación ordenada y lista para sincronización en la nube corporativa (Zoho WorkDrive), el repositorio central `REPOSITORIO MMCALL ACADEMY` se organiza por marcas y familias de equipos:

```
REPOSITORIO MMCALL ACADEMY/
├── 01_MMCALL_HOSPITALARIO/
│   └── Luces_Hospital_N131/
│       ├── Manual_Configuracion_Luces_Hospital_N131_SyD.docx / .pdf
│       ├── Ficha_Comercial_Marketing_Luces_Hospital_N131.docx / .pdf
│       ├── presentacion_autocontenida.html
│       ├── Video_Tecnico_SyD_Luces_Hospital_N131.mp4
│       ├── Video_Marketing_Ventas_Luces_Hospital_N131.mp4
│       ├── esquema_arquitectura_rf_luz_n131_220v.png
│       └── _herramientas/
├── 02_MMCALL_PAGERS_Y_RELOJES/
│   ├── Relojes_NS818/
│   └── Pagers_Discos/
├── 03_MMCALL_PANTALLAS_LED/
│   ├── Pantalla_M4/
│   └── Pantalla_K2/
├── 04_MMCALL_TRANSMISORES/
│   ├── Transmisor_T02/
│   └── Transmisores_CA_TA_Pulsadores/
├── 05_MMCALL_SENSORES_E_IOT/
│   └── Sensores_Inalambricos/
│       ├── Sensor_Portillon/
│       ├── Sensor_Movimiento_PIR/
│       ├── Sensor_Puerta_Magnetico/
│       └── Sensor_Tamper_Presion/
├── 06_MINEW_SUNPAI_ESL/
│   ├── 01_General_y_Ventas_ESL/
│   ├── 02_Gateway_Minew_ESL/
│   └── 03_ESL_Modo_Standalone_Celular/
├── 07_MMCALL_SOFTWARE_Y_CENTRALES/
│   ├── Central_Windows_Industrial/
│   ├── MMCall_Android_Gateway/
│   └── Diag_Tool_Software/
└── 08_CORPORATIVO_Y_METODOLOGIAS/
    ├── Presentaciones_Oficiales/
    ├── Metodos_y_Procedimientos/
    └── Articulos_Tecnicos_USA/
```

### Reglas de Higiene de Directorios
*   **Cero Carpetas Huérfanas en la Raíz:** Prohibido dejar carpetas genéricas o descontextualizadas en la raíz del repositorio (ej. `Imagenes proyectos`, `Otros`, etc.). Todo activo debe residir dentro de su familia de producto o en `08_CORPORATIVO_Y_METODOLOGIAS`.
*   **Raíz Impecable en Módulos:** En la raíz de cada módulo solo deben residir los archivos finales de entrega para el cliente y el equipo técnico.
*   **Prohibido Archivos Sueltos:** Todos los scripts auxiliares (`build_*.py`, `capture_*.py`, caches de video) deben almacenarse exclusivamente dentro de `_herramientas/`.
*   **Preservación Histórica:** Ningún archivo previo debe eliminarse arbitrariamente; todo borrador o versión previa debe archivarse ordenadamente en `Borradores_Anteriores/` o `Presentación antigua/`.
*   **Compatibilidad de Rutas:** Si scripts o accesos heredados requieren rutas antiguas (ej. `SyD/`), se deben mantener enlaces de directorio (Directory Junctions) de NTFS para compatibilidad total sin duplicar almacenamiento.

---

## 8. Scripts y Plantillas de Recursos

En la carpeta `resources/` de esta skill se encuentran disponibles los scripts base listos para ejecutar:
*   `template_manual.py`: Generador de manuales técnicos en Word y PDF con soporte para capturas móviles de 2 columnas y tablas cebra.
*   `template_guion_presentador.py`: Generador del "Torpedo" / Guion del presentador con tarjetas parada por parada, cronología y manejo de objeciones.
*   `template_pptx.py`: Generador de presentaciones PowerPoint 16:9 bajo el estándar MMCall Dark Studio.
*   `template_presentation.html`: Plantilla base HTML con partículas holográficas, visores interactivos, diapositivas 16:9 secuenciales y quiz con Web Audio.
*   `template_video_narrated.py`: Pipeline completo de locución neuronal con la voz de Victor Cayupil y renderizado de video 1080p con FFmpeg.

---

## 9. Framework B2B de Marketing, Ventas y Sales Enablement (MMCall Academy)

Para asegurar que todo entregable de MMCall Academy (manuales, torpedos, diapositivas y videos) maximice su impacto comercial y facilite el cierre de ventas corporativas, se deben aplicar de forma obligatoria los siguientes pilares derivados del estándar B2B de *Conversion Factory / Corey Haines* (disponibles en la subcarpeta `b2b_marketing/`):

### 1. Fichas de Batalla ("Battle Cards") y Manejo de Objeciones (`sales-enablement`)
* **Comparativas Objetivas frente a la Competencia:** Estructurar matrices de criterios duros (Costo inicial, método de conexión, velocidad de refresco, vida útil de baterías, cumplimiento SERNAC) frente a fabricantes globales (Pricer, SES-imagotag, Hanshow, SoluM).
* **Manejo Proactivo de Objeciones (Sección Obligatoria en Torpedos):**
  * *¿Por qué cambiar etiquetas si el papel es barato?* -> Cuantificar el costo oculto de horas/hombre de cambio manual de flejes, errores humanos de digitación y las severas multas de la Ley del Consumidor (SERNAC) por discrepancia de precio entre góndola y caja.
  * *¿Qué pasa si se cae la red o el internet?* -> Explicar la autonomía del hardware bi-estable E-ink: la pantalla retiene la imagen y el precio de forma permanente sin energía; la modalidad StandAlone funciona 100% offline con el teléfono celular vía BLE 5.0.
  * *¿Por qué empezar con MMCall?* -> Escalabilidad garantizada: el cliente puede adquirir 50 etiquetas en modo StandAlone hoy y sumar Gateways Minew G1 mañana sin tener que cambiar el hardware.

### 2. Matriz de Propuesta de Valor y Posicionamiento (`product-marketing`)
* **Traducción de Especificación Técnica a Beneficio de Negocio:**
  * En lugar de: *"Pantalla E-ink bi-estable de 2.9 pulgadas con chip Bluetooth 5.0"*
  * Redactar: *"Pantalla de tinta electrónica tipo papel de 2.9 pulgadas: 180° de ángulo de visión sin reflejos, cero consumo de energía mientras la imagen está congelada, y actualización directa en pasillo en 3 a 5 segundos con el teléfono del operador."*
* **Segmentación por Audiencia (ICP):**
  * **Gerentes de Operaciones y Retail:** Énfasis en reducción de mermas, optimización de tiempos de reposición y preparación rápida de pedidos (*Pick-by-Light LED* multicolor).
  * **Gerentes de Finanzas / Compras:** Retorno de la inversión (ROI amortizable en 6 a 12 meses) y protección de inversión escalable.
  * **Jefes de Tienda y Operadores:** Facilidad de uso, cero cables en góndola y aplicación móvil intuitiva en español.

### 3. Redacción Persuasiva y Control Editorial (`copywriting` & `copy-editing`)
* **PROHIBICIÓN TOTAL DE "AI TELLS" (Clichés de Inteligencia Artificial):**
  * Está estrictamente prohibido usar frases vacías o típicas de LLM como:
    * *"En un mundo digital en constante evolución..."*
    * *"No es solo una solución, es una revolución..."*
    * *"Descubra una experiencia sin fisuras (*seamless*)..."*
    * *"Sumérjase en la vanguardia tecnológica..."*
* **Regla de Oro:** Utilizar voz activa, oraciones directas, datos duros y beneficios tangibles. Cada afirmación comercial debe ser demostrable técnicamente.

### 4. Estructura de Guiones de Video Comercial (`edit-anatomy`)
Todo video comercial o promocional de producto producido por MMCall Academy debe seguir la cronología modular de 6 actos:
1. **Acto 1 — El Gancho y la Problemática (0 a 18s):** El dolor del retail tradicional (costos de papel, demoras y multas SERNAC).
2. **Acto 2 — La Solución Ágil (18 a 35s):** Presentación del producto base (ej. Modo StandAlone con Cloud Tag BLE 5.0).
3. **Acto 3 — La Solución Centralizada (35 a 54s):** Automatización masiva y conectividad cloud (ej. Modo Online con Gateway G1 y LED Pick-by-Light).
4. **Acto 4 — Cuadro Comparativo de Decisión (54 a 72s):** Matriz visual de criterios estratégicos según el tamaño del local.
5. **Acto 5 — Escalabilidad y Protección de la Inversión (72 a 90s):** Garantía de compatibilidad de hardware entre ambas modalidades.
6. **Acto 6 — Respaldo Corporativo y Llamado a la Acción (90 a 108s):** Garantía, soporte local chileno, integración con ERP/POS y datos de contacto oficiales.

### 5. Estructuración de Ofertas Piloto de Bajo Riesgo (`offers`)
* Facilitar el cierre comercial ofreciendo programas piloto (*PoC - Proof of Concept*):
  * Kit de evaluación de 50 etiquetas + App móvil StandAlone sin costo de servidores ni cableado.
  * Garantía de migración transparente hacia red de Gateways Online.

