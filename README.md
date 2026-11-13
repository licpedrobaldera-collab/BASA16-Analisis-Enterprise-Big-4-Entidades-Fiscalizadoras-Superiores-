# Sistema de Auditoría Forense y Supervisión CCRD (Ley 18-24)

Aplicación Full-Stack en **React 19 + TypeScript + Express + Vite + Google GenAI (`@google/genai`)** diseñada para la evaluación forense de expedientes reales (Acta de Lectura, Informe Preliminar, Réplicas, Contraréplicas e Informe Final).

## Características de Producción (Cero Datos Simulados)
- **Extracción Real Multiformato**: Procesa archivos reales `.pdf` (`unpdf`), `.docx` (`mammoth`), `.xlsx/.csv` (`xlsx`) y `.txt/.json` con generación de huella digital `SHA-256` por archivo.
- **Motor Forense Ley 18-24 CCRD & ISSAI**: Evalúa cada capítulo y hallazgo en escala `1/10` y porcentaje, contrastando *Lo que Existe vs. Recomendación 10/10*, base legal y procedimiento aplicado.
- **Segregación por Directores y Búsqueda de Decretos**: Usa Grounding con Google Search para localizar decretos oficiales de nombramiento y fuentes verificables.
- **Exportación Directa**: Descarga en Excel con **cada hallazgo en una columna separada** (`Matriz_Hallazgos_x_Columna`) y dictamen ejecutivo en Word.

## Instalación y Ejecución
1. Instalar dependencias:
   ```bash
   npm install
   ```
2. Configurar variables de entorno:
   ```bash
   cp .env.example .env
   # Editar .env y agregar su GEMINI_API_KEY
   ```
3. Ejecutar en desarrollo:
   ```bash
   npm run dev
   ```
4. Compilar y ejecutar en producción:
   ```bash
   npm run build
   npm start
   ```
