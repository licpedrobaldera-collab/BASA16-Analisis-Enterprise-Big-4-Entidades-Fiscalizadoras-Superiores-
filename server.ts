import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '50mb' }));

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

const FORENSIC_SYSTEM_INSTRUCTION = `Eres un Auditor Forense Senior (nivel Big-4 y Supervisor de Auditoría de la Cámara de Cuentas de la República Dominicana - CCRD).
Reglas estrictas de producción:
1. PROHIBIDO inventar, simular o rellenar datos demo/ficticios. Todo análisis debe basarse ÚNICAMENTE en el texto real extraído del documento cargado por el usuario y en legislación real verificable. Si un dato (como fecha o monto) no consta en el documento, indica explícitamente "No especificado en el documento".
2. Evalúa el documento según su fase procesal (Acta de Lectura, Informe Preliminar, Réplica/Contraréplica o Informe Final) bajo la Ley núm. 18-24 de la Cámara de Cuentas de la República Dominicana, su Reglamento de Aplicación General (Res. ADM-2025-040) y las normas ISSAI.
3. Asigna puntuación real de 1 a 10 y porcentaje (0-100%) al documento global y a cada capítulo/punto evaluado.
4. Para cada hallazgo o incidencia detectada en el texto real, desglosa:
   - Qué existe actualmente en el documento (cita/resumen fiel).
   - Brecha técnica o legal detectada.
   - Recomendación exacta para llevarlo a 10/10.
   - Base legal y normativa exacta (Ley 18-24, Ley 340-06, Ley 10-07, ISSAI).
   - Procedimiento forense para determinar y sustentar lo razonado.
   - Período y titular/director responsable según el documento y búsqueda de decreto oficial aplicable.
   - Si se evalúa una Réplica/Contraréplica: monto original, monto descargado validado con evidencia y monto persistente.`;

app.post('/api/analyze', async (req, res) => {
  try {
    const { documents, phase, entityName, auditPeriod, previousContext } = req.body;

    if (!process.env.GEMINI_API_KEY) {
      return res.status(500).json({ error: 'Falta configurar GEMINI_API_KEY en el entorno (.env).' });
    }

    if (!documents || !Array.isArray(documents) || documents.length === 0) {
      return res.status(400).json({ error: 'Debe cargar al menos un archivo real con contenido extraído.' });
    }

    const combinedText = documents
      .map((d: any, idx: number) => `=== DOCUMENTO REAL #${idx + 1}: ${d.name} (${d.type}) ===\n${d.content}`)
      .join('\n\n');

    const prompt = `Analiza con rigor forense el siguiente expediente real cargado por el usuario:
- Entidad Auditada: ${entityName || 'Extraer del documento'}
- Período(s) Auditado(s): ${auditPeriod || 'Extraer del documento'}
- Fase Procesal Actual: ${phase}
${previousContext ? `- Contexto / Iteración Previa para Seguimiento de Réplicas: ${JSON.stringify(previousContext)}` : ''}

CONTENIDO REAL DE LOS DOCUMENTOS CARGADOS:
${combinedText.slice(0, 850000)}

Devuelve EXCLUSIVAMENTE un JSON estructurado siguiendo el esquema solicitado, basado 100% en los datos reales proporcionados.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-pro',
      contents: prompt,
      config: {
        systemInstruction: FORENSIC_SYSTEM_INSTRUCTION,
        temperature: 0.1,
        tools: [{ googleSearch: {} }],
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            entidad: { type: Type.STRING },
            faseEvaluada: { type: Type.STRING },
            puntuacionGlobal10: { type: Type.NUMBER },
            porcentajeBlindaje: { type: Type.NUMBER },
            resumenEjecutivo: { type: Type.STRING },
            evaluacionPorCapitulos: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  capitulo: { type: Type.STRING },
                  puntuacionActual10: { type: Type.NUMBER },
                  porcentajeActual: { type: Type.NUMBER },
                  loQueExiste: { type: Type.STRING },
                  recomendacion10de10: { type: Type.STRING },
                  justificacionLegal: { type: Type.STRING },
                  procedimientoDeterminacion: { type: Type.STRING }
                },
                required: ['capitulo', 'puntuacionActual10', 'porcentajeActual', 'loQueExiste', 'recomendacion10de10', 'justificacionLegal', 'procedimientoDeterminacion']
              }
            },
            hallazgos: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  codigo: { type: Type.STRING },
                  titulo: { type: Type.STRING },
                  periodoIncidencia: { type: Type.STRING },
                  directorResponsable: { type: Type.STRING },
                  decretoNombramiento: { type: Type.STRING },
                  fuenteOficialDecreto: { type: Type.STRING },
                  puntuacionActual10: { type: Type.NUMBER },
                  porcentajeActual: { type: Type.NUMBER },
                  loQueExisteEnInforme: { type: Type.STRING },
                  brechaDetectada: { type: Type.STRING },
                  recomendacionAjuste10de10: { type: Type.STRING },
                  baseLegalNormativa: { type: Type.STRING },
                  procedimientoForense: { type: Type.STRING },
                  montoOriginalRD: { type: Type.NUMBER },
                  montoDescargadoRD: { type: Type.NUMBER },
                  montoPersistenteRD: { type: Type.NUMBER },
                  estadoReplica: { type: Type.STRING }
                },
                required: [
                  'codigo', 'titulo', 'periodoIncidencia', 'directorResponsable',
                  'decretoNombramiento', 'fuenteOficialDecreto', 'puntuacionActual10',
                  'porcentajeActual', 'loQueExisteEnInforme', 'brechaDetectada',
                  'recomendacionAjuste10de10', 'baseLegalNormativa', 'procedimientoForense',
                  'montoOriginalRD', 'montoDescargadoRD', 'montoPersistenteRD', 'estadoReplica'
                ]
              }
            },
            segregacionResponsabilidad: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  director: { type: Type.STRING },
                  cargo: { type: Type.STRING },
                  decretoYFuente: { type: Type.STRING },
                  periodoGestion: { type: Type.STRING },
                  cantidadHallazgos: { type: Type.NUMBER },
                  montoTotalImputadoRD: { type: Type.NUMBER },
                  montoDescargadoRD: { type: Type.NUMBER },
                  montoPersistenteRD: { type: Type.NUMBER }
                },
                required: ['director', 'cargo', 'decretoYFuente', 'periodoGestion', 'cantidadHallazgos', 'montoTotalImputadoRD', 'montoDescargadoRD', 'montoPersistenteRD']
              }
            }
          },
          required: ['entidad', 'faseEvaluada', 'puntuacionGlobal10', 'porcentajeBlindaje', 'resumenEjecutivo', 'evaluacionPorCapitulos', 'hallazgos', 'segregacionResponsabilidad']
        }
      }
    });

    const parsed = JSON.parse(response.text || '{}');
    const groundingChunks = response.candidates?.[0]?.groundingMetadata?.groundingChunks || [];
    const webSources = groundingChunks
      .filter((c: any) => c.web?.uri)
      .map((c: any) => ({ title: c.web.title, uri: c.web.uri }));

    res.json({ analysis: parsed, webSources });
  } catch (err: any) {
    console.error('Error en /api/analyze:', err);
    res.status(500).json({ error: err.message || 'Error interno en el motor forense.' });
  }
});

if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, 'dist')));
  app.get('*', (_req, res) => {
    res.sendFile(path.join(__dirname, 'dist', 'index.html'));
  });
}

const PORT = Number(process.env.PORT) || 3000;
app.listen(PORT, () => {
  console.log(`Servidor Auditoría Forense CCRD ejecutándose en puerto ${PORT}`);
});
