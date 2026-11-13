import * as XLSX from 'xlsx';

export function exportForensicExcel(analysis: any, webSources: any[], documents: any[]) {
  const wb = XLSX.utils.book_new();

  // HOJA 1: MATRIZ TRANSPUESTA (CADA HALLAZGO EN UNA COLUMNA SEPARADA)
  const hallazgos = analysis.hallazgos || [];
  const rowLabels = [
    'Código Hallazgo',
    'Título de la Incidencia',
    'Período de la Incidencia',
    'Director / Titular Responsable',
    'Decreto de Nombramiento',
    'Fuente Oficial / Ubicación del Decreto',
    'Calificación Actual (1/10)',
    'Porcentaje Blindaje Actual (%)',
    'Lo que Existe Actualmente en el Informe',
    'Brecha Técnica / Legal Detectada',
    'Recomendación Exacta para Nivel 10/10',
    'Justificación / Base Legal y Normativa (Ley 18-24 / ISSAI)',
    'Procedimiento Forense de Determinación',
    'Monto Original Observado (RD$)',
    'Monto Descargado Validado (RD$)',
    'Monto Persistente / Responsabilidad (RD$)',
    'Estado Procesal / Réplica'
  ];

  const matrixAoA: any[][] = [];
  const headerRow = ['CRITERIO DE EVALUACIÓN FORENSE', ...hallazgos.map((h: any, i: number) => h.codigo || `HALLAZGO_${i + 1}`)];
  matrixAoA.push(headerRow);

  const fieldKeys = [
    'codigo', 'titulo', 'periodoIncidencia', 'directorResponsable',
    'decretoNombramiento', 'fuenteOficialDecreto', 'puntuacionActual10',
    'porcentajeActual', 'loQueExisteEnInforme', 'brechaDetectada',
    'recomendacionAjuste10de10', 'baseLegalNormativa', 'procedimientoForense',
    'montoOriginalRD', 'montoDescargadoRD', 'montoPersistenteRD', 'estadoReplica'
  ];

  rowLabels.forEach((label, rIdx) => {
    const row = [label, ...hallazgos.map((h: any) => h[fieldKeys[rIdx]] ?? '')];
    matrixAoA.push(row);
  });

  const wsMatriz = XLSX.utils.aoa_to_sheet(matrixAoA);
  XLSX.utils.book_append_sheet(wb, wsMatriz, 'Matriz_Hallazgos_x_Columna');

  // HOJA 2: EVALUACIÓN PUNTO POR PUNTO / CAPÍTULOS
  const capRows = (analysis.evaluacionPorCapitulos || []).map((c: any) => ({
    'Capítulo / Sección': c.capitulo,
    'Puntaje Actual (1/10)': c.puntuacionActual10,
    'Porcentaje (%)': c.porcentajeActual,
    'Lo que Existe': c.loQueExiste,
    'Recomendación 10/10': c.recomendacion10de10,
    'Justificación / Base Legal': c.justificacionLegal,
    'Procedimiento Aplicado': c.procedimientoDeterminacion
  }));
  const wsCap = XLSX.utils.json_to_sheet(capRows);
  XLSX.utils.book_append_sheet(wb, wsCap, 'Evaluacion_Capitulos');

  // HOJA 3: SEGREGACIÓN POR DIRECTORES Y DECRETOS
  const segRows = (analysis.segregacionResponsabilidad || []).map((s: any) => ({
    'Director / Titular': s.director,
    'Cargo': s.cargo,
    'Decreto y Fuente Oficial': s.decretoYFuente,
    'Período de Gestión': s.periodoGestion,
    'Cantidad Hallazgos': s.cantidadHallazgos,
    'Monto Imputado (RD$)': s.montoTotalImputadoRD,
    'Monto Descargado (RD$)': s.montoDescargadoRD,
    'Monto Persistente (RD$)': s.montoPersistenteRD
  }));
  const wsSeg = XLSX.utils.json_to_sheet(segRows);
  XLSX.utils.book_append_sheet(wb, wsSeg, 'Segregacion_Directores');

  // HOJA 4: CADENA DE CUSTODIA SHA-256 Y FUENTES VERIFICADAS
  const docRows = documents.map((d: any) => ({
    'Archivo Analizado': d.name,
    'Formato': d.type,
    'Tamaño (Bytes)': d.sizeBytes,
    'Huella Digital SHA-256': d.sha256,
    'Fecha Extracción': d.extractedAt
  }));
  const wsDocs = XLSX.utils.json_to_sheet(docRows);
  XLSX.utils.book_append_sheet(wb, wsDocs, 'Cadena_Custodia_SHA256');

  if (webSources && webSources.length > 0) {
    const wsSources = XLSX.utils.json_to_sheet(webSources.map((w: any) => ({
      'Fuente Oficial Consultada': w.title,
      'URL Verificada': w.uri
    })));
    XLSX.utils.book_append_sheet(wb, wsSources, 'Fuentes_Decretos_Web');
  }

  const fileName = `Matriz_Forense_CCRD_${(analysis.entidad || 'Auditoria').replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
