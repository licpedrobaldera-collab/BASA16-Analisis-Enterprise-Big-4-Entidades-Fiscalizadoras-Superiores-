import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { extractText, getDocumentProxy } from 'unpdf';

export interface ParsedDocument {
  id: string;
  name: string;
  type: string;
  sizeBytes: number;
  sha256: string;
  content: string;
  extractedAt: string;
}

async function computeSha256(buffer: ArrayBuffer): Promise<string> {
  const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

export async function parseRealFile(file: File): Promise<ParsedDocument> {
  const buffer = await file.arrayBuffer();
  const sha256 = await computeSha256(buffer);
  const ext = file.name.split('.').pop()?.toLowerCase() || '';
  let content = '';

  if (ext === 'pdf') {
    const pdf = await getDocumentProxy(new Uint8Array(buffer));
    const { text } = await extractText(pdf, { mergePages: true });
    content = text;
  } else if (ext === 'docx') {
    const result = await mammoth.extractRawText({ arrayBuffer: buffer });
    content = result.value;
  } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheets: string[] = [];
    workbook.SheetNames.forEach((sheetName) => {
      const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      sheets.push(`--- HOJA: ${sheetName} ---\n${csv}`);
    });
    content = sheets.join('\n\n');
  } else {
    content = new TextDecoder('utf-8').decode(buffer);
  }

  if (!content.trim()) {
    throw new Error(`El archivo "${file.name}" no contiene texto legible extraíble.`);
  }

  return {
    id: crypto.randomUUID(),
    name: file.name,
    type: ext.toUpperCase(),
    sizeBytes: file.size,
    sha256,
    content: content.trim(),
    extractedAt: new Date().toISOString(),
  };
}
