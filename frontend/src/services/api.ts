import axios from 'axios';

export interface ExcelData {
  success: boolean;
  fileName: string;
  totalRows: number;
  totalColumns: number;
  columns: string[];
  data: Record<string, unknown>[];
  error?: string;
}

export interface PhotoUploadResponse {
  success: boolean;
  totalImages: number;
  imageNames: string[];
  error?: string;
}

export interface MatchResponse {
  success: boolean;
  matched: number;
  unmatched: number;
  data: MatchedRow[];
  error?: string;
}

export interface MatchedRow extends Record<string, unknown> {
  _photoMatched: boolean;
  _photoKey: string | null;
}

export interface GenerateResponse {
  success: boolean;
  fileName: string;
  downloadUrl: string;
  totalSlides: number;
  error?: string;
}

export async function uploadExcel(file: File): Promise<ExcelData> {
  const formData = new FormData();
  formData.append('file', file);

  const response = await axios.post<ExcelData>('/api/excel/upload', formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  });

  return response.data;
}

export async function uploadPhotos(file: File): Promise<PhotoUploadResponse> {
  const formData = new FormData();
  formData.append('file', file);

  const response = await axios.post<PhotoUploadResponse>('/api/photos/upload', formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  });

  return response.data;
}

export async function matchPhotos(
  excelData: Record<string, unknown>[],
  matchColumn: string,
  imageNames: string[]
): Promise<MatchResponse> {
  const response = await axios.post<MatchResponse>('/api/photos/match', {
    excelData,
    matchColumn,
    imageNames,
  });

  return response.data;
}

export async function generatePPT(
  excelData: MatchedRow[],
  columns: string[]
): Promise<GenerateResponse> {
  const response = await axios.post<GenerateResponse>('/api/ppt/generate', {
    excelData,
    columns,
  });

  return response.data;
}
