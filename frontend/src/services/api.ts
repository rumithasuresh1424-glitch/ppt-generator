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
