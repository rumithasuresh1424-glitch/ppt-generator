import { useState, useCallback, useRef } from 'react';
import { Upload, Search, X, FileSpreadsheet } from 'lucide-react';
import { uploadExcel, ExcelData } from '../services/api';
import PreviewTable from '../components/PreviewTable';

export default function GeneratePPT() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = async (file: File | null) => {
    if (!file) return;

    const allowedTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];
    const ext = file.name.split('.').pop()?.toLowerCase();

    if (!allowedTypes.includes(file.type) && ext !== 'xls' && ext !== 'xlsx') {
      setError('Invalid file type. Please upload an Excel file (.xls or .xlsx)');
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      setError('File is too large. Maximum size is 10MB.');
      return;
    }

    setExcelFile(file);
    setError(null);
    setIsUploading(true);

    try {
      const data = await uploadExcel(file);
      if (data.success) {
        setExcelData(data);
        setError(null);
      } else {
        setError(data.error || 'Failed to parse Excel file');
        setExcelData(null);
      }
    } catch (err) {
      setError('Unable to parse Excel file. Please check the file format.');
      setExcelData(null);
    } finally {
      setIsUploading(false);
    }
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    handleFileChange(file);
  }, []);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    handleFileChange(file || null);
  };

  const clearFile = () => {
    setExcelFile(null);
    setExcelData(null);
    setError(null);
    setSearchQuery('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="container mx-auto px-4 py-12">
      <div className="max-w-4xl mx-auto space-y-8">
        <div className="text-center">
          <h1 className="text-3xl font-bold mb-2">Generate PowerPoint</h1>
          <p className="text-muted-foreground">
            Upload your Excel file to preview employee data
          </p>
        </div>

        <div className="space-y-6">
          <div className="border rounded-lg p-6">
            <h2 className="text-lg font-medium mb-4">Upload Excel</h2>
            {!excelFile ? (
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`flex flex-col items-center justify-center w-full h-40 border-2 border-dashed rounded-lg cursor-pointer transition-colors ${
                  isDragging
                    ? 'border-primary bg-primary/5'
                    : 'hover:bg-muted/50 border-muted-foreground/25'
                }`}
              >
                <Upload className="h-10 w-10 mb-3 text-muted-foreground" />
                <p className="text-sm text-muted-foreground mb-1">
                  Drag & drop your Excel file here
                </p>
                <p className="text-xs text-muted-foreground">or click to browse</p>
                <p className="text-xs text-muted-foreground mt-2">.xls, .xlsx (max 10MB)</p>
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  onChange={handleInputChange}
                />
              </div>
            ) : (
              <div className="flex items-center justify-between p-4 bg-muted rounded-lg">
                <div className="flex items-center gap-3">
                  <FileSpreadsheet className="h-8 w-8 text-primary" />
                  <div>
                    <p className="font-medium">{excelFile.name}</p>
                    <p className="text-sm text-muted-foreground">
                      {(excelFile.size / 1024).toFixed(1)} KB
                    </p>
                  </div>
                </div>
                <button
                  onClick={clearFile}
                  className="p-2 hover:bg-muted-foreground/10 rounded-lg"
                >
                  <X className="h-5 w-5" />
                </button>
              </div>
            )}

            {isUploading && (
              <div className="mt-4 text-center text-sm text-muted-foreground">
                Processing file...
              </div>
            )}
          </div>

          {error && (
            <div className="p-4 bg-destructive/10 border border-destructive/20 rounded-lg text-destructive text-sm">
              {error}
            </div>
          )}

          {excelData && (
            <>
              <div className="border rounded-lg p-6">
                <div className="grid grid-cols-3 gap-4">
                  <div>
                    <p className="text-sm text-muted-foreground">File Name</p>
                    <p className="font-medium truncate">{excelData.fileName}</p>
                  </div>
                  <div>
                    <p className="text-sm text-muted-foreground">Total Rows</p>
                    <p className="font-medium">{excelData.totalRows.toLocaleString()}</p>
                  </div>
                  <div>
                    <p className="text-sm text-muted-foreground">Total Columns</p>
                    <p className="font-medium">{excelData.totalColumns}</p>
                  </div>
                </div>
              </div>

              <div className="border rounded-lg p-6">
                <div className="flex items-center gap-4 mb-4">
                  <h2 className="text-lg font-medium">Preview</h2>
                  <div className="flex-1 relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                    <input
                      type="text"
                      placeholder="Search across all columns..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full h-10 pl-10 pr-4 border rounded-md bg-background text-sm"
                    />
                  </div>
                </div>

                <PreviewTable
                  columns={excelData.columns}
                  data={excelData.data}
                  searchQuery={searchQuery}
                />
              </div>

              <div className="border rounded-lg p-6">
                <h2 className="text-lg font-medium mb-4">Match Photos Using</h2>
                <select className="w-full h-10 px-3 py-2 border rounded-md bg-background text-sm">
                  <option value="">Select column...</option>
                  {excelData.columns.map((col) => (
                    <option key={col} value={col}>
                      {col}
                    </option>
                  ))}
                </select>
              </div>

              <button
                disabled
                className="w-full bg-primary text-primary-foreground px-6 py-3 rounded-lg font-medium opacity-50 cursor-not-allowed"
              >
                Generate PowerPoint
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
