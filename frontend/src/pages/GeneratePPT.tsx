import { useState, useCallback, useRef, useEffect } from 'react';
import { Upload, Search, X, FileSpreadsheet, Archive, CheckCircle, XCircle, Loader2, Download } from 'lucide-react';
import { uploadExcel, uploadPhotos, matchPhotos, generatePPT, ExcelData, MatchedRow, PhotoUploadResponse } from '../services/api';
import PreviewTable from '../components/PreviewTable';

export default function GeneratePPT() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [zipFile, setZipFile] = useState<File | null>(null);
  const [photoData, setPhotoData] = useState<PhotoUploadResponse | null>(null);
  const [matchedData, setMatchedData] = useState<MatchedRow[]>([]);
  const [isDraggingExcel, setIsDraggingExcel] = useState(false);
  const [isDraggingZip, setIsDraggingZip] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [isMatching, setIsMatching] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [generateSuccess, setGenerateSuccess] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedColumn, setSelectedColumn] = useState<string>('');
  const excelInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);

  // Auto-match when both ZIP is uploaded and column is selected
  useEffect(() => {
    const performMatch = async () => {
      if (excelData && zipFile && photoData && selectedColumn) {
        setIsMatching(true);
        setGenerateSuccess(false);
        setDownloadUrl(null);
        try {
          const result = await matchPhotos(excelData.data, selectedColumn, photoData.imageNames);
          if (result.success) {
            setMatchedData(result.data);
            setError(null);
          } else {
            setError(result.error || 'Failed to match photos');
          }
        } catch (err) {
          setError('Unable to match photos');
        } finally {
          setIsMatching(false);
        }
      }
    };

    performMatch();
  }, [excelData, zipFile, photoData, selectedColumn]);

  const handleExcelChange = async (file: File | null) => {
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

    // Reset state
    setZipFile(null);
    setPhotoData(null);
    setMatchedData([]);
    setSelectedColumn('');
    setGenerateSuccess(false);
    setDownloadUrl(null);

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

  const handleZipChange = async (file: File | null) => {
    if (!file) return;

    const ext = file.name.split('.').pop()?.toLowerCase();

    if (ext !== 'zip') {
      setError('Invalid file type. Please upload a ZIP file.');
      return;
    }

    if (file.size > 50 * 1024 * 1024) {
      setError('File is too large. Maximum size is 50MB.');
      return;
    }

    setZipFile(file);
    setError(null);
    setIsUploading(true);
    setGenerateSuccess(false);
    setDownloadUrl(null);

    try {
      const data = await uploadPhotos(file);
      if (data.success) {
        setPhotoData(data);
        setError(null);
      } else {
        setError(data.error || 'Failed to extract ZIP file');
        setPhotoData(null);
      }
    } catch (err) {
      setError('Unable to extract ZIP file. Please check if the file is corrupted.');
      setPhotoData(null);
    } finally {
      setIsUploading(false);
    }
  };

  const handleExcelDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingExcel(true);
  }, []);

  const handleExcelDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingExcel(false);
  }, []);

  const handleExcelDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingExcel(false);
    const file = e.dataTransfer.files[0];
    handleExcelChange(file);
  }, []);

  const handleZipDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingZip(true);
  }, []);

  const handleZipDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingZip(false);
  }, []);

  const handleZipDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingZip(false);
    const file = e.dataTransfer.files[0];
    handleZipChange(file);
  }, []);

  const clearExcel = () => {
    setExcelFile(null);
    setExcelData(null);
    setZipFile(null);
    setPhotoData(null);
    setMatchedData([]);
    setError(null);
    setSearchQuery('');
    setSelectedColumn('');
    setGenerateSuccess(false);
    setDownloadUrl(null);
    if (excelInputRef.current) {
      excelInputRef.current.value = '';
    }
  };

  const clearZip = () => {
    setZipFile(null);
    setPhotoData(null);
    setMatchedData([]);
    setGenerateSuccess(false);
    setDownloadUrl(null);
    if (zipInputRef.current) {
      zipInputRef.current.value = '';
    }
  };

  const handleGenerate = async () => {
    if (!excelData || !matchedData.length || !selectedColumn) return;

    setIsGenerating(true);
    setError(null);

    try {
      const columns = excelData.columns;
      const result = await generatePPT(matchedData, columns);

      if (result.success) {
        setGenerateSuccess(true);
        // Use full backend URL for download (backend runs on port 5000)
        const backendUrl = 'http://localhost:5000';
        setDownloadUrl(`${backendUrl}${result.downloadUrl}`);
        setError(null);
      } else {
        setError(result.error || 'Failed to generate PowerPoint');
        setGenerateSuccess(false);
      }
    } catch (err) {
      setError('Unable to generate PowerPoint. Please try again.');
      setGenerateSuccess(false);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDownload = async () => {
    if (!downloadUrl) return;

    try {
      const response = await fetch(downloadUrl);
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = downloadUrl.split('/').pop() || 'Employee_Presentation.pptx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Download failed:', err);
    }
  };

  const canGenerate = excelData && zipFile && selectedColumn && matchedData.length > 0;
  const tableData: MatchedRow[] = matchedData.length > 0 ? matchedData : (excelData?.data as MatchedRow[]) || [];

  const matchedCount = matchedData.filter((row) => row._photoMatched).length;
  const unmatchedCount = matchedData.filter((row) => !row._photoMatched).length;

  return (
    <div className="container mx-auto px-4 py-12">
      <div className="max-w-4xl mx-auto space-y-8">
        <div className="text-center">
          <h1 className="text-3xl font-bold mb-2">Generate PowerPoint</h1>
          <p className="text-muted-foreground">
            Upload your Excel file and ZIP of photos to generate employee presentations
          </p>
        </div>

        <div className="space-y-6">
          {/* Upload Excel Section */}
          <div className="border rounded-lg p-6">
            <h2 className="text-lg font-medium mb-4">Upload Excel</h2>
            {!excelFile ? (
              <div
                onDragOver={handleExcelDragOver}
                onDragLeave={handleExcelDragLeave}
                onDrop={handleExcelDrop}
                onClick={() => excelInputRef.current?.click()}
                className={`flex flex-col items-center justify-center w-full h-40 border-2 border-dashed rounded-lg cursor-pointer transition-colors ${
                  isDraggingExcel
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
                  ref={excelInputRef}
                  type="file"
                  className="hidden"
                  accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  onChange={(e) => handleExcelChange(e.target.files?.[0] || null)}
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
                  onClick={clearExcel}
                  className="p-2 hover:bg-muted-foreground/10 rounded-lg"
                >
                  <X className="h-5 w-5" />
                </button>
              </div>
            )}

            {isUploading && !zipFile && (
              <div className="mt-4 text-center text-sm text-muted-foreground">
                Processing file...
              </div>
            )}
          </div>

          {/* Error Display */}
          {error && (
            <div className="p-4 bg-destructive/10 border border-destructive/20 rounded-lg text-destructive text-sm">
              {error}
            </div>
          )}

          {/* File Info & Preview - shown only after Excel upload */}
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

              {/* Upload ZIP Section - Enabled after Excel upload */}
              <div className={`border rounded-lg p-6 ${!excelData ? 'opacity-60' : ''}`}>
                <h2 className="text-lg font-medium mb-4">Upload ZIP</h2>
                {!zipFile ? (
                  <div
                    onDragOver={handleZipDragOver}
                    onDragLeave={handleZipDragLeave}
                    onDrop={handleZipDrop}
                    onClick={() => excelData && zipInputRef.current?.click()}
                    className={`flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer transition-colors ${
                      isDraggingZip
                        ? 'border-primary bg-primary/5'
                        : excelData
                          ? 'hover:bg-muted/50 border-muted-foreground/25 cursor-pointer'
                          : 'border-muted-foreground/25 cursor-not-allowed'
                    }`}
                  >
                    <Archive className="h-8 w-8 mb-2 text-muted-foreground" />
                    <p className="text-sm text-muted-foreground">Choose ZIP</p>
                    <p className="text-xs text-muted-foreground mt-2">.zip (max 50MB)</p>
                    <input
                      ref={zipInputRef}
                      type="file"
                      className="hidden"
                      accept=".zip"
                      onChange={(e) => handleZipChange(e.target.files?.[0] || null)}
                      disabled={!excelData}
                    />
                  </div>
                ) : (
                  <div className="flex items-center justify-between p-4 bg-muted rounded-lg">
                    <div className="flex items-center gap-3">
                      <Archive className="h-8 w-8 text-primary" />
                      <div>
                        <p className="font-medium">{zipFile.name}</p>
                        <p className="text-sm text-muted-foreground">
                          {(zipFile.size / 1024).toFixed(1)} KB • {photoData?.totalImages} images
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={clearZip}
                      className="p-2 hover:bg-muted-foreground/10 rounded-lg"
                    >
                      <X className="h-5 w-5" />
                    </button>
                  </div>
                )}

                {isUploading && zipFile && (
                  <div className="mt-4 text-center text-sm text-muted-foreground">
                    Extracting images...
                  </div>
                )}
              </div>

              {/* Match Photos Using Section */}
              <div className="border rounded-lg p-6">
                <h2 className="text-lg font-medium mb-4">Match Photos Using</h2>
                <select
                  value={selectedColumn}
                  onChange={(e) => setSelectedColumn(e.target.value)}
                  disabled={!excelData || !zipFile}
                  className="w-full h-10 px-3 py-2 border rounded-md bg-background text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <option value="">Select column...</option>
                  {excelData?.columns.map((col) => (
                    <option key={col} value={col}>
                      {col}
                    </option>
                  ))}
                </select>
                {excelData && (
                  <p className="text-xs text-muted-foreground mt-2">
                    {excelData.columns.length > 0
                      ? `${excelData.columns.length} columns available`
                      : 'Upload an Excel file to see columns'}
                  </p>
                )}
              </div>

              {/* Match Summary - shown after matching */}
              {matchedData.length > 0 && (
                <div className="border rounded-lg p-6">
                  <h2 className="text-lg font-medium mb-4">Summary</h2>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="flex items-center gap-3 p-4 bg-green-50 rounded-lg">
                      <CheckCircle className="h-6 w-6 text-green-600" />
                      <div>
                        <p className="text-sm text-green-600">Matched Photos</p>
                        <p className="text-2xl font-bold text-green-700">{matchedCount}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-3 p-4 bg-orange-50 rounded-lg">
                      <XCircle className="h-6 w-6 text-orange-600" />
                      <div>
                        <p className="text-sm text-orange-600">Unmatched</p>
                        <p className="text-2xl font-bold text-orange-700">{unmatchedCount}</p>
                      </div>
                    </div>
                  </div>
                  {isMatching && (
                    <p className="text-sm text-muted-foreground text-center mt-4">
                      Matching photos...
                    </p>
                  )}
                </div>
              )}

              {/* Preview Table */}
              <div className="border rounded-lg p-6">
                <div className="flex items-center gap-4 mb-4">
                  <h2 className="text-lg font-medium">Preview</h2>
                  {matchedData.length > 0 && (
                    <span className="text-sm text-muted-foreground">
                      (Photos matched: {matchedCount}/{matchedData.length})
                    </span>
                  )}
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
                  data={tableData}
                  searchQuery={searchQuery}
                />
              </div>
            </>
          )}

          {/* Generate PowerPoint Button */}
          {!generateSuccess ? (
            <>
              <button
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  handleGenerate();
                }}
                disabled={!canGenerate || isGenerating}
                className={`w-full px-6 py-3 rounded-lg font-medium transition-colors flex items-center justify-center gap-2 ${
                  canGenerate && !isGenerating
                    ? 'bg-primary text-primary-foreground hover:bg-primary/90 cursor-pointer'
                    : 'bg-muted text-muted-foreground opacity-50 cursor-not-allowed'
                }`}
              >
                {isGenerating ? (
                  <>
                    <Loader2 className="h-5 w-5 animate-spin" />
                    Generating PowerPoint...
                  </>
                ) : (
                  'Generate PowerPoint'
                )}
              </button>
              {!canGenerate && excelData && (
                <p className="text-xs text-muted-foreground text-center -mt-4">
                  {!zipFile
                    ? 'Upload a ZIP file with photos'
                    : !selectedColumn
                      ? 'Select a column to match photos'
                      : 'Processing...'}
                </p>
              )}
            </>
          ) : (
            <div className="space-y-4">
              <div className="flex items-center justify-center gap-2 p-4 bg-green-50 rounded-lg border border-green-200">
                <CheckCircle className="h-6 w-6 text-green-600" />
                <span className="text-green-700 font-medium">PowerPoint Generated Successfully!</span>
              </div>
              <button
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  handleDownload();
                }}
                className="w-full px-6 py-3 rounded-lg font-medium transition-colors bg-green-600 text-white hover:bg-green-700 cursor-pointer flex items-center justify-center gap-2"
              >
                <Download className="h-5 w-5" />
                Download PowerPoint
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
