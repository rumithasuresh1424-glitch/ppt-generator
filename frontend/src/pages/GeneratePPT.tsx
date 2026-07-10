import { useState } from 'react';
import { Upload } from 'lucide-react';

export default function GeneratePPT() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [zipFile, setZipFile] = useState<File | null>(null);

  return (
    <div className="container mx-auto px-4 py-12">
      <div className="max-w-2xl mx-auto space-y-8">
        <div className="text-center">
          <h1 className="text-3xl font-bold mb-2">Generate PowerPoint</h1>
          <p className="text-muted-foreground">
            Upload your Excel file and ZIP of photos to generate employee presentations
          </p>
        </div>

        <div className="space-y-6">
          <div className="border rounded-lg p-6">
            <h2 className="text-lg font-medium mb-4">Upload Excel</h2>
            <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer hover:bg-muted/50 transition-colors">
              <div className="flex flex-col items-center justify-center pt-5 pb-6">
                <Upload className="h-8 w-8 mb-2 text-muted-foreground" />
                <p className="text-sm text-muted-foreground">
                  {excelFile ? excelFile.name : 'Choose File'}
                </p>
              </div>
              <input
                type="file"
                className="hidden"
                accept=".xlsx,.xls"
                onChange={(e) => setExcelFile(e.target.files?.[0] || null)}
              />
            </label>
          </div>

          <div className="border rounded-lg p-6">
            <h2 className="text-lg font-medium mb-4">Upload ZIP</h2>
            <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer hover:bg-muted/50 transition-colors">
              <div className="flex flex-col items-center justify-center pt-5 pb-6">
                <Upload className="h-8 w-8 mb-2 text-muted-foreground" />
                <p className="text-sm text-muted-foreground">
                  {zipFile ? zipFile.name : 'Choose ZIP'}
                </p>
              </div>
              <input
                type="file"
                className="hidden"
                accept=".zip"
                onChange={(e) => setZipFile(e.target.files?.[0] || null)}
              />
            </label>
          </div>

          <div className="border rounded-lg p-6">
            <h2 className="text-lg font-medium mb-4">Match Photos Using</h2>
            <select className="w-full h-10 px-3 py-2 border rounded-md bg-background text-sm">
              <option value="">Select column...</option>
            </select>
          </div>

          <button
            disabled
            className="w-full bg-primary text-primary-foreground px-6 py-3 rounded-lg font-medium opacity-50 cursor-not-allowed"
          >
            Generate PowerPoint
          </button>
        </div>
      </div>
    </div>
  );
}
