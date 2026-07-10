import { useMemo, useState } from 'react';
import { ChevronLeft, ChevronRight, User } from 'lucide-react';
import { MatchedRow } from '../services/api';

interface PreviewTableProps {
  columns: string[];
  data: MatchedRow[];
  searchQuery: string;
}

const ROWS_PER_PAGE = 10;

export default function PreviewTable({ columns, data, searchQuery }: PreviewTableProps) {
  const [currentPage, setCurrentPage] = useState(1);

  const filteredData = useMemo(() => {
    if (!searchQuery.trim()) return data;

    const query = searchQuery.toLowerCase();
    return data.filter((row) =>
      columns.some((col) => {
        const value = row[col];
        if (value == null) return false;
        return String(value).toLowerCase().includes(query);
      })
    );
  }, [data, columns, searchQuery]);

  const totalPages = Math.ceil(filteredData.length / ROWS_PER_PAGE);
  const startIndex = (currentPage - 1) * ROWS_PER_PAGE;
  const paginatedData = filteredData.slice(startIndex, startIndex + ROWS_PER_PAGE);

  const formatCellValue = (value: unknown): string => {
    if (value == null) return '';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  };

  if (filteredData.length === 0) {
    return (
      <div className="text-center py-8 text-muted-foreground">
        {searchQuery ? 'No matching results found' : 'No data available'}
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="overflow-x-auto border rounded-lg">
        <table className="w-full min-w-max">
          <thead className="bg-muted sticky top-0">
            <tr>
              <th className="px-4 py-3 text-left text-sm font-medium text-muted-foreground border-b w-16">
                #
              </th>
              <th className="px-4 py-3 text-left text-sm font-medium text-muted-foreground border-b w-[80px]">
                Photo
              </th>
              {columns.map((col) => (
                <th
                  key={col}
                  className="px-4 py-3 text-left text-sm font-medium text-muted-foreground border-b min-w-[150px]"
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paginatedData.map((row, idx) => (
              <tr key={startIndex + idx} className="hover:bg-muted/50">
                <td className="px-4 py-3 text-sm text-muted-foreground border-b">
                  {startIndex + idx + 1}
                </td>
                <td className="px-4 py-3 text-sm border-b">
                  {row._photoMatched ? (
                    <div className="w-[60px] h-[60px] bg-muted rounded-md flex items-center justify-center overflow-hidden">
                      <User className="h-6 w-6 text-muted-foreground" />
                    </div>
                  ) : (
                    <div className="w-[60px] h-[60px] bg-muted rounded-md flex items-center justify-center text-xs text-muted-foreground">
                      No Photo
                    </div>
                  )}
                </td>
                {columns.map((col) => (
                  <td key={col} className="px-4 py-3 text-sm border-b">
                    {formatCellValue(row[col])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {totalPages > 1 && (
        <div className="flex items-center justify-between text-sm text-muted-foreground">
          <span>
            Showing {startIndex + 1} to {Math.min(startIndex + ROWS_PER_PAGE, filteredData.length)} of{' '}
            {filteredData.length} rows
          </span>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              disabled={currentPage === 1}
              className="p-2 rounded-md hover:bg-muted disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
            <span>
              Page {currentPage} of {totalPages}
            </span>
            <button
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              className="p-2 rounded-md hover:bg-muted disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <ChevronRight className="h-4 w-4" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
