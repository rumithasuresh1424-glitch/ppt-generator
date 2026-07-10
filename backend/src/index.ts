import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import path from 'path';

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (_req, file, cb) => {
    const allowedTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.xls',
      '.xlsx',
    ];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(file.mimetype) || ext === '.xls' || ext === '.xlsx') {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only Excel files (.xls, .xlsx) are allowed.'));
    }
  },
});

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.post('/api/excel/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ success: false, error: 'No file selected' });
      return;
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const data: Record<string, unknown>[] = XLSX.utils.sheet_to_json(worksheet);

    if (data.length === 0) {
      res.status(400).json({ success: false, error: 'Excel file is empty' });
      return;
    }

    const columns = Object.keys(data[0]);

    res.json({
      success: true,
      fileName: req.file.originalname,
      totalRows: data.length,
      totalColumns: columns.length,
      columns,
      data,
    });
  } catch (error) {
    console.error('Error processing Excel file:', error);
    res.status(500).json({ success: false, error: 'Unable to parse Excel file' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
