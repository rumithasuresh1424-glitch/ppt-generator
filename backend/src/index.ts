import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import AdmZip from 'adm-zip';
import path from 'path';
import fs from 'fs';

const app = express();
const PORT = process.env.PORT || 5000;

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, '..', 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Multer configuration for Excel files
const excelUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (_req, file, cb) => {
    const allowedTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(file.mimetype) || ext === '.xls' || ext === '.xlsx') {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only Excel files (.xls, .xlsx) are allowed.'));
    }
  },
});

// Multer configuration for ZIP files
const zipUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit for ZIP
  },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const mimeType = file.mimetype;
    if (ext === '.zip' || mimeType === 'application/zip' || mimeType === 'application/x-zip-compressed') {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only ZIP files are allowed.'));
    }
  },
});

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' });
});

// Excel upload endpoint
app.post('/api/excel/upload', excelUpload.single('file'), (req, res) => {
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

// Photos upload endpoint
app.post('/api/photos/upload', zipUpload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ success: false, error: 'No file selected' });
      return;
    }

    const allowedExtensions = ['.jpg', '.jpeg', '.png', '.webp'];
    const zip = new AdmZip(req.file.buffer);
    const entries = zip.getEntries();

    const imageNames: string[] = [];
    const seenNames = new Set<string>();

    for (const entry of entries) {
      if (entry.isDirectory) continue;

      const ext = path.extname(entry.entryName).toLowerCase();
      if (!allowedExtensions.includes(ext)) continue;

      const baseName = path.basename(entry.entryName, ext);
      if (!baseName) continue;

      // Check for duplicate filenames
      if (seenNames.has(baseName.toLowerCase())) {
        console.warn(`Duplicate filename detected: ${baseName}`);
        continue;
      }

      seenNames.add(baseName.toLowerCase());
      imageNames.push(baseName);
    }

    imageNames.sort();

    res.json({
      success: true,
      totalImages: imageNames.length,
      imageNames,
    });
  } catch (error) {
    console.error('Error extracting ZIP file:', error);
    res.status(500).json({ success: false, error: 'Unable to extract ZIP file or file is corrupted' });
  }
});

// Photo match endpoint
app.post('/api/photos/match', (req, res) => {
  try {
    const { excelData, matchColumn, imageNames } = req.body as {
      excelData: Record<string, unknown>[];
      matchColumn: string;
      imageNames: string[];
    };

    if (!excelData || !Array.isArray(excelData)) {
      res.status(400).json({ success: false, error: 'Invalid Excel data' });
      return;
    }

    if (!matchColumn) {
      res.status(400).json({ success: false, error: 'Match column is required' });
      return;
    }

    if (!imageNames || !Array.isArray(imageNames)) {
      res.status(400).json({ success: false, error: 'Image names are required' });
      return;
    }

    const imageSet = new Set(imageNames.map((name) => name.toLowerCase()));
    let matched = 0;
    let unmatched = 0;

    const data = excelData.map((row) => {
      const matchValue = row[matchColumn];
      const matchKey = matchValue != null ? String(matchValue).toLowerCase() : '';

      const hasPhoto = imageSet.has(matchKey);
      if (hasPhoto) {
        matched++;
      } else {
        unmatched++;
      }

      return {
        ...row,
        _photoMatched: hasPhoto,
        _photoKey: hasPhoto ? matchValue : null,
      };
    });

    res.json({
      success: true,
      matched,
      unmatched,
      data,
    });
  } catch (error) {
    console.error('Error matching photos:', error);
    res.status(500).json({ success: false, error: 'Unable to match photos' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
