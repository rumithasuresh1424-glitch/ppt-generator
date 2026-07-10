import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import AdmZip from 'adm-zip';
import pptxgen from 'pptxgenjs';
import path from 'path';
import fs from 'fs';

const app = express();
const PORT = process.env.PORT || 5000;

// Create directories if they don't exist
const uploadsDir = path.join(__dirname, '..', 'uploads');
const photosDir = path.join(uploadsDir, 'photos');
const generatedDir = path.join(__dirname, '..', 'generated');

[uploadsDir, photosDir, generatedDir].forEach((dir) => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Serve static files from generated directory
app.use('/generated', express.static(generatedDir));

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

// Photos upload endpoint - extracts and saves images to uploads/photos/
app.post('/api/photos/upload', zipUpload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ success: false, error: 'No file selected' });
      return;
    }

    const allowedExtensions = ['.jpg', '.jpeg', '.png', '.webp'];
    const zip = new AdmZip(req.file.buffer);

    // Clear previous photos
    if (fs.existsSync(photosDir)) {
      fs.readdirSync(photosDir).forEach((file) => {
        fs.unlinkSync(path.join(photosDir, file));
      });
    }

    const imageNames: string[] = [];
    const imagePaths: Record<string, string> = {};
    const seenNames = new Set<string>();

    for (const entry of zip.getEntries()) {
      if (entry.isDirectory) continue;

      const ext = path.extname(entry.entryName).toLowerCase();
      if (!allowedExtensions.includes(ext)) continue;

      const baseName = path.basename(entry.entryName, ext);
      if (!baseName) continue;

      // Check for duplicate filenames
      const key = baseName.toLowerCase();
      if (seenNames.has(key)) {
        console.warn(`Duplicate filename detected: ${baseName}`);
        continue;
      }

      seenNames.add(key);
      imageNames.push(baseName);

      // Save image to uploads/photos/
      const imageFileName = `${baseName}${ext}`;
      const imagePath = path.join(photosDir, imageFileName);
      fs.writeFileSync(imagePath, entry.getData());
      imagePaths[key] = imagePath;
    }

    imageNames.sort();

    res.json({
      success: true,
      totalImages: imageNames.length,
      imageNames,
      imagePaths,
    });
  } catch (error) {
    console.error('Error extracting ZIP file:', error);
    res.status(500).json({ success: false, error: 'Unable to extract ZIP file or file is corrupted' });
  }
});

// Photo match endpoint - now accepts imagePaths and returns them
app.post('/api/photos/match', (req, res) => {
  try {
    const { excelData, matchColumn, imageNames, imagePaths } = req.body as {
      excelData: Record<string, unknown>[];
      matchColumn: string;
      imageNames: string[];
      imagePaths: Record<string, string>;
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
        _photoPath: hasPhoto && imagePaths ? imagePaths[matchKey] : null,
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

// PowerPoint generation endpoint
app.post('/api/ppt/generate', async (req, res) => {
  try {
    const { excelData, columns } = req.body as {
      excelData: Record<string, unknown>[];
      columns: string[];
    };

    if (!excelData || !Array.isArray(excelData) || excelData.length === 0) {
      res.status(400).json({ success: false, error: 'No data to generate PowerPoint' });
      return;
    }

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'HR PowerPoint Generator';
    pptx.title = 'Employee Presentation';

    // Dark blue color for header
    const headerColor = '1a365d';
    const labelColor = '4a5568';
    const valueColor = '2d3748';
    const footerColor = '718096';

    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };

      // Header
      slide.addText('Employee Details', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.6,
        fontSize: 28,
        bold: true,
        color: headerColor,
        align: 'center',
      });

      // Left side - Photo area (2.2 x 2.8 inches)
      const photoX = 0.5;
      const photoY = 1.2;
      const photoW = 2.2;
      const photoH = 2.8;

      const hasPhoto = row._photoMatched === true;
      const photoPath = row._photoPath as string | null;

      if (hasPhoto && photoPath && fs.existsSync(photoPath)) {
        // Add actual photo using slide.addImage() with file path
        try {
          slide.addImage({
            x: photoX,
            y: photoY,
            w: photoW,
            h: photoH,
            path: photoPath,
          });
        } catch (imgError) {
          // Fallback to placeholder if image fails to load
          console.error(`Failed to load image: ${photoPath}`, imgError);
          slide.addShape(pptx.ShapeType.rect, {
            x: photoX,
            y: photoY,
            w: photoW,
            h: photoH,
            fill: { color: 'f7fafc' },
            line: { color: 'e2e8f0', width: 1, dashType: 'dash' },
          });
          slide.addText('No Photo Available', {
            x: photoX,
            y: photoY + photoH / 2 - 0.3,
            w: photoW,
            h: 0.6,
            fontSize: 12,
            color: 'a0aec0',
            align: 'center',
          });
        }
      } else {
        // No photo available box
        slide.addShape(pptx.ShapeType.rect, {
          x: photoX,
          y: photoY,
          w: photoW,
          h: photoH,
          fill: { color: 'f7fafc' },
          line: { color: 'e2e8f0', width: 1, dashType: 'dash' },
        });
        slide.addText('No Photo Available', {
          x: photoX,
          y: photoY + photoH / 2 - 0.3,
          w: photoW,
          h: 0.6,
          fontSize: 12,
          color: 'a0aec0',
          align: 'center',
        });
      }

      // Right side - Employee details
      const rightX = 3.0;
      const rightW = 6.5;
      let currentY = 1.3;
      const lineHeight = 0.45;

      for (const col of columns) {
        const value = row[col];
        if (value === undefined || value === null) continue;

        const displayValue = typeof value === 'object' ? JSON.stringify(value) : String(value);

        // Label
        slide.addText(col, {
          x: rightX,
          y: currentY,
          w: rightW,
          h: 0.25,
          fontSize: 11,
          color: labelColor,
          bold: false,
        });

        // Value
        slide.addText(displayValue, {
          x: rightX,
          y: currentY + 0.22,
          w: rightW,
          h: 0.25,
          fontSize: 14,
          color: valueColor,
          bold: false,
        });

        currentY += lineHeight + 0.2;
      }

      // Footer
      slide.addText('Generated by HR PowerPoint Generator', {
        x: 0.5,
        y: 5.2,
        w: 9,
        h: 0.3,
        fontSize: 10,
        color: footerColor,
        align: 'right',
      });

      // Slide number
      slide.addText(`Slide ${i + 1} of ${excelData.length}`, {
        x: 0.5,
        y: 5.2,
        w: 2,
        h: 0.3,
        fontSize: 10,
        color: footerColor,
        align: 'left',
      });
    }

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const fileName = `Employee_Presentation_${timestamp}.pptx`;
    const filePath = path.join(generatedDir, fileName);

    // Write the file
    await pptx.writeFile({ fileName: filePath });

    res.json({
      success: true,
      fileName,
      downloadUrl: `/generated/${fileName}`,
      totalSlides: excelData.length,
    });
  } catch (error) {
    console.error('Error generating PowerPoint:', error);
    res.status(500).json({ success: false, error: 'Failed to generate PowerPoint' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
