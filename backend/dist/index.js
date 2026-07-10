"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const multer_1 = __importDefault(require("multer"));
const XLSX = __importStar(require("xlsx"));
const adm_zip_1 = __importDefault(require("adm-zip"));
const pptxgenjs_1 = __importDefault(require("pptxgenjs"));
const path_1 = __importDefault(require("path"));
const fs_1 = __importDefault(require("fs"));
const app = (0, express_1.default)();
const PORT = process.env.PORT || 5000;
// Create directories if they don't exist
const uploadsDir = path_1.default.join(__dirname, '..', 'uploads');
const generatedDir = path_1.default.join(__dirname, '..', 'generated');
if (!fs_1.default.existsSync(uploadsDir)) {
    fs_1.default.mkdirSync(uploadsDir, { recursive: true });
}
if (!fs_1.default.existsSync(generatedDir)) {
    fs_1.default.mkdirSync(generatedDir, { recursive: true });
}
app.use((0, cors_1.default)());
app.use(express_1.default.json({ limit: '50mb' }));
// Serve static files from generated directory
app.use('/generated', express_1.default.static(generatedDir));
// Multer configuration for Excel files
const excelUpload = (0, multer_1.default)({
    storage: multer_1.default.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024, // 10MB limit
    },
    fileFilter: (_req, file, cb) => {
        const allowedTypes = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];
        const ext = path_1.default.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(file.mimetype) || ext === '.xls' || ext === '.xlsx') {
            cb(null, true);
        }
        else {
            cb(new Error('Invalid file type. Only Excel files (.xls, .xlsx) are allowed.'));
        }
    },
});
// Multer configuration for ZIP files
const zipUpload = (0, multer_1.default)({
    storage: multer_1.default.memoryStorage(),
    limits: {
        fileSize: 50 * 1024 * 1024, // 50MB limit for ZIP
    },
    fileFilter: (_req, file, cb) => {
        const ext = path_1.default.extname(file.originalname).toLowerCase();
        const mimeType = file.mimetype;
        if (ext === '.zip' || mimeType === 'application/zip' || mimeType === 'application/x-zip-compressed') {
            cb(null, true);
        }
        else {
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
        const data = XLSX.utils.sheet_to_json(worksheet);
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
    }
    catch (error) {
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
        const zip = new adm_zip_1.default(req.file.buffer);
        const imageNames = [];
        const seenNames = new Set();
        const imageBuffers = {};
        for (const entry of zip.getEntries()) {
            if (entry.isDirectory)
                continue;
            const ext = path_1.default.extname(entry.entryName).toLowerCase();
            if (!allowedExtensions.includes(ext))
                continue;
            const baseName = path_1.default.basename(entry.entryName, ext);
            if (!baseName)
                continue;
            // Check for duplicate filenames
            const key = baseName.toLowerCase();
            if (seenNames.has(key)) {
                console.warn(`Duplicate filename detected: ${baseName}`);
                continue;
            }
            seenNames.add(key);
            imageNames.push(baseName);
            imageBuffers[key] = { buffer: entry.getData(), ext };
        }
        imageNames.sort();
        res.json({
            success: true,
            totalImages: imageNames.length,
            imageNames,
        });
    }
    catch (error) {
        console.error('Error extracting ZIP file:', error);
        res.status(500).json({ success: false, error: 'Unable to extract ZIP file or file is corrupted' });
    }
});
// Photo match endpoint
app.post('/api/photos/match', (req, res) => {
    try {
        const { excelData, matchColumn, imageNames } = req.body;
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
            }
            else {
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
    }
    catch (error) {
        console.error('Error matching photos:', error);
        res.status(500).json({ success: false, error: 'Unable to match photos' });
    }
});
// PowerPoint generation endpoint
app.post('/api/ppt/generate', async (req, res) => {
    try {
        const { excelData, columns } = req.body;
        if (!excelData || !Array.isArray(excelData) || excelData.length === 0) {
            res.status(400).json({ success: false, error: 'No data to generate PowerPoint' });
            return;
        }
        const pptx = new pptxgenjs_1.default();
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
            // Left side - Photo placeholder (2.2 x 2.8 inches)
            const hasPhoto = row._photoMatched === true;
            const photoX = 0.5;
            const photoY = 1.2;
            const photoW = 2.2;
            const photoH = 2.8;
            if (hasPhoto) {
                // Add a placeholder box for the photo
                slide.addShape(pptx.ShapeType.rect, {
                    x: photoX,
                    y: photoY,
                    w: photoW,
                    h: photoH,
                    fill: { color: 'e2e8f0' },
                    line: { color: 'cbd5e0', width: 1 },
                });
                slide.addText('[Photo]', {
                    x: photoX,
                    y: photoY + photoH / 2 - 0.3,
                    w: photoW,
                    h: 0.6,
                    fontSize: 14,
                    color: 'a0aec0',
                    align: 'center',
                });
            }
            else {
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
                if (value === undefined || value === null)
                    continue;
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
        const filePath = path_1.default.join(generatedDir, fileName);
        // Write the file
        await pptx.writeFile({ fileName: filePath });
        res.json({
            success: true,
            fileName,
            downloadUrl: `/generated/${fileName}`,
            totalSlides: excelData.length,
        });
    }
    catch (error) {
        console.error('Error generating PowerPoint:', error);
        res.status(500).json({ success: false, error: 'Failed to generate PowerPoint' });
    }
});
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
//# sourceMappingURL=index.js.map