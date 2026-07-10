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
const path_1 = __importDefault(require("path"));
const fs_1 = __importDefault(require("fs"));
const app = (0, express_1.default)();
const PORT = process.env.PORT || 5000;
// Create uploads directory if it doesn't exist
const uploadsDir = path_1.default.join(__dirname, '..', 'uploads');
if (!fs_1.default.existsSync(uploadsDir)) {
    fs_1.default.mkdirSync(uploadsDir, { recursive: true });
}
app.use((0, cors_1.default)());
app.use(express_1.default.json({ limit: '50mb' }));
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
        const entries = zip.getEntries();
        const imageNames = [];
        const seenNames = new Set();
        for (const entry of entries) {
            if (entry.isDirectory)
                continue;
            const ext = path_1.default.extname(entry.entryName).toLowerCase();
            if (!allowedExtensions.includes(ext))
                continue;
            const baseName = path_1.default.basename(entry.entryName, ext);
            if (!baseName)
                continue;
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
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
//# sourceMappingURL=index.js.map