const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const ExcelService = require('../services/excelService');

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = path.join(__dirname, '../../uploads');
        if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        cb(null, `${timestamp}_${file.originalname}`);
    }
});

const upload = multer({
    storage,
    limits: { fileSize: 10 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (['.xlsx', '.xls'].includes(ext)) cb(null, true);
        else cb(new Error('Only .xlsx and .xls files are allowed'));
    }
});

const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');
function loadRevisions() {
    if (fs.existsSync(revisionsFile)) return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    return [];
}
function saveRevisions(revisions) {
    fs.writeFileSync(revisionsFile, JSON.stringify(revisions, null, 2));
}

router.post('/', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        const excelService = new ExcelService();
        const summary = await excelService.parsePoData(req.file.path);
        const revision = {
            id: uuidv4(),
            filename: req.file.filename,
            originalName: req.file.originalname,
            uploadDate: new Date().toISOString(),
            size: req.file.size,
            summary: { totalRows: summary.totalRows, doNumbers: summary.doNumbers, customerCodes: summary.customerCodes, dateRange: summary.dateRange }
        };
        const revisions = loadRevisions();
        revisions.unshift(revision);
        saveRevisions(revisions);
        res.json({ success: true, message: 'File uploaded successfully', revision });
    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: 'Upload failed', message: error.message });
    }
});

module.exports = router;
