const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const ExcelService = require('../services/excelService');

// Storage configuration
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = path.join(__dirname, '../../uploads');
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir, { recursive: true });
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        const revision = uuidv4();
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const ext = path.extname(file.originalname);
        cb(null, `${timestamp}_${revision}${ext}`);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xlsx', '.xls'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files (.xlsx, .xls) are allowed'));
        }
    },
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Revision storage (in production, use database)
const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');

function loadRevisions() {
    if (fs.existsSync(revisionsFile)) {
        return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    }
    return [];
}

function saveRevisions(revisions) {
    fs.writeFileSync(revisionsFile, JSON.stringify(revisions, null, 2));
}

// Upload endpoint
router.post('/', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const filePath = req.file.path;
        
        // Validate and parse Excel file
        const excelService = new ExcelService();
        const parseResult = await excelService.parsePoData(filePath);
        
        if (!parseResult.success) {
            // Remove invalid file
            fs.unlinkSync(filePath);
            return res.status(400).json({ error: parseResult.error });
        }

        // Create revision record
        const revision = {
            id: uuidv4(),
            filename: req.file.filename,
            originalName: req.file.originalname,
            uploadDate: new Date().toISOString(),
            size: req.file.size,
            recordCount: parseResult.recordCount,
            deliveryDate: parseResult.deliveryDate,
            customers: parseResult.customers
        };

        // Save revision
        const revisions = loadRevisions();
        revisions.unshift(revision);
        saveRevisions(revisions);

        res.json({
            success: true,
            message: 'File uploaded successfully',
            revision: revision
        });

    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: 'Upload failed', message: error.message });
    }
});

module.exports = router;
