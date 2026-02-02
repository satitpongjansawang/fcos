const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const ExcelService = require('../services/excelService');

const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');

function loadRevisions() {
    if (fs.existsSync(revisionsFile)) {
        return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    }
    return [];
}

// Export as Issue D/O
router.get('/issue-do/:revisionId', async (req, res) => {
    try {
        const revisions = loadRevisions();
        const revision = revisions.find(r => r.id === req.params.revisionId);
        
        if (!revision) {
            return res.status(404).json({ error: 'Revision not found' });
        }

        const sourcePath = path.join(__dirname, '../../uploads', revision.filename);
        
        if (!fs.existsSync(sourcePath)) {
            return res.status(404).json({ error: 'Source file not found' });
        }

        const excelService = new ExcelService();
        const exportPath = await excelService.generateIssueDO(sourcePath, revision.id);

        const filename = `Issue_DO_${new Date().toISOString().split('T')[0]}.xlsx`;
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        
        const fileStream = fs.createReadStream(exportPath);
        fileStream.pipe(res);
        
        // Clean up after sending
        fileStream.on('end', () => {
            fs.unlinkSync(exportPath);
        });

    } catch (error) {
        console.error('Export Issue D/O error:', error);
        res.status(500).json({ error: 'Export failed', message: error.message });
    }
});

// Export as Delivery Daily Report
router.get('/delivery-daily/:revisionId', async (req, res) => {
    try {
        const revisions = loadRevisions();
        const revision = revisions.find(r => r.id === req.params.revisionId);
        
        if (!revision) {
            return res.status(404).json({ error: 'Revision not found' });
        }

        const sourcePath = path.join(__dirname, '../../uploads', revision.filename);
        
        if (!fs.existsSync(sourcePath)) {
            return res.status(404).json({ error: 'Source file not found' });
        }

        const excelService = new ExcelService();
        const exportPath = await excelService.generateDeliveryDaily(sourcePath, revision.id);

        const filename = `Delivery_Daily_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        
        const fileStream = fs.createReadStream(exportPath);
        fileStream.pipe(res);
        
        // Clean up after sending
        fileStream.on('end', () => {
            fs.unlinkSync(exportPath);
        });

    } catch (error) {
        console.error('Export Delivery Daily error:', error);
        res.status(500).json({ error: 'Export failed', message: error.message });
    }
});

// Preview data (for checking before export)
router.get('/preview/:revisionId', async (req, res) => {
    try {
        const revisions = loadRevisions();
        const revision = revisions.find(r => r.id === req.params.revisionId);
        
        if (!revision) {
            return res.status(404).json({ error: 'Revision not found' });
        }

        const sourcePath = path.join(__dirname, '../../uploads', revision.filename);
        
        if (!fs.existsSync(sourcePath)) {
            return res.status(404).json({ error: 'Source file not found' });
        }

        const excelService = new ExcelService();
        const previewData = await excelService.getPreviewData(sourcePath);

        res.json({ success: true, data: previewData });

    } catch (error) {
        console.error('Preview error:', error);
        res.status(500).json({ error: 'Preview failed', message: error.message });
    }
});

module.exports = router;
