const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const ExcelService = require('../services/excelService');

const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');
function loadRevisions() {
    if (fs.existsSync(revisionsFile)) return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    return [];
}

router.get('/issue-do/:revisionId', async (req, res) => {
    try {
        const rev = loadRevisions().find(r => r.id === req.params.revisionId);
        if (!rev) return res.status(404).json({ error: 'Revision not found' });
        const sourcePath = path.join(__dirname, '../../uploads', rev.filename);
        if (!fs.existsSync(sourcePath)) return res.status(404).json({ error: 'Source file not found' });
        const excelService = new ExcelService();
        const exportPath = await excelService.generateIssueDO(sourcePath, rev.id);
        const filename = `Issue_DO_${new Date().toISOString().split('T')[0]}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        const stream = fs.createReadStream(exportPath);
        stream.pipe(res);
        stream.on('end', () => { try { fs.unlinkSync(exportPath); } catch (e) {} });
    } catch (e) { console.error('Export error:', e); res.status(500).json({ error: 'Export failed', message: e.message }); }
});

router.get('/delivery-daily/:revisionId', async (req, res) => {
    try {
        const rev = loadRevisions().find(r => r.id === req.params.revisionId);
        if (!rev) return res.status(404).json({ error: 'Revision not found' });
        const sourcePath = path.join(__dirname, '../../uploads', rev.filename);
        if (!fs.existsSync(sourcePath)) return res.status(404).json({ error: 'Source file not found' });
        const excelService = new ExcelService();
        const exportPath = await excelService.generateDeliveryDaily(sourcePath, rev.id);
        const filename = `Delivery_Daily_${new Date().toISOString().split('T')[0]}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        const stream = fs.createReadStream(exportPath);
        stream.pipe(res);
        stream.on('end', () => { try { fs.unlinkSync(exportPath); } catch (e) {} });
    } catch (e) { console.error('Export error:', e); res.status(500).json({ error: 'Export failed', message: e.message }); }
});

router.get('/preview/:revisionId', async (req, res) => {
    try {
        const rev = loadRevisions().find(r => r.id === req.params.revisionId);
        if (!rev) return res.status(404).json({ error: 'Revision not found' });
        const sourcePath = path.join(__dirname, '../../uploads', rev.filename);
        const excelService = new ExcelService();
        const data = await excelService.parsePoData(sourcePath);
        res.json({ success: true, data });
    } catch (e) { res.status(500).json({ error: 'Preview failed', message: e.message }); }
});

module.exports = router;
