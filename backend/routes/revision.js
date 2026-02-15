const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');

const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');
function loadRevisions() {
    if (fs.existsSync(revisionsFile)) return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    return [];
}

router.get('/', (req, res) => {
    try { res.json({ success: true, revisions: loadRevisions() }); }
    catch (e) { res.status(500).json({ error: 'Failed to load revisions' }); }
});

router.get('/:id', (req, res) => {
    try {
        const rev = loadRevisions().find(r => r.id === req.params.id);
        if (!rev) return res.status(404).json({ error: 'Revision not found' });
        res.json({ success: true, revision: rev });
    } catch (e) { res.status(500).json({ error: 'Failed to load revision' }); }
});

router.delete('/:id', (req, res) => {
    try {
        let revisions = loadRevisions();
        const rev = revisions.find(r => r.id === req.params.id);
        if (!rev) return res.status(404).json({ error: 'Revision not found' });
        const filePath = path.join(__dirname, '../../uploads', rev.filename);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        revisions = revisions.filter(r => r.id !== req.params.id);
        fs.writeFileSync(revisionsFile, JSON.stringify(revisions, null, 2));
        res.json({ success: true, message: 'Revision deleted' });
    } catch (e) { res.status(500).json({ error: 'Failed to delete revision' }); }
});

module.exports = router;
