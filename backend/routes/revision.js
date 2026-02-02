const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');

const revisionsFile = path.join(__dirname, '../../uploads/revisions.json');

function loadRevisions() {
    if (fs.existsSync(revisionsFile)) {
        return JSON.parse(fs.readFileSync(revisionsFile, 'utf8'));
    }
    return [];
}

// Get all revisions
router.get('/', (req, res) => {
    try {
        const revisions = loadRevisions();
        res.json({ success: true, revisions });
    } catch (error) {
        res.status(500).json({ error: 'Failed to load revisions' });
    }
});

// Get specific revision
router.get('/:id', (req, res) => {
    try {
        const revisions = loadRevisions();
        const revision = revisions.find(r => r.id === req.params.id);
        
        if (!revision) {
            return res.status(404).json({ error: 'Revision not found' });
        }
        
        res.json({ success: true, revision });
    } catch (error) {
        res.status(500).json({ error: 'Failed to load revision' });
    }
});

// Delete revision
router.delete('/:id', (req, res) => {
    try {
        let revisions = loadRevisions();
        const revision = revisions.find(r => r.id === req.params.id);
        
        if (!revision) {
            return res.status(404).json({ error: 'Revision not found' });
        }
        
        // Delete file
        const filePath = path.join(__dirname, '../../uploads', revision.filename);
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
        }
        
        // Remove from revisions
        revisions = revisions.filter(r => r.id !== req.params.id);
        fs.writeFileSync(revisionsFile, JSON.stringify(revisions, null, 2));
        
        res.json({ success: true, message: 'Revision deleted' });
    } catch (error) {
        res.status(500).json({ error: 'Failed to delete revision' });
    }
});

module.exports = router;
