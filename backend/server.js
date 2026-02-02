const express = require('express');
const cors = require('cors');
const path = require('path');
const uploadRoutes = require('./routes/upload');
const exportRoutes = require('./routes/export');
const revisionRoutes = require('./routes/revision');

const app = express();
const PORT = process.env.PORT || 8080;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

// Routes
app.use('/api/upload', uploadRoutes);
app.use('/api/export', exportRoutes);
app.use('/api/revisions', revisionRoutes);

// Serve frontend
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

// Error handling
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ error: 'Something went wrong!', message: err.message });
});

app.listen(PORT, () => {
    console.log(`FCOS Server running on http://localhost:${PORT}`);
});

module.exports = app;
