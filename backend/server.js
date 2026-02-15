const express = require('express');
const cors = require('cors');
const path = require('path');
const uploadRoutes = require('./routes/upload');
const exportRoutes = require('./routes/export');
const revisionRoutes = require('./routes/revision');

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

app.use('/api/upload', uploadRoutes);
app.use('/api/export', exportRoutes);
app.use('/api/revisions', revisionRoutes);

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

app.get('/fluctuation', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/fluctuation.html'));
});

app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ error: 'Something went wrong!', message: err.message });
});

app.listen(PORT, () => {
    console.log(`FCOS Server running on http://localhost:${PORT}`);
});
