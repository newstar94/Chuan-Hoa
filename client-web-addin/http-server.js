const http = require('http');
const path = require('path');
const express = require('express');
const cors = require('cors');

const app = express();
app.use(cors());

// Static directories
const taskpaneDir = path.join(__dirname, 'src', 'taskpane');
const assetsDir = path.join(__dirname, 'assets');

app.use(express.static(taskpaneDir));
app.use('/assets', express.static(assetsDir));

// Fallback routes
app.get('/', (req, res) => {
  res.sendFile(path.join(taskpaneDir, 'taskpane.html'));
});

app.get('/taskpane.html', (req, res) => {
  res.sendFile(path.join(taskpaneDir, 'taskpane.html'));
});

app.get('/index.html', (req, res) => {
  res.sendFile(path.join(taskpaneDir, 'taskpane.html'));
});

const server = http.createServer(app);
server.listen(3000, () => {
  console.log('[HTTP Server] Running cleanly at http://localhost:3000');
  console.log('[Taskpane URL] http://localhost:3000/taskpane.html');
});
