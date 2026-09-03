const https = require('https');
const http = require('http');
const fs = require('fs');
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
  res.sendFile(path.join(taskpaneDir, 'index.html'));
});

app.get('/taskpane.html', (req, res) => {
  res.sendFile(path.join(taskpaneDir, 'taskpane.html'));
});

const certDir = path.join(process.env.USERPROFILE, '.office-addin-dev-certs');
const keyPath = path.join(certDir, 'localhost.key');
const certPath = path.join(certDir, 'localhost.crt');

if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
  const options = {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath)
  };

  https.createServer(options, app).listen(3000, () => {
    console.log('[HTTPS Server] Running at https://localhost:3000');
    console.log('[Taskpane URL] https://localhost:3000/taskpane.html');
  });
} else {
  http.createServer(app).listen(3000, () => {
    console.log('[HTTP Server] Running at http://localhost:3000 (No dev certs found)');
  });
}
