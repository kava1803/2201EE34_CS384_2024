const express = require('express');
const multer = require('multer');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const app = express();
const upload = multer({ dest: 'uploads/' }); // Directory for uploaded files
app.use(cors())
app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const pythonProcess = spawn('python', [
    path.join(__dirname, 'process_file.py'), // Your Python script path
    req.file.path, // Path to the uploaded file
  ]);

  pythonProcess.on('close', (code) => {
    if (code !== 0) {
      return res.status(500).send(`Python script exited with code ${code}`);
    }

    const outputFilePath = path.join(__dirname, '/output/output-file.xlsx');

    fs.access(outputFilePath, fs.constants.F_OK, (err) => {
      if (err) {
        return res.status(500).send('Output file not found.');
      }
      res.download(outputFilePath, 'output-file.xlsx', (downloadErr) => {
        if (downloadErr) {
          console.error('Error sending file:', downloadErr);
          res.status(500).send('Error sending file.');
        }
      });
    });
  });
  console.log("here")

  pythonProcess.stderr.on('data', (data) => {
    console.error(`Python script error: ${data}`);
  });
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
