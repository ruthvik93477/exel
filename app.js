const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Multer configuration for file upload
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  }
});

const upload = multer({ storage: storage });

// Home route - render the upload form
app.get('/', (req, res) => {
  res.render('upload');
});

// Post route for file upload
app.post('/upload', upload.single('excel'), (req, res) => {
  const filePath = path.join(__dirname, req.file.path);

  // Read Excel file
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  
  // Convert Excel data to JSON
  const data = xlsx.utils.sheet_to_json(sheet);

  res.render('table', { data: data });
});

// Post route to save changes
app.post('/save', (req, res) => {
  // Assuming you want to update the Excel file here
  // You can access the updated data from req.body and write it back to the Excel file
  res.send('Changes saved successfully!');
});


app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
