const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = 3000;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(__dirname));

// Handle form submission
app.post('/submit', (req, res) => {
  const { name, email } = req.body;

  const filePath = path.join(__dirname, 'contact_data.xlsx');
  let workbook;
  let worksheet;

  // Check if file exists
  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contacts');
  }

  // Convert worksheet to JSON, add new data
  const data = XLSX.utils.sheet_to_json(worksheet);
  data.push({ Name: name, Email: email, Time: new Date().toLocaleString() });

  // Update worksheet
  const updatedSheet = XLSX.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;

  // Write back to file
  XLSX.writeFile(workbook, filePath);

  res.send('<h2 style="text-align:center;margin-top:50px;">Thanks! Your info was saved. <br><a href="/">Go Back</a></h2>');
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
