const express = require('express');
const app = express();
const port = 5000;

app.use(express.json());  // To parse JSON bodies

// Utility functions for calculations (same as front-end)
const calculateSum = (values) => values.reduce((acc, val) => acc + parseFloat(val || 0), 0);
const calculateAverage = (values) => calculateSum(values) / values.length;
const calculateMax = (values) => Math.max(...values.map(val => parseFloat(val || -Infinity)));
const calculateMin = (values) => Math.min(...values.map(val => parseFloat(val || Infinity)));

// Endpoint for handling formula calculations
app.post('/calculate', (req, res) => {
  const { formula, values } = req.body;

  let result = null;
  switch (formula) {
    case 'SUM':
      result = calculateSum(values);
      break;
    case 'AVERAGE':
      result = calculateAverage(values);
      break;
    case 'MAX':
      result = calculateMax(values);
      break;
    case 'MIN':
      result = calculateMin(values);
      break;
    default:
      result = 'Invalid formula';
  }

  res.json({ result });
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
