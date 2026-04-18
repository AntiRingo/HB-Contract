const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = Number(process.env.PORT) || 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

app.use('/api/templates', require('./routes/templates'));    
app.use('/api/purchase-orders', require('./routes/purchase-orders'));
app.use('/api/mappings', require('./routes/mappings'));

app.get(/.*/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public/index.html'));
}); 

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
     
