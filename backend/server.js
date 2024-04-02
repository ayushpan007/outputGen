const path = require('path')
const express = require('express')
const dotenv = require('dotenv').config()
const port = process.env.PORT || 3000
const cors = require('cors');
const routes = require('./routes');
const bodyParser = require("body-parser");
const helmet = require("helmet");
const connectDB = require('./config/db')
const authorize=require("./middleware/authMiddleware")

connectDB()
const app = express()
app.use(express.json())
app.use(helmet());
app.use(express.urlencoded({ extended: false }))
app.use(helmet.crossOriginResourcePolicy({ policy: "cross-origin" }));
app.use(bodyParser.json());
app.listen(port, () => console.log(`Server started on ${port}`))
//app.use(authorize);
app.use('/', routes); 