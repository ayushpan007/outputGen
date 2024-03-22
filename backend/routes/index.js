const express = require('express');
const router = express.Router();
const pingRoutes = require('./ping');
const generatePptRoute=require("./generatePpt")
const xmlControl=require("./xmlControl")
const generateTokenRoutes=require("./generateToken")
const slideTextRoutes=require("./slideText")

pingRoutes(router);
generatePptRoute(router)
xmlControl(router)
generateTokenRoutes(router)
slideTextRoutes(router)

module.exports = router;