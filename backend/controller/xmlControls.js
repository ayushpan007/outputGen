const multer = require('multer');
const path = require('path');
const fs = require('fs');

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const uploadDir = path.resolve(__dirname, "../", '../asset/');
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        cb(null, 'sample.xml');
    }
});

const upload = multer({
    storage: storage
}).single('xmlFile');

const xmlUpload = async (req, res) => {
    try {
        upload(req, res, function (err) {
            if (err instanceof multer.MulterError) {
                res.status(500).json({ error: err.message });
            } else if (err) {
                res.status(500).json({ error: 'Unknown error occurred.' });
            } else {
                res.status(200).json({ message: 'File uploaded successfully.' });
            }
        });
    } catch (error) {
        res.status(500).json({ error: 'Internal Server Error' });
    }
}

const getXmlFile = async (req, res) => {
    try {
        const filePath = path.resolve(__dirname, "../",'../asset/sample.xml');
        if (fs.existsSync(filePath)) {
            fs.readFile(filePath, 'utf8', (err, data) => {
                if (err) {
                    res.status(500).json({ error: 'Error reading the XML file.' });
                } else {
                    res.status(200).send(data);
                }
            });
        } else {
            res.status(404).json({ error: 'XML file not found.' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Internal Server Error' });
    }
}

const deleteXmlFile = async (req, res) => {
    try {
        const filePath = path.resolve(__dirname, "../",'../asset/sample.xml');
        if (fs.existsSync(filePath)) {
            fs.unlink(filePath, (err) => {
                if (err) {
                    res.status(500).json({ error: 'Error deleting the XML file.' });
                } else {
                    res.status(200).json({ message: 'XML file deleted successfully.' });
                }
            });
        } else {
            res.status(404).json({ error: 'XML file not found.' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Internal Server Error' });
    }
}

module.exports = {
    xmlUpload,
    getXmlFile,
    deleteXmlFile
};
