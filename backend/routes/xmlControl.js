const { xmlUpload,getXmlFile,deleteXmlFile } = require('../controller/xmlControls');

module.exports = (router) => {
  router.post('/uploadXml', xmlUpload);
  router.get("/getXml",getXmlFile)
  router.delete("/deleteXml",deleteXmlFile)
};