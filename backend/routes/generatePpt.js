const {generateAndSavePresentation}=require("../controller/pptCreate")

module.exports = (router) => {
    router.get('/generatePpt',generateAndSavePresentation);
  };