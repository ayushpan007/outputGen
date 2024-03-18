const { getToken,registerUser } = require('../controller/generateToken');

module.exports = (router) => {
  router.post('/registerUser',registerUser );
  router.post("/getToken",getToken)
};