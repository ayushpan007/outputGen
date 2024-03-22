const { addSlideText, getSlideText, updateSlideText, deleteSlideText } = require('../controller/slideText');

module.exports = (router) => {
    router.post('/addSlideText',addSlideText);
    router.get('/getSlideText/:slideName',getSlideText)
    router.put('/updateSlideText/:slideName',updateSlideText)
    router.delete('/deleteSlideText/:slideName',deleteSlideText)
};