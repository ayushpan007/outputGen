const mongoose = require('mongoose')
const slideSchema  = mongoose.Schema({
    slideName: {
        type: String,
        unique: [true,"Slide name must be unique"],
        required: [true, 'Please add Slide Name']
    },
    textContent: {
        type: String,
        required: [true,'Please add the text for the slide']
    }
},)

module.exports = mongoose.model('SlideText', slideSchema)