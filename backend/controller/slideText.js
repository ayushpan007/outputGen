const SlideText = require("../models/slideTextModel")

const addSlideText = async (req, res) => {
    try {
        const slide = await SlideText.create(req.body);
        res.status(201).json(slide);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
};

const getSlideText = async (req, res) => {
    try {
        const { slideName } = req.params;
        const slide = await SlideText.findOne({ slideName });
        if (!slide) {
            return res.status(404).json({ message: 'Slide not found' });
        }
        res.json(slide);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};

const updateSlideText = async (req, res) => {
    try {
        const { slideName } = req.params;
        const slide = await SlideText.findOneAndUpdate({ slideName }, req.body, { new: true });
        if (!slide) {
            return res.status(404).json({ message: 'Slide not found' });
        }
        res.json(slide);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
};

const deleteSlideText = async (req, res) => {
    try {
        const { slideName } = req.params;
        const slide = await SlideText.findOneAndDelete({ slideName });
        if (!slide) {
            return res.status(404).json({ message: 'Slide not found' });
        }
        res.status(204).json("Slide Content Deleted successfully")
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
};

module.exports = {
    addSlideText,
    getSlideText,
    updateSlideText,
    deleteSlideText
};