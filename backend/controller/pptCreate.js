const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { convertXMLtoJS, querySIPS, getYearAndMonth } = require("../utils/utils")
const xmlDataPath = path.resolve(__dirname, "../", '../asset/sample.xml');

let pptx = new pptxgen();
pptx.defineLayout({ name: 'standard', width: 13.32992126, height: 7.5 });
pptx.layout = 'standard';

const imageLine1 = path.resolve(__dirname, "../", '../images/line1.svg');
const imageLine2 = path.resolve(__dirname, "../", '../images/line2.svg');

const pageNumber = 0

const slide1 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const title = await querySIPS("case.client.title", json)
        const advisorName = await querySIPS("case.advisor.name", json)
        const advisorFirmName = await querySIPS("case.advisor.description", json)
        const advisorFirmAddress = "TBD Firm Address"
        const advisorFirmPhone = "TBD Firm Phone"
        const initialPlanDate = await querySIPS("case.client.initialPlanDate", json)
        const yearPlanCreated = (await getYearAndMonth(json)).planYear

        const titleOpts = {
            x: 2.7598425197,
            y: 1.2322834646,
            w: 9.6653543307,
            h: 0.8070866142,
            fontSize: 48,
            color: '376C8A',
            fontFace: "Barlow Black"
        };
        const subtitleOpts1 = {
            x: 2.7598425197,
            y: 2.3464566929,
            w: 9.6653543307,
            h: 1.2480314961,
            fontSize: 32,
            color: '000000',
            fontFace: "Barlow",
            isTextBox: true,
        };
        const contentOps = {
            x: 2.7598425197,
            y: 4.2401574803,
            w: 9.6653543307,
            h: 2.0275590551,
            fontSize: 18,
            color: "000000",
            fontFace: "Barlow",
            isTextBox: true,
        };
        slide.addText(`${title}`, titleOpts);
        slide.addText(`${yearPlanCreated} Financial Plan\nInitial Plan: ${initialPlanDate}`, subtitleOpts1);
        slide.addText(
            [
                { text: "Prepared by:\n", options: { bold: true } },
                { text: `${advisorName}\n` },
                { text: `${advisorFirmName}\n` },
                { text: `${advisorFirmAddress}\n` },
                { text: `${advisorFirmPhone}` },
            ],
            contentOps
        );
        slide.addShape(pptx.shapes.LINE, {
            x: 0.0,
            y: 3.9173228346,
            w: 13.334645669,
            h: 0.0,
            line: { color: "376C8A", width: 2 },
        });
        const imagePath1 = path.resolve(__dirname, "../", '../images/slide1/slide1.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images//slide1/slide1.2.svg');

        slide.addImage({ path: imagePath1, x: 0.905511811, y: 3.2125984252, w: 1.405511811, h: 1.405511811 });
        slide.addImage({ path: imagePath2, x: 1.2834645669, y: 3.594488189, w: 0.6456692913, h: 0.6456692913 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 1' });
    }
}


const slide2 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth
        const imagePath1 = path.resolve(__dirname, "../", '../images/slide2/slide2.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide2/slide2.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide2/slide2.3.svg');
        const imagePath4 = path.resolve(__dirname, "../", '../images/slide2/slide2.4.svg');
        const imagePath5 = path.resolve(__dirname, "../", '../images/slide2/slide2.5.svg');
        const imagePath6 = path.resolve(__dirname, "../", '../images/slide2/slide2.6.svg');
        const imagePath7 = path.resolve(__dirname, "../", '../images/slide2/slide2.7.svg');
        const imagePath8 = path.resolve(__dirname, "../", '../images/slide2/slide2.8.svg');
        slide.addImage({ path: imagePath1, x: 2.6338582677, y: 3.4291338583, w: 2.4133858268, h: 2.0905511811 });
        slide.addImage({ path: imagePath2, x: 3.1968503937, y: 4.062992126, w: 1.1732283465, h: 0.8267716535 });
        slide.addImage({ path: imagePath3, x: 4.5590551181, y: 2.3858267717, w: 2.1850393701, h: 2.0905511811 });
        slide.addImage({ path: imagePath4, x: 5.1220472441, y: 2.9724409449, w: 1.0551181102, h: 0.9133858268 });
        slide.addImage({ path: imagePath5, x: 6.3622047244, y: 3.4291338583, w: 2.3031496063, h: 2.0905511811 });
        slide.addImage({ path: imagePath6, x: 6.9724409449, y: 4.0196850394, w: 0.9133858268, h: 0.9133858268 });
        slide.addImage({ path: imagePath7, x: 8.1771653543, y: 2.3858267717, w: 2.1850393701, h: 2.0905511811 });
        slide.addImage({ path: imagePath8, x: 8.9094488189, y: 2.9724409449, w: 0.8267716535, h: 0.8267716535 });

        slide.addText("A Tailored Financial Plan for the Rest of Your Life", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, });
        slide.addText("Validate your current situation", { x: 2.6338582677, y: 5.7204724409, w: 2.4881889764, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Your current situation, will you have enough?", { x: 4.2795275591, y: 1.6181102362, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Your Plan after income and tax optimizations", { x: 6.094488189, y: 5.7440944882, w: 2.7283464567, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Focusing on your priorities", { x: 8.0905511811, y: 1.5511811024, w: 1.9960629921, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });

        //footer
        //think about creating one master template for all of this
        slide.addShape(pptx.shapes.LINE, {
            x: 0.0,
            y: 7.0826771654,
            w: 13.334645669,
            h: 0.0,
            line: { color: "A6A6A6", width: 0.1 },
        });
        slide.addText(`${month} ${year}`, { x: 0.6614173228, y: 7.1811023622, w: 3, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, });
        slide.addText(`${year} Financial Plan Initial Plan`, { x: 4.4173228346, y: 7.1811023622, w: 4.5, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, align: "center" });
        slide.addText(`${pageNumber}`, { x: 9.6732283465, y: 7.1811023622, w: 3, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, align: "right" });
        //two lines at the left corner
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.688976378, h: 0.688976378 });
    } catch {
        res.status(500).json({ message: 'Error in slide 2' });
    }
}

const slide3 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth
        const imagePath1 = path.resolve(__dirname, "../", '../images/slide3/slide3.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide3/slide3.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide3/slide3.3.svg');
        const imagePath4 = path.resolve(__dirname, "../", '../images/slide3/slide3.4.svg');
        const imagePath5 = path.resolve(__dirname, "../", '../images/slide3/slide3.5.svg');
        const imagePath6 = path.resolve(__dirname, "../", '../images/slide3/slide3.6.svg');
        const imagePath7 = path.resolve(__dirname, "../", '../images/slide3/slide3.7.svg');
        const imagePath8 = path.resolve(__dirname, "../", '../images/slide3/slide3.8.svg');
        const imagePath9 = path.resolve(__dirname, "../", '../images/slide3/slide3.9.svg');
        const imagePath10 = path.resolve(__dirname, "../", '../images/slide3/slide3.10.svg');
        slide.addImage({ path: imagePath1, x: 1.8385826772, y: 3.4330708661, w: 2.4133858268, h: 2.0905511811 });
        slide.addImage({ path: imagePath2, x: 2.405511811, y: 4.0669291339, w: 1.1732283465, h: 0.8267716535 });
        slide.addImage({ path: imagePath3, x: 3.7677165354, y: 2.3897637795, w: 2.1850393701, h: 2.0905511811 });
        slide.addImage({ path: imagePath4, x: 4.3307086614, y: 2.9763779528, w: 1.0551181102, h: 0.9133858268 });
        slide.addImage({ path: imagePath5, x: 5.5708661417, y: 3.4330708661, w: 2.3031496063, h: 2.0905511811 });
        slide.addImage({ path: imagePath6, x: 6.2519685039, y: 4.0669291339, w: 0.8267716535, h: 0.8267716535 });
        slide.addImage({ path: imagePath7, x: 7.38582677173, y: 2.3897637795, w: 2.1850393701, h: 2.0905511811 });
        slide.addImage({ path: imagePath8, x: 8.0157480315, y: 2.9763779528, w: 0.9133858268, h: 0.9133858268 });
        slide.addImage({ path: imagePath9, x: 9.188976378, y: 3.4330708661, w: 2.3031496063, h: 2.0905511811 });
        slide.addImage({ path: imagePath10, x: 9.92519685049, y: 4.0669291339, w: 0.8267716535, h: 0.8267716535 });

        slide.addText("A Tailored Financial Plan for the Rest of Your Life", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, });
        slide.addText("Validate your current situation", { x: 1.8385826772, y: 5.7244094488, w: 2.4881889764, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Your current situation, will you have enough?", { x:3.4842519685, y: 1.6181102362, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Stress testing your plan", { x: 5.5118110236, y: 5.7244094488, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Your Plan after income and tax optimizations", { x: 6.9448818898, y: 1.6181102362, w: 2.7283464567, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("Focusing on your priorities", { x: 9.1850393701, y: 5.7244094488, w: 1.9960629921, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });

        //footer
        //think about creating one master template for all of this
        slide.addShape(pptx.shapes.LINE, {
            x: 0.0,
            y: 7.0826771654,
            w: 13.334645669,
            h: 0.0,
            line: { color: "A6A6A6", width: 0.1 },
        });
        slide.addText(`${month} ${year}`, { x: 0.6614173228, y: 7.1811023622, w: 3, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, });
        slide.addText(`${year} Financial Plan Initial Plan`, { x: 4.4173228346, y: 7.1811023622, w: 4.5, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, align: "center" });
        slide.addText(`${pageNumber}`, { x: 9.6732283465, y: 7.1811023622, w: 3, h: 0.1692913386, fontSize: 10, color: "898989", fontFace: "Calibri (Body)", isTextBox: true, align: "right" });
        //two lines at the left corner
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.688976378, h: 0.688976378 });
    } catch {
        res.status(500).json({ message: 'Error in slide 3' });
    }
}


const generateAndSavePresentation = async (_, res) => {
    const finalPresentationFolder = path.join(__dirname, "../", 'finalPresentation');
    const filePath = path.join(finalPresentationFolder, "presentation.pptx");

    try {
        if (!fs.existsSync(finalPresentationFolder)) {
            fs.mkdirSync(finalPresentationFolder);
        }
        await slide1();
        await slide2()
        await slide3();

        await pptx.writeFile({ fileName: filePath });
        res.status(200).json({ message: 'Presentation created and saved successfully' });
    } catch (error) {
        console.error("Error saving presentation:", error);
        res.status(500).json({ message: 'Error saving presentation' });
    }
};
module.exports = { generateAndSavePresentation };