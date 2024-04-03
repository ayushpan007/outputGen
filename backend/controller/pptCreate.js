// const pptxgen = require("pptxgenjs");
// const fs = require("fs");
// const path = require("path");

// // Create a new presentation
// let pptx = new pptxgen();

// // Add a slide with a title and content
// let slide = pptx.addSlide();
// slide.addText("My Presentation", { x:1.5, y:0.5, fontSize:24, color:"363636" });
// slide.addText("Hello, world!", { x:1.5, y:2.0, fontSize:18, color:"363636" });

// // Add another slide
// let slide2 = pptx.addSlide();
// slide2.addText("Second Slide", { x:1.5, y:0.5, fontSize:24, color:"363636" });
// slide2.addText("This is the content of the second slide.", { x:1.5, y:2.0, fontSize:18, color:"363636" });

// const finalPresentationFolder = path.join(__dirname, "../",'finalPresentation');

// // Save the presentation to a file in the finalPresentation folder
// const filePath = path.join(finalPresentationFolder, "my_presentation.pptx");
// pptx.writeFile({ fileName: filePath }, (error) => {
//     if (error) {
//         console.log("Error saving presentation:", error);
//     } else {
//         console.log("Presentation saved successfully!");
//     }
// });


// const pptxgen = require("pptxgenjs");
// const fs = require("fs");
// const path = require("path");

// let pptx = new pptxgen();
// pptx.defineLayout({ name: 'standard', width: 13.33, height: 7.5 });
// pptx.layout = 'standard';

// function slide1() {
//     let slide = pptx.addSlide();
//     slide.addText("You Will Not Have Enough Income for Your Desired Lifestyle", { x: 0.5, y: 0.75, w: "100%", fontSize: "32", color: "376C8A", fontFace: "Barlow Condensed SemiBold" });
//     slide.addText("Plan Summary", { x: 1, y: 1, fontSize: 24, color: "376C8A", fontFace: "Arial" });

//     const tableData = [
//         ["$", "Starting"],
//         ["Total after-tax income", ""],
//         ["Savings at end of plan", ""],
//         ["Lifetime Estate Value", ""],
//         ["Total taxes paid by end of plan", ""]
//     ];

//     const tableOpts = {
//         x: 1.5,
//         y: 2.5,
//         w: 5,
//         h: 4,
//         fill: { color: "FFFFFF" },
//         fontSize: 14,
//         color: "000000",
//         border: { pt: 1, color: "376C8A" },
//         align: "center",
//         valign: "middle"
//     };
//     slide.addTable(tableData, tableOpts);

//     const age = [62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102];
//     const seriesSavings = [2245956, 2313547, 2407270, 2504085, 2538300, 2572425, 2606022, 2639055, 2671469, 2703204, 2734189, 2764346, 2793587, 2821820, 2846070, 2868764, 2889771, 2908862, 2925097, 2938101, 2947726, 2953645, 2955713, 2953600, 2946998, 2935729, 2919283, 2956438, 3007229, 3054323, 3097278, 3135963, 3169963, 3199331, 3223230, 3241819, 3254687, 3262216, 3264144, 3260288];
//     const seriesWithdrawals = [41679, 43896, 21112, 22671, 90078, 91843, 94038, 96243, 98472, 100729, 103022, 105353, 107729, 110149, 115491, 118207, 120971, 123880, 127629, 131609, 135572, 139692, 143768, 147978, 152285, 156542, 161072, 165490, 171114, 177258, 183654, 189971, 196484, 202708, 209531, 215918, 222446, 228300, 234144, 239887];
//     const targetAfterTaxIncome = [120000, 123000, 126075, 129227, 132458, 135769, 139163, 142642, 146208, 149864, 153610, 157450, 161387, 165421, 169557, 173796, 178141, 182594, 187159, 191838, 196634, 201550, 206588, 211753, 217047, 222473, 228035, 233736, 239579, 245569, 251708, 258001, 264451, 271062, 277838, 284784, 291904, 299202, 306682, 314349];
//     const sumPension = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
//     const sumSocialSecurity = [0, 0, 30926, 31699, 57961, 59282, 60635, 62018, 63433, 64881, 66362, 67878, 69429, 71016, 72639, 74299, 75999, 77737, 79516, 81336, 83197, 85103, 87052, 89047, 91087, 93174, 95312, 97497, 99733, 102023, 104364, 106760, 109211, 111720, 114287, 116913, 119601, 122351, 125165, 128044];
//     const wages = [85000, 85850, 86708, 87576, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

//     // // Add the area chart with secondary axis
//      const dataChartArea = [
//         {
//             name: "seriesSavings",
//             labels: age,
//             values: seriesSavings,
//             options: { yAxis: 'secondary' }
//         }
//     ];
//     slide.addChart(pptx.ChartType.area, dataChartArea, { x: 1.5, y: 1.5, w: 4.0, h: 2.0 },);

//     //Add the line chart
//     const dataChartLine = [
//         {
//             name: "targetAfterTaxIncome",
//             labels: age,
//             values: targetAfterTaxIncome
//         }
//     ];
//     slide.addChart(pptx.ChartType.line, dataChartLine, { x: 1.5, y: 4.0, w: 4.0, h: 2.0 });

//     // Add the stacked column chart
//     // const dataChartColumn = [
//     //     {
//     //         name: "seriesWithdrawals",
//     //         labels: age,
//     //         values: seriesWithdrawals
//     //     },
//     //     {
//     //         name: "sumPension",
//     //         labels: age,
//     //         values: sumPension
//     //     },
//     //     {
//     //         name: "sumSocialSecurity",
//     //         labels: age,
//     //         values: sumSocialSecurity
//     //     },
//     //     {
//     //         name: "wages",
//     //         labels: age,
//     //         values: wages
//     //     }
//     // ];
//     // slide.addChart(pptx.ChartType.column, dataChartColumn, { x: 6.0, y: 1.5, w: 4.0, h: 3.0 });
// }

// const generateAndSavePresentation = async (_, res) => {
//     const slideFunctions = [slide1];
//     const finalPresentationFolder = path.join(__dirname, "../", 'finalPresentation');
//     const filePath = path.join(finalPresentationFolder, "presentation.pptx");

//     try {
//         if (!fs.existsSync(finalPresentationFolder)) {
//             fs.mkdirSync(finalPresentationFolder);
//         }

//         slideFunctions.forEach(slideFunction => {
//             slideFunction();
//         });

//         await pptx.writeFile({ fileName: filePath });
//         res.status(200).json({ message: 'Presentation created and saved successfully' });
//     } catch (error) {
//         console.log("Error saving presentation:", error);
//         res.status(500).json({ message: 'Error saving presentation' });
//     }
// };

// module.exports = { generateAndSavePresentation };
// const slide1=async () =>{
//     let slide = pptx.addSlide();
//     const json = await convertXMLtoJS(xmlDataPath);
//     const value= await querySIPS("case.client.title", json);
//     slide.addText("You Will Not Have Enough Income for Your Desired Lifestyle", { x: 0.5, y: 0.75, w: "100%", fontSize: "32", color: "376C8A", fontFace: "Barlow Condensed SemiBold" });
//     slide.addText("Plan Summary", { x: 1, y: 1, fontSize: 24, color: "376C8A", fontFace: "Arial" });

//     const tableData = [
//         ["$", "Starting"],
//         ["Total after-tax income", "$64726"],
//         ["Savings at end of plan", "$3434"],
//         ["Lifetime Estate Value", "$467348"],
//         ["Total taxes paid by end of plan", "$38443"]
//     ];
//     const tableOpts = {
//         x: 0.5,
//         y: 2.5,
//         w: 5,
//         h: 3,
//         fill: { color: "FFFFFF" },
//         fontSize: 14,
//         color: "000000",
//         border: { pt: 1, color: "376C8A" },
//         align: "center",
//         valign: "middle"
//     };
//     slide.addTable(tableData, tableOpts);

//     let opts = {
//         x: 5.5, 
//         y: 2.0,  
//         w: 7.0,   
//         h: 5.0,
//         barDir: "col",
//         catAxisLabelColor: "666666",
//         catAxisLabelFontFace: "Arial",
//         catAxisLabelFontSize: 12,
//         catAxisOrientation: "minMax",
//         showLegend: false,
//         showTitle: false,

//         valAxes: [
//             {
//                 showValAxisTitle: true,
//                 valAxisTitle: "Primary Value Axis",
//             },
//             {
//                 showValAxisTitle: true,
//                 valAxisTitle: "Secondary Value Axis",
//                 valGridLine: { style: "none" },
//             },
//         ],

//         catAxes: [
//             {
//                 catAxisTitle: "Primary Category Axis",
//             },
//             {
//                 catAxisHidden: false,
//             },
//         ],
//     };

//     const age = [62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102];
//     let chartTypes = [
//         {
//             type: pptx.charts.AREA,
//             data: [
//                 {
//                     name: "Savings at end of plan",
//                     labels: age,
//                     values: [2245956,2313547,2407270,2504085,2538300,2572425,2606022,2639055,2671469,2703204,2734189,2764346,2793587,2821820,2846070,2868764,2889771,2908862,2925097,2938101,2947726,2953645,2955713,2953600,2946998,2935729,2919283,2956438,3007229,3054323,3097278,3135963,3169963,3199331,3223230,3241819,3254687,3262216,3264144,3260288]
//                 },
//             ],
//             options: {
//                 chartColors: ["#D6DCE5"],
//                 //barGrouping: "standard",
//                 secondaryValAxis: !!opts.valAxes,
//                 secondaryCatAxis: !!opts.catAxes,
//             },
//         },
//         {
//             type: pptx.charts.BAR,
//             data: [
//                 {
//                     name: "Total after-tax income",
//                     labels: age,
//                     values:[41679,43896,21112,22671,90078,91843,94038,96243,98472,100729,103022,105353,107729,110149,115491,118207,120971,123880,127629,131609,135572,139692,143768,147978,152285,156542,161072,165490,171114,177258,183654,189971,196484,202708,209531,215918,222446,228300,234144,239887],

//                 },
//             ],
//             options: {
//                 chartColors: ["0000FF"],
//                 barGrouping: "stacked",
//             },
//         },
//         {
//             type: pptx.charts.LINE,
//             data: [
//                 {
//                     name: "Current",
//                     labels: age,
//                     values: [1200000, 1230000, 126075, 129227, 132458, 135769, 139163, 142642, 146208, 149864, 153610, 157450, 161387, 165421, 169557, 173796, 178141, 182594, 187159, 191838, 196634, 201550, 206588, 211753, 217047, 222473, 228035, 233736, 239579, 245569, 251708, 258001, 264451, 271062, 277838, 284784, 291904, 299202, 306682, 314349],
//                 },
//             ],
//             options: {
//                 barGrouping: "standard",
//                 secondaryValAxis: !!opts.valAxes,
//                 secondaryCatAxis: !!opts.catAxes,
//             },
//         },
//     ];
//     slide.addChart(chartTypes, opts);

// }

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { convertXMLtoJS, querySIPS, getYearAndMonth, alternateTextSlide44, getTextForSlide } = require("../utils/utils")
const { dataForSlide5 } = require("./slideContentFetcher")
const SlideText = require("../models/slideTextModel")
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
        res.status(500).json({ message: 'Error in slide 1', error: err.message });
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

        slide.addText("A Tailored Financial Plan for the Rest of Your Life", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true });
        slide.addText("Validate your current situation", { x: 2.6338582677, y: 5.7204724409, w: 2.4881889764, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Your current situation, will you have enough?", { x: 4.2795275591, y: 1.6181102362, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Your Plan after income and tax optimizations", { x: 6.094488189, y: 5.7440944882, w: 2.7283464567, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Focusing on your priorities", { x: 8.0905511811, y: 1.5511811024, w: 1.9960629921, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });

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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 2', error: err.message });
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
        slide.addText("Validate your current situation", { x: 1.8385826772, y: 5.7244094488, w: 2.4881889764, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Your current situation, will you have enough?", { x: 3.4842519685, y: 1.6181102362, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Stress testing your plan", { x: 5.5118110236, y: 5.7244094488, w: 2.657480315, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Your Plan after income and tax optimizations", { x: 6.9448818898, y: 1.6181102362, w: 2.7283464567, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });
        slide.addText("Focusing on your priorities", { x: 9.1850393701, y: 5.7244094488, w: 1.9960629921, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", lineSpacingMultiple: 1.14 });

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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 3', error: err.message });
    }
}

const slide4 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        const imagePath1 = path.resolve(__dirname, "../", '../images/slide4/slide4.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide4/slide4.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide4/slide4.3.svg');
        const imagePath4 = path.resolve(__dirname, "../", '../images/slide4/slide4.4.svg');
        const imagePath5 = path.resolve(__dirname, "../", '../images/slide4/slide4.5.svg');

        slide.addImage({ path: imagePath1, x: 6.7007874016, y: 2.1299212598, w: 0.937007874, h: 0.937007874 });
        slide.addImage({ path: imagePath1, x: 6.7007874016, y: 3.2637795276, w: 0.937007874, h: 0.937007874 });
        slide.addImage({ path: imagePath1, x: 6.7007874016, y: 4.3976377953, w: 0.937007874, h: 0.937007874 });
        slide.addImage({ path: imagePath1, x: 6.7007874016, y: 5.531496063, w: 0.937007874, h: 0.937007874 });

        slide.addImage({ path: imagePath2, x: 6.8976377953, y: 2.2952755906, w: 0.5393700787, h: 0.6023622047 });
        slide.addText("Careful tax avoidance by sequencing withdrawal timing.", { x: 7.79921259847, y: 2.2677165354, w: 4.1496062992, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true });

        slide.addImage({ path: imagePath3, x: 6.8779527559, y: 3.4409448819, w: 0.5826771654, h: 0.5826771654 });
        slide.addText("Return optimization by improved asset allocation.", { x: 7.79921259847, y: 3.4015748031, w: 4.1496062992, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true });

        slide.addImage({ path: imagePath4, x: 6.8779527559, y: 4.5748031496, w: 0.5787401575, h: 0.5787401575 });
        slide.addText("Improved budgeting and goal setting for the retirement period.", { x: 7.79921259847, y: 4.5354330709, w: 4.1496062992, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true });

        slide.addImage({ path: imagePath5, x: 6.8779527559, y: 5.7559055118, w: 0.5787401575, h: 0.4842519685 });
        slide.addText("Optimization of Social Security withdrawal timing.", { x: 7.79921259847, y: 5.6692913386, w: 4.1496062992, h: 0.6535433071, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true });


        slide.addText("How We Measure Success", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, });
        slide.addText("Lifetime Estate Value", { x: 1.3818897638, y: 1.5905511811, w: 4.7204724409, h: 0.3779527559, fontSize: 22, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });
        slide.addText("Our Wealth Drivers", { x: 6.7007874016, y: 1.5905511811, w: 5.2519685039, h: 0.3779527559, fontSize: 22, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });
        slide.addText([
            { text: "Lifetime estate value measures the wealth your plan can deliver during retirement.", options: { fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, lineSpacingMultiple: 1.14, bullet: { code: "2022", color: "000000" } } },
            { text: "We will see how changes to your plan can change the lifetime estate value.", options: { fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, lineSpacingMultiple: 1.14, bullet: { code: "2022", color: "000000" } } }
        ], {
            x: 1.3818897638,
            y: 2.1062992126,
            w: 4.7204724409,
            h: 1.7755905512,
        })

        slide.addText("Total After Tax Income", { x: 2.3228346457, y: 4.1338582677, w: 2.842519685, h: 0.311023622, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });
        slide.addText("+  Savings at End of Plan", { x: 2.3228346457, y: 4.5590551181, w: 2.842519685, h: 0.311023622, fontSize: 18, color: "000000", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center" });

        //line under saving at end of plan
        slide.addShape(pptx.shapes.LINE, {
            x: 2.3228346457,
            y: 4.9842519685,
            w: 2.842519685,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        slide.addText("Lifetime Estate Value", { x: 2.3228346457, y: 5.1023622047, w: 2.842519685, h: 0.311023622, fontSize: 18, color: "E6B127", fontFace: "Barlow", wordWrap: true, isTextBox: true, align: "center", bold: true });
        //line under the lifetime estate value
        slide.addShape(pptx.shapes.LINE, {
            x: 1.3818897638,
            y: 2.031496063,
            w: 4.7204724409,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        //line under our wealth driver
        slide.addShape(pptx.shapes.LINE, {
            x: 6.7007874016,
            y: 2.031496063,
            w: 5.2519685039,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 4', error: err.message });
    }
}

const slide5 = async () => {
    try {
        let slide = pptx.addSlide()
        const {
            c1FirstName,
            c1LastName,
            c2FirstName,
            c2LastName,
            c1Age,
            c2Age,
            dependents,
            beneficiaries,
            c1RetAge,
            c2RetAge,
            c1Income,
            c2Income,
            c1SSAmt,
            c2SSAmt,
            c1SSAmtCurrently,
            c2SSAmtCurrently,
            totalSavings,
            c1PensionAmt,
            c2PensionAmt,
            marginalTaxRate,
            federalTaxRate,
            sumTaxDeferred,
            sumTaxableSaving,
            sumTaxFreeRoth,
            liabilities,
            liabilitiesTotal,
            totalIncome,
            fedRate,
            stateRes,
            year,
            month,
        } = await dataForSlide5(xmlDataPath)

        slide.addText("Let’s Validate Your Information First", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, });

        let tableData1 = [
            [
                { text: "Family Members", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `${c1FirstName}\n${c1LastName}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `age ${c1Age}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `${c2FirstName}\n${c2LastName}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `age ${c2Age}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
        ];

        // if (beneficiaries) {
        //     tableData1.push([
        //         { text: `Beneficiaries`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
        //         { text: `${beneficiaries}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", bullet: true } }
        //     ]);
        // }
        if (beneficiaries) {
            const beneficiaryNames = beneficiaries.map(beneficiary => beneficiary).join('\n'); // Concatenate beneficiary names into a single string
            tableData1.push([
                { text: `Beneficiaries`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${beneficiaryNames}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", bullet: true } } // Display concatenated names in a single column
            ]);
        }

        const tableOpts1 = {
            x: 0.7244094488,
            y: 1.2992125984,
            w: 3.4015748031,
            h: 2.0354330709,
            rowH: [0.25, 0.25, 0.25, 0.25],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData1, tableOpts1);


        let tableData2 = [
            [
                { text: "Planning Priorities", options: { fontSize: 14, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: "Priority 1", options: { fontSize: 14, fontFace: "Barlow", valign: "middle", bullet: true } }
            ]
        ];
        const tableOpts2 = {
            x: 0.7244094488,
            y: 4.0393700787,
            w: 3.4015748031,
            h: 1.1338582677,
            rowH: [0.25, 0.25],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData2, tableOpts2);

        let tableData3 = [
            [
                { text: "Tax Information", options: { fontSize: 14, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `${marginalTaxRate}% marginal tax bracket`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `${federalTaxRate}% effective federal tax rate`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ]
        ];
        const tableOpts3 = {
            x: 0.7244094488,
            y: 5.8779527559,
            w: 3.4015748031,
            h: 1.0196850394,
            rowH: [0.25, 0.25],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData3, tableOpts3);

        let tableData4 = [
            [
                { text: "Current Income", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `${c1FirstName}\nretire at age ${c1RetAge}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${c1Income}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `${c2FirstName}\nretire at age ${c2RetAge}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${c2Income}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
        ];
        const tableOpts4 = {
            x: 4.4803149606,
            y: 1.2992125984,
            w: 4.5157480315,
            h: 1.468503937,
            rowH: [0.25, 0.25, 0.25],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData4, tableOpts4);

        let tableData5 = [
            [
                { text: "Retirement Income", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
        ];

        if (c1SSAmt !== '$0') {
            tableData5.push(
                [
                    { text: `${c1FirstName} Social Security\n at age 67`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                    { text: `${c1SSAmt}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                ],
            )
        }

        if (c2SSAmt !== '$0') {
            tableData5.push(
                [
                    { text: `${c2FirstName} Social Security\n at age 67`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                    { text: `${c2SSAmt}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                ],
            )
        }

        if (c1SSAmtCurrently !== '$0') {
            tableData5.push(
                [
                    { text: `${c1FirstName} Social Security\n currently`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                    { text: `${c1SSAmtCurrently}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                ],
            )
        }

        if (c2SSAmtCurrently !== '$0') {
            tableData5.push(
                [
                    { text: `${c2FirstName} Social Security\n currently`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                    { text: `${c2SSAmtCurrently}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                ],
            )
        }

        if (c1PensionAmt !== '$0') {
            tableData5.push
                (
                    [
                        { text: `${c1FirstName} Pension`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                        { text: `${c1PensionAmt}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                    ],
                )
        }

        if (c2PensionAmt !== '$0') {
            (
                [
                    { text: `${c2FirstName} Pension`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                    { text: `${c2PensionAmt}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
                ]
            )
        }

        const tableOpts5 = {
            x: 4.4803149606,
            y: 2.937007874,
            w: 4.5157480315,
            h: 2.1338582677,
            rowH: [0.25, 0.25, 0.25, 0.25],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData5, tableOpts5);

        let tableData6 = [
            [
                { text: "Other Income", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `Income Name`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `Income Amt`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ]
        ];
        const tableOpts6 = {
            x: 4.4803149606,
            y: 5.3622047244,
            w: 4.5157480315,
            h: 1.6653543307,
            rowH: 0.25,
            border: { color: "376C8A" },
        };
        slide.addTable(tableData6, tableOpts6);

        let tableData7 = [
            [
                { text: "Target Retirement Income (after tax)", options: { fontSize: 14, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `${totalIncome}/ year`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } },
            ]
        ];
        const tableOpts7 = {
            x: 9.3503937008,
            y: 1.3031496063,
            w: 3.4015748031,
            h: 0.6653543307,
            rowH: 0.25,
            border: { color: "376C8A" },
        };
        slide.addTable(tableData7, tableOpts7);

        let tableData8 = [
            [
                { text: "Savings", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `Taxable\nSavings`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${sumTaxableSaving}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `Tax-Deferred\nIRA, 401(k), etc`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${sumTaxDeferred}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `Tax-Free\nRoth IRA`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `${sumTaxFreeRoth}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],
            [
                { text: `Total Savings`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left", bold: true } },
                { text: `${totalSavings}`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center", bold: true } }
            ]
        ];
        const tableOpts8 = {
            x: 9.3503937008,
            y: 2.4173228346,
            w: 3.3858267717,
            h: 2.3661417323,
            rowH: 0.25,
            border: { color: "376C8A" },
        };
        slide.addTable(tableData8, tableOpts8);

        let tableData9 = [
            [
                { text: "Other Assets and Liabilities", options: { fontSize: 14, fontFace: "Barlow", bold: true, colspan: 2, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: `Asset Name`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "left" } },
                { text: `Asset Value`, options: { fontSize: 14, fontFace: "Barlow", valign: "middle", align: "center" } }
            ],

        ];
        const tableOpts9 = {
            x: 9.3503937008,
            y: 5.2283464567,
            w: 3.3228346457,
            h: 1.6653543307,
            rowH: 0.25,
            border: { color: "376C8A" },
        };
        slide.addTable(tableData9, tableOpts9);



        //footer
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 5', error: err.message });
    }
}


const slide37 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        //title
        slide.addText("Appendix – Inflation", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });
        slide.addText("U.S. Historical Annual Inflation Rate", { x: 0.6614173228, y: 1.4212598425, w: 12.011811024, h: 0.2401574803, fontSize: 14, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });

        slide.addShape(pptx.shapes.LINE, {
            x: 0.6614173228,
            y: 1.7244094488,
            w: 12.011811024,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        const imagePath1 = path.resolve(__dirname, "../", '../images/slide37/slide37.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide37/slide37.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide37/slide37.3.svg');

        slide.addImage({ path: imagePath1, x: 0.6614173228, y: 1.7440944882, w: 12.011811024, h: 2.2007874016 });
        slide.addText(
            [
                { text: "1982: U.S. Federal Reserve implements modern-day inflation management toolkit", options: { fontSize: 10, fontFace: "Barlow" } }
            ],
            { x: 6.125984252, y: 2.1850393701, w: 2.0157480315, h: 0.6732283465, align: "center", isTextBox: true, wordWrap: true, line: { color: "7F7F7F", width: "1" }, lineSpacingMultiple: 1.14 }
        );
        slide.addImage({ path: imagePath2, x: 0.6614173228, y: 4.1102362205, w: 10.173228346, h: 1.4015748031 });
        slide.addText(
            [
                { text: "In your plan", options: { fontSize: 10, fontFace: "Barlow" } }
            ],
            { x: 11.578740157, y: 4.8346456693, w: 1.0196850394, h: 0.2913385827, align: "center", isTextBox: true, wordWrap: true, line: { color: "E6B127", width: "1" }, lineSpacingMultiple: 1.14 }
        );
        slide.addShape(pptx.shapes.LINE, {
            x: 10.905511811,
            y: 4.9803149606,
            w: 0.6732283465,
            h: 0.0,
            line: { color: "E6B127", width: "1", beginArrowType: "oval", },
        });

        slide.addImage({ path: imagePath3, x: 5.5118110236, y: 2.5236220472, w: 0.6141732283, h: 0.6181102362 });
        slide.addText(
            [
                { text: "References", options: { fontSize: 12, color: "767171", fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bold: true } },
                { text: "https://www.macrotrends.net/countries/USA/united-states/inflation-rate-cpi", options: { hyperlink: { url: "https://www.macrotrends.net/countries/USA/united-states/inflation-rate-cpi" }, color: "0070C0", bullet: { type: "number" } }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14 },
                { text: "https://www.federalreservehistory.org/essays/great-inflation", options: { hyperlink: { url: "https://www.federalreservehistory.org/essays/great-inflation" }, color: "0070C0", bullet: { type: "number" } }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, },
                { text: "3.7% inflation is used in stress testing your plan", options: { fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bullet: { type: "number" }, color: "767171" } },
                { text: "2.2% inflation is used as the inflation assumption for your plan", options: { fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bullet: { type: "number" }, color: "767171" } }
            ],
            { x: 0.6614173228, y: 6.0078740157, w: 12.011811024, h: 1.0118110236, isTextBox: true, wordWrap: true }
        )

        //footer
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 37', error: err.message });
    }
}

const slide38 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        //title
        slide.addText("Appendix – Social Security Cost of Living Adjustment (COLA)", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });
        slide.addText("U.S Historical Social Security Annual Cost of Living Adjustments", { x: 0.6614173228, y: 1.4212598425, w: 12.011811024, h: 0.2401574803, fontSize: 14, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });

        slide.addShape(pptx.shapes.LINE, {
            x: 0.6614173228,
            y: 1.7244094488,
            w: 12.011811024,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        const imagePath1 = path.resolve(__dirname, "../", '../images/slide38/slide38.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide38/slide38.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide38/slide38.3.svg');

        slide.addImage({ path: imagePath1, x: 0.6614173228, y: 1.7440944882, w: 12.011811024, h: 2.2007874016 });
        slide.addText(
            [
                { text: "1982: U.S. Federal Reserve implements modern-day inflation management toolkit", options: { fontSize: 10, fontFace: "Barlow" } }
            ],
            { x: 3.8818897638, y: 2.1850393701, w: 2.0157480315, h: 0.6732283465, align: "center", isTextBox: true, wordWrap: true, line: { color: "7F7F7F", width: "1" }, lineSpacingMultiple: 1.14 }
        );
        slide.addImage({ path: imagePath2, x: 0.6614173228, y: 4.1102362205, w: 10.173228346, h: 1.4015748031 });
        slide.addText(
            [
                { text: "In your plan", options: { fontSize: 10, fontFace: "Barlow" } }
            ],
            { x: 11.578740157, y: 4.8346456693, w: 1.0196850394, h: 0.2913385827, align: "center", isTextBox: true, wordWrap: true, line: { color: "E6B127", width: "1" }, lineSpacingMultiple: 1.14 }
        );
        slide.addShape(pptx.shapes.LINE, {
            x: 10.905511811,
            y: 4.9803149606,
            w: 0.6732283465,
            h: 0.0,
            line: { color: "E6B127", width: "1", beginArrowType: "oval", },
        });

        slide.addImage({ path: imagePath3, x: 3.2677165354, y: 2.5236220472, w: 0.6141732283, h: 0.6181102362 });
        slide.addText(
            [
                { text: "References", options: { fontSize: 12, color: "767171", fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bold: true } },
                { text: "https://www.ssa.gov/oact/cola/colaseries.html", options: { hyperlink: { url: "https://www.ssa.gov/oact/cola/colaseries.html" }, color: "0070C0", bullet: { type: "number" } }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14 },
                { text: "https://www.federalreservehistory.org/essays/great-inflation", options: { hyperlink: { url: "https://www.federalreservehistory.org/essays/great-inflation" }, color: "0070C0", bullet: { type: "number" } }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, },
                { text: "2.3% Social Security annual cost of living adjustment (COLA) is used in your plan", options: { fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bullet: { type: "number" }, color: "767171" } },
            ],
            { x: 0.6614173228, y: 6.0078740157, w: 12.011811024, h: 1.0118110236, isTextBox: true, wordWrap: true }
        )

        //footer
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 38', error: err.message });
    }
}

const slide39 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        //title
        slide.addText("Appendix – Conventional Wisdom Withdrawal Strategy", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });

        slide.addText("Withdrawing One Account At a Time", { x: 0.6614173228, y: 1.4212598425, w: 12.011811024, h: 0.2401574803, fontSize: 14, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });
        slide.addShape(pptx.shapes.LINE, {
            x: 0.6614173228,
            y: 1.7244094488,
            w: 8.3149606299,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        const imagePath1 = path.resolve(__dirname, "../", '../images/slide39/slide39.1.svg');
        slide.addImage({ path: imagePath1, x: 0.6614173228, y: 1.7598425197, w: 8.3149606299, h: 2.2913385827 });

        slide.addText("Taxes Owed", { x: 0.6614173228, y: 4.1456692913, w: 8.3149606299, h: 0.2401574803, fontSize: 14, color: "376C8A", fontFace: "Barlow", wordWrap: true, isTextBox: true, bold: true });
        slide.addShape(pptx.shapes.LINE, {
            x: 0.6614173228,
            y: 4.4527559055,
            w: 8.3149606299,
            h: 0.0,
            line: { color: "376C8A", width: 0.1 },
        });

        const imagePath2 = path.resolve(__dirname, "../", '../images/slide39/slide39.2.svg');
        slide.addImage({ path: imagePath2, x: 0.6614173228, y: 4.5275590551, w: 8.3149606299, h: 1.6968503937 });
        slide.addText(
            [
                { text: "References", options: { fontSize: 12, color: "767171", fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bold: true } },
                { text: "https://www.fidelity.com/viewpoints/retirement/tax-savvy-withdrawals", options: { hyperlink: { url: "https://www.fidelity.com/viewpoints/retirement/tax-savvy-withdrawals" }, color: "0070C0", bullet: { type: "number" } }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14 },
                { text: "2 Assumes 5% annual return on all investments", options: { fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bullet: { type: "number" }, color: "767171" } },
            ],
            { x: 0.6614173228, y: 6.4094488189, w: 12.011811024, h: 0.6062992126, isTextBox: true, wordWrap: true }
        )

        let tableData1 = [
            [
                { text: "Conventional Wisdom Withdrawal Strategy", options: { fontSize: 12, fontFace: "Barlow", bold: true, color: "FFFFFF", valign: "middle", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: "Traditionally, tax professionals suggest withdrawing first from taxable accounts, then tax-deferred accounts, and finally Roth accounts where withdrawals are tax-free. The goal is to allow tax-deferred assets to grow longer and faster.", options: { fontSize: 12, fontFace: "Barlow", color: "000000", lineSpacingMultiple: 1.14, valign: "middle", } }
            ],

        ];
        const tableOpts1 = {
            x: 9.2952755906,
            y: 1.4212598425,
            w: 3.3779527559,
            h: 2.1417322835,
            rowH: [0.4, 1.8],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData1, tableOpts1);

        let tableData2 = [
            [
                { text: "Hypothetical Example", options: { fontSize: 12, fontFace: "Barlow", bold: true, color: "FFFFFF", valign: "middle", align: "center", fill: { color: "376C8A" } } },
            ],
            [
                { text: "62-year old single retiree with $200,000 in taxable accounts, $250,000 in tax-deferred, and $50,000 in tax-free. The retiree also receives $25,000 per year in Social Security and has a total after-tax income need of $60,000 per year.", options: { fontSize: 12, fontFace: "Barlow", color: "000000", lineSpacingMultiple: 1.14, valign: "middle", } }
            ],
        ];
        const tableOpts2 = {
            x: 9.2952755906,
            y: 3.7874015748,
            w: 3.3779527559,
            h: 2.1417322835,
            rowH: [0.4, 1.8],
            border: { color: "376C8A" },
        };
        slide.addTable(tableData2, tableOpts2);


        //footer
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 39', error: err.message });
    }
}

const slide40 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        //title
        slide.addText("Net Returns Used for Savings Accounts", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });

        const imagePath1 = path.resolve(__dirname, "../", '../images/slide40/slide40.1.svg');
        const imagePath2 = path.resolve(__dirname, "../", '../images/slide40/slide40.2.svg');
        const imagePath3 = path.resolve(__dirname, "../", '../images/slide40/slide40.3.svg');

        slide.addImage({ path: imagePath1, x: 0.6614173228, y: 1.2755905512, w: 12.011811024, h: 1.9606299213 });
        slide.addImage({ path: imagePath2, x: 0.6614173228, y: 3.4251968504, w: 12.011811024, h: 1.7874015748 });
        slide.addImage({ path: imagePath3, x: 0.6614173228, y: 5.4015748031, w: 12.011811024, h: 1.0866141732 });

        //footer
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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 40', error: err.message });
    }
}

const slide41 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth

        //title
        slide.addText("Ability to Spend", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });



        slide.addText(
            [
                { text: "According to a Consumer Expenditure Survey by the Bureau of Labor Statistics, spending drops 24% from ages 64 – 75 vs after age 75.", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: true, indentLevel: 0 } },
                { text: "https://www.bls.gov/cex/", options: { hyperlink: { url: "https://www.bls.gov/cex/" }, bullet: { code: "2500" }, indentLevel: 1, color: "0070C0", fontSize: 16, fontFace: "Barlow", lineSpacingMultiple: 1.3, } },
                { text: "Your plan stops increasing your target income with inflation at age 81 to more accurately model your ability to spend.", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: true, indentLevel: 0 } },
                { text: "By age 90, this is a 10% spending drop", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: { code: "2500" }, indentLevel: 1 } },
                { text: "By age 100, this is a 20% spending drop (still below the average 24% spending drop in the US population)", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: { code: "2500" }, indentLevel: 1 } },
                { text: "An article in the New York Times that describes your ability to spend as you age:", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: true, indentLevel: 0 } },
                { text: "“Housing costs remained steady and health care expenses increased, but nearly every other category — transportation, entertainment, clothing, food and drink — declined sharply.”", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: { code: "2500" }, indentLevel: 1 } },
                { text: "“Take that cruise when you’re 65 or 70 because you’re probably not going to be able to take it when you’re 80.”", options: { fontSize: 16, color: "767171", fontFace: "Barlow", lineSpacingMultiple: 1.3, color: "000000", bullet: { code: "2500" }, indentLevel: 1 } },
                { text: "https://www.nytimes.com/2018/11/29/business/retirement/retirement-spending-calculators.html ", options: { hyperlink: { url: "https://www.nytimes.com/2018/11/29/business/retirement/retirement-spending-calculators.html" }, bullet: { code: "2500" }, indentLevel: 1, color: "0070C0", fontSize: 16, fontFace: "Barlow", lineSpacingMultiple: 1.3, } },
            ],
            { x: 1.6377952756, y: 1.74409448823, w: 11.039370079, h: 4.3267716535, isTextBox: true, wordWrap: true }
        )
        const imagePath1 = path.resolve(__dirname, "../", '../images/slide41/slide41.1.svg');

        slide.addImage({ path: imagePath1, x: 0.6614173228, y: 3.4330708661, w: 0.8385826772, h: 0.9488188976 });



        //footer

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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });
    } catch (err) {
        res.status(500).json({ message: 'Error in slide 41', error: err.message });
    }
}


const slide42 = async () => {
    let slide = pptx.addSlide()
    const json = await convertXMLtoJS(xmlDataPath)
    const year = (await getYearAndMonth(json)).planYear
    const month = (await getYearAndMonth(json)).updateMonth
    //title
    slide.addText("Appendix – Roth IRA Conversion", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true, bold: true });

    let arrTextObjects1 = [
        { text: "You make contributions with after-tax dollars that can potentially grow tax-free", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14 } },
        { text: "During retirement, your withdrawals are tax-free if certain conditions are met", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14 } },
    ]

    let tableData1 = [
        [
            { text: "What a Roth IRA is", options: { fontSize: 11, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects1 }
        ],

    ];

    const tableOpts1 = {
        x: 0.6614173228,
        y: 1.2519685039,
        w: 6.0078740157,
        h: 1.1732283465,
        rowH: [0.25, 0.75],
        border: { color: "376C8A" },
    };
    slide.addTable(tableData1, tableOpts1);

    let tableData2 = [
        [
            { text: "What a Roth IRA Conversion is", options: { fontSize: 11, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: "A Roth IRA conversion occurs when an account owner transfers or rolls over savings from a traditional IRA or employer sponsored tax-deferred retirement account (e.g. 401(k),  403(b), 457(b) account)", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14 } }
        ],

    ];

    const tableOpts2 = {
        x: 6.9527559055,
        y: 1.2519685039,
        w: 5.7204724409,
        h: 1.1732283465,
        rowH: [0.25, 0.75],
        border: { color: "376C8A" },
    };
    slide.addTable(tableData2, tableOpts2);


    let arrTextObjects3 = [
        { text: "Tax-Free withdrawals if certain conditions are met", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
        { text: "No Required Minimum Distributions", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
        { text: "Assets passed to beneficiaries to grow and be withdrawn tax-free", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
        { text: "In contrast with Tax-deferred assets that require beneficiaries to add any withdrawal of inherited tax-deferred accounts to their income tax calculation in the year of withdrawal", options: { fontSize: 11, bullet: { code: "2500" }, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 1 } },

    ]

    let tableData3 = [
        [
            { text: "Advantages of a Roth IRA", options: { fontSize: 11, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects3 }
        ],

    ];

    const tableOpts3 = {
        x: 0.6614173228,
        y: 2.5275590551,
        w: 6.0078740157,
        h: 1.8307086614,
        rowH: [0.25, 0.75],
        border: { color: "376C8A" },
    };
    slide.addTable(tableData3, tableOpts3);


    let arrTextObjects4 = [
        { text: "The account owner adds the taxable amount of the converted account to their income tax calculation in the year of conversion.", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
        { text: "The more the converted funds grow in the Roth IRA account before being withdrawn the more potential savings account owners can realize from taxes assuming income tax rates do not decline", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
    ]

    let tableData4 = [
        [
            { text: "Considerations when completing a Roth IRA Conversion", options: { fontSize: 11, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects4 }
        ],

    ];

    const tableOpts4 = {
        x: 6.9527559055,
        y: 2.5275590551,
        w: 5.7204724409,
        h: 1.8307086614,
        rowH: [0.25, 1.325],
        border: { color: "376C8A" },
    };
    slide.addTable(tableData4, tableOpts4);



    let arrTextObjects5 = [
        { text: "Roth IRA assets can be withdrawn tax-free, as long as assets are held within a Roth IRA for a 5-year aging period and at least one of the following conditions has been met: ", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
        { text: "The owner is at least 59½ years of age.", options: { fontSize: 11, bullet: { code: "2500" }, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 1 } },
        { text: "A distribution is made to a beneficiary after the death of the original Roth IRA owner.", options: { fontSize: 11, bullet: { code: "2500" }, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 1 } },
        { text: "The distributing Roth IRA owner is disabled under the applicable IRS definition.", options: { fontSize: 11, bullet: { code: "2500" }, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 1 } },
        { text: "The distribution is for a qualifying first-time home-buyer expense, up to a $10,000 lifetime maximum amount, per individual IRA.", options: { fontSize: 11, bullet: { code: "2500" }, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 1 } },
        { text: "If these criteria are not met, a Roth IRA owner will generally be subject to tax, and possibly an additional 10% penalty for early withdrawal. However, withdrawals of contributions from a Roth IRA are generally tax-free.", options: { fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.14, indentLevel: 0 } },
    ]

    let tableData5 = [
        [
            { text: "Roth IRA Tax-Free withdrawal conditions", options: { fontSize: 11, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects5 }
        ],

    ];

    const tableOpts5 = {
        x: 0.6614173228,
        y: 4.3937007874,
        w: 12.011811024,
        h: 2.0669291339,
        rowH: [0.40, 1.325],
        border: { color: "376C8A" },
    };
    slide.addTable(tableData5, tableOpts5);


    slide.addText(
        [
            { text: "References\n", options: { fontSize: 12, color: "767171", fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, bold: true } },
            { text: "https://www.fidelity.com/viewpoints/wealth-management/insights/roth-ira-conversion", options: { hyperlink: { url: "https://www.fidelity.com/viewpoints/wealth-management/insights/roth-ira-conversion" }, color: "0070C0" }, fontSize: 12, fontFace: "Calibri (Body)", lineSpacingMultiple: 1.14, }
        ],
        { x: 0.6614173228, y: 6.6141732283, w: 12.011811024, h: 0.405511811, isTextBox: true, wordWrap: true }

    )

    //footer
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
}

const slide43 = async () => {

    let slide = pptx.addSlide()
    const json = await convertXMLtoJS(xmlDataPath)
    const year = (await getYearAndMonth(json)).planYear
    const month = (await getYearAndMonth(json)).updateMonth
    //title
    slide.addText("How Much Long-Term Care Will You Need?", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true });

    slide.addText(
        [
            { text: "The duration and level of LTC varies for each person and can change over time.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 0 } },
            { text: "Costs of LTC vary widely depending on state of residence and type of care.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1 } },
            { text: "A person turning 65 today has a 70% chance of needing some type of long-term care or support.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1 } },
            { text: "The average length of time people use any form of LTC is three years, although women may need more (average 3.7 years) and men may need less (average 2.2 years).", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 0 } },
            { text: "The average length of time receiving any form of care at home is 2 years.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1 } },
            { text: "The average length of time receiving any form of care in a facility is 3 years.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1 } },
            { text: "20% of people may need up to 5 years of LTC.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1 } },
        ],
        { x: 0.9173228346, y: 1.1023622047, w: 11.5, h: 2.342519685, wordWrap: true, isTextBox: true, lineSpacingMultiple: 1.14 }
    );

    let arrTextObjects1 = [
        { text: "LTC expenses in your plan are based on:", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 0, lineSpacingMultiple: 1.14 } },
        { text: "3 years of care", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1, lineSpacingMultiple: 1.14 } },
        { text: "semi-private room at a nursing home facility", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1, lineSpacingMultiple: 1.14 } },
        { text: "beginning age 85", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1, lineSpacingMultiple: 1.14 } },
        { text: "5% annual inflation for LTC costs", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 1, lineSpacingMultiple: 1.14 } },
        { text: "We deliberately use a conservative estimate to ensure clients can afford as much care as they need.", options: { fontSize: 14, fontFace: "Barlow", bullet: true, color: "000000", indentLevel: 0, lineSpacingMultiple: 1.14 } },

    ]

    let tableData1 = [
        [
            { text: "How we estimate LTC costs", options: { fontSize: 14, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects1 },
        ],
    ];

    const tableOpts1 = {
        x: 0.3661417323,
        y: 3.5511811024,
        w: 12.307086614,
        h: 2.2086614173,
        rowH: 0.25,
        border: { color: "376C8A" },
    };

    slide.addTable(tableData1, tableOpts1);


    let arrTextObjects2 = [
        { text: "Genworth Cost of Care Survey", options: { hyperlink: { url: "https://www.genworth.com/aging-and-you/finances/cost-of-care.html" }, fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.2 } },
        { text: "LongTermCare.gov - How Much Care Will You Need?", options: { hyperlink: { url: "https://acl.gov/ltc/basic-needs/how-much-care-will-you-need" }, fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.2 } },
    ]
    let arrTextObjects3 = [
        { text: "ACHA/NCAL - Facts and Figures", options: { hyperlink: { url: "https://www.ahcancal.org/Assisted-Living/Facts-and-Figures/Pages/default.aspx" }, fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.2 } },
        { text: "NIH Study - Lengths of Stay for Older Adults Residing in Nursing Homes at the End of Life ", options: { hyperlink: { url: "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC2945440/" }, fontSize: 11, bullet: true, fontFace: "Barlow", lineSpacingMultiple: 1.2 } },
    ]
    let tableData2 = [
        [
            { text: "More LTC Resources", options: { fontSize: 14, fontFace: "Barlow", bold: true, color: "FFFFFF", align: "center", colspan: 2, fill: { color: "376C8A" } } },
        ],
        [
            { text: arrTextObjects2 },
            { text: arrTextObjects3 }
        ],

    ];

    const tableOpts2 = {
        x: 0.3622047244,
        y: 5.7992125984,
        w: 12.307086614,
        h: 0.842519685,
        rowH: 0.25,
        border: { color: "376C8A" },
    };

    slide.addTable(tableData2, tableOpts2);




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
}

const slide44 = async () => {
    try {
        let slide = pptx.addSlide()
        const json = await convertXMLtoJS(xmlDataPath)
        const year = (await getYearAndMonth(json)).planYear
        const month = (await getYearAndMonth(json)).updateMonth
        const slideData = await getTextForSlide("slide44");
        const text = slideData.textContent || alternateTextSlide44


        slide.addText("Important Disclosures", { x: 0.6614173228, y: 0.4330708661, w: 12.011811024, h: 0.5393700787, fontSize: 32, color: "376C8A", fontFace: "Barlow Condensed SemiBold", wordWrap: true, isTextBox: true });
        slide.addText(`${text}`, { x: 0.562992126, y: 1.2480314961, w: 12.11023622, h: 5.811023622, fontSize: 8, color: "000000", fontFace: "Arial", isTextBox: true, lineSpacingMultiple: 1.14, margin: { bottom: 6 }, });

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
        slide.addImage({ path: imageLine1, x: 0, y: 0, w: 0.688976378, h: 0.688976378 })
        slide.addImage({ path: imageLine2, x: 0, y: 0, w: 0.8188976378, h: 0.8267716535 });

    } catch (err) {
        res.status(500).json({ message: 'Error in slide 44', error: err.message });
    }
}

const generateAndSavePresentation = async (_, res) => {
    try {
        // Add slides to the presentation
        await slide1();
        await slide2();
        await slide3();
        await slide4();
        await slide5();
        await slide37();
        await slide38();
        await slide39();
        await slide40();
        await slide41();
        await slide42();
        await slide43();
        await slide44();

        // Generate the presentation file in memory
        await pptx.writeFile({ fileName: 'jadoo.pptx' });


        fs.readFile('jadoo.pptx', function (err, data) {
            if (err) {
                throw err;
            }

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
            res.setHeader('Content-Disposition', 'attachment; filename=presentation.pptx');
            res.send(data);
        });
    } catch (error) {
        console.error("Error generating and downloading presentation:", error);
        res.status(500).json({ message: 'Error generating and downloading presentation' });
    }
};

module.exports = { generateAndSavePresentation, slide1, slide2, slide3, slide4 };