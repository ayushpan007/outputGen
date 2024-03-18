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
        await slide3();

        await pptx.writeFile({ fileName: filePath });
        res.status(200).json({ message: 'Presentation created and saved successfully' });
    } catch (error) {
        console.error("Error saving presentation:", error);
        res.status(500).json({ message: 'Error saving presentation' });
    }
};
module.exports = { generateAndSavePresentation };