const jsonata = require("jsonata");
const path = require("path");
const fs = require("fs");
var xml2js = require("xml2js");

async function querySIPS(expression, data) {
  try {
    const exp = jsonata(expression);
    return await exp.evaluate(data);
  } catch (error) {
    throw error;
  }
}

function convertXMLtoJS(xmlPath) {
  const xml = xmlPath;
  var parser = new xml2js.Parser({ explicitArray: false });

  return new Promise((resolve, reject) => {
    fs.readFile(xml, (err, data) => {
      if (err) {
        reject(err);
      } else {
        parser.parseString(data, (err, result) => {
          if (err) {
            reject(err);
          }
          resolve(result);
        });
      }
    });
  });
}

async function getYearAndMonth(json) {
  try {
    const updateDate = await querySIPS("case.client.updateDate", json);
    const [monthName, day, year] = updateDate.split("/").map(Number);
    const updateDateObject = new Date(year, monthName - 1, day);
    const updateMonth = updateDateObject.toLocaleDateString("en-US", {
      month: "long",
    });
    const planYear = updateDateObject.getFullYear();
    return {
      updateMonth,
      planYear,
    };
  } catch (error) {
    console.error("Error in function getYearAndMonth:", error);
    throw error;
  }
}
function formatCurrency(value) {
  const formattedValue = Math.abs(value).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 0
  });
  return value < 0 ? "-" + formattedValue : formattedValue;
}

const generateAgeArray = (client1Age, client2Age) => {
  const startAge = Math.max(
    parseInt(client1Age) || 0,
    parseInt(client2Age) || 0
  );
  const ageArray = [];
  for (let age = startAge; age <= startAge + 40; age++) {
    ageArray.push(age);
  }
  return ageArray;
};

const replaceNegativeWithZero = async (array) => {
  if (!Array.isArray(array)) {
    throw new Error('Input is not an array');
  }
  const result = array.map(value => Math.max(value, 0));

  return result;
};
let pageNumber = 0;

const isSlideValid = async (slide, xmlDataPath) => {
  const json = await convertXMLtoJS(xmlDataPath);
  let scenarioSc09, scenarioSc10, scenarioPP1501, scenarioSc0616;
  switch (slide.name) {
    case "slide1":
      pageNumber++;
      return { pageNumber, isValid: true };;
    case "slide2":
      scenarioSc09 = await querySIPS("case.scenario[title='SC09 Income Planning Scenario']", json);
      if (scenarioSc09) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide3":
      scenarioSc10 = await querySIPS("case.scenario[title='SC10 Estate Planning Scenario']", json);
      if (scenarioSc10) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide4":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide5":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide6":
      scenarioSc09 = await querySIPS("case.scenario[title='SC09 Income Planning Scenario']", json);
      if (scenarioSc09) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide7":
      scenarioSc10 = await querySIPS("case.scenario[title='SC10 Estate Planning Scenario']", json);
      if (scenarioSc10) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide8":
      const scenarioSc02 = (await querySIPS("case.scenario[title='SC02 Stress Test']", json)) || (await querySIPS("case.scenario[title='SC02 Stress Test Scenario']", json));
      if (scenarioSc02) {
        pageNumber++;
        return { pageNumber, isValid: true };;
      }
      return { isValid: false };
    case "slide9":
      const scenarioSc04 = await querySIPS("case.scenario[title='SC04 Optimized Scenario']", json);
      if (scenarioSc04) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide10":
      const scenarioSc03 = await querySIPS("case.scenario[title='SC03 Income Longevity Scenario']", json);
      if (scenarioSc03) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide11":
      const scenarioSc05 = (await querySIPS("case.scenario[title='SC05 Bucket Plan Scenario']", json)) || (await querySIPS("case.scenario[title='SC05 Bucket Planning Scenario']", json));
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide12":
      const scenarioSc0601 = await querySIPS("case.scenario[title='SC06.01 Roth IRA Conversion']", json);
      if (scenarioSc0601) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide13":
      const scenarioSc0602 = await querySIPS("case.scenario[title='SC06.02 Major Life Event']", json);
      if (scenarioSc0602) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide14":
      const scenarioSc0603 = await querySIPS("case.scenario[title='SC06.03 Charitable Giving']", json);
      if (scenarioSc0603) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide15":
      const scenarioSc0604 = await querySIPS("case.scenario[title='SC06.04 Build Savings']", json);
      if (scenarioSc0604) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide16":
      const scenarioSc0605 = await querySIPS("case.scenario[title='SC06.05 Sale of Property']", json);
      if (scenarioSc0605) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide17":
      const scenarioSc0606 = await querySIPS("case.scenario[title='SC06.06 Debt Payoff']", json);
      if (scenarioSc0606) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide18":
      const scenarioSc0607 = await querySIPS("case.scenario[title='SC06.07 Healthcare Expenses']", json);
      if (scenarioSc0607) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide19":
      const scenarioSc0608 = await querySIPS("case.scenario[title='SC06.08 Education']", json);
      if (scenarioSc0608) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide20":
      const scenarioSc0609 = await querySIPS(
        "case.scenario[title='SC06.09 Travelling/Vacations']",
        json
      );
      if (scenarioSc0609) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide21":
      const scenarioSc0610 = await querySIPS(
        "case.scenario[title='SC06.10 Major Life Event']",
        json
      );
      if (scenarioSc0610) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide22":
      const scenarioSc0611 = await querySIPS(
        "case.scenario[title='SC06.11 Buying a Vehicle']",
        json
      );
      if (scenarioSc0611) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide23":
      const scenarioSc0612 = await querySIPS(
        "case.scenario[title='SC06.12 Purchase of Property']",
        json
      );
      if (scenarioSc0612) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide24":
      const scenarioSc0613 = await querySIPS(
        "case.scenario[title='SC06.13 Early Retirement']",
        json
      );
      if (scenarioSc0613) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide25":
      const scenarioSc0614 = await querySIPS(
        "case.scenario[title='SC06.14 Leaving an Inheritance']",
        json
      );
      if (scenarioSc0614) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide26":
      scenarioPP1501 = await querySIPS(
        "case.scenario[title='PP15.1 Estimate Cost of Long-Term Care']",
        json
      );
      if (scenarioPP1501) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide27":
      scenarioPP1501 = await querySIPS(
        "case.scenario[title='PP15.1 Estimate Cost of Long-Term Care']",
        json) || await querySIPS("case.scenario[title='PP15.1 Estimate Cost of Long Term Care']", json);
      if (scenarioPP1501) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide28":
      scenarioSc0616 = await querySIPS(
        "case.scenario[title='SC06.16 Survivorship']",
        json
      );
      if (scenarioSc0616) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide29":
      scenarioSc0616 = await querySIPS(
        "case.scenario[title='SC06.16 Survivorship']",
        json
      );
      if (scenarioSc0616) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide30":
      const scenarioSc061601 = await querySIPS(
        "case.scenario[title='SC06.16.1 Survivorship w/ Life Insurance']",
        json
      );
      if (scenarioSc061601) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide31":
      const sc09ORsc10 = await querySIPS(
        "case.scenario[title='SC09 Income Planning Scenario']",
        json
      ) || await querySIPS(
        "case.scenario[title='SC10 Estate Planning Scenario']",
        json
      );
      if (sc09ORsc10) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide32":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide33":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide34":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide35":
      const sc09ORsc04 = await querySIPS(
        "case.scenario[title='SC04 Optimized Scenario']",
        json
      ) && await querySIPS(
        "case.scenario[title='SC09 Income Planning Scenario']",
        json
      )
      if (sc09ORsc04) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide36":
      const sc02ORsc10 = await querySIPS(
        "case.scenario[title='SC10 Estate Planning Scenario']",
        json
      ) && await querySIPS(
        "case.scenario[title='SC02 Stress Test']",
        json
      );
      if (sc02ORsc10) {
        pageNumber++;
        return { pageNumber, isValid: true };
      }
      return { isValid: false }
    case "slide37":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide38":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide39":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide40":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide41":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide42":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide43":
      pageNumber++;
      return { pageNumber, isValid: true };
    case "slide44":
      pageNumber++;
      return { pageNumber, isValid: true };
    default:
      return { isValid: false };
  }
}





module.exports = { getYearAndMonth, formatCurrency, querySIPS, generateAgeArray, replaceNegativeWithZero, isSlideValid, convertXMLtoJS };
