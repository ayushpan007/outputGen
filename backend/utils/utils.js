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
    minimumFractionDigits: 0,
    useGrouping: true
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




const alternateTextSlide44 = "This page is the ‘Disclosures page' and must be included with all presentations made to the client.\n\nASSUMPTIONS – This plan is intended to provide an analysis of your financial position and potential income in retirement. This plan incorporates the information provided by you, the client, with respect to your \nincome, expenses and asset holdings. Income plans can offer one or more of the following characteristics: lifetime guarantees, flexibility, principal preservation and growth potential. Our goal is to help you build a \nplan that takes these needs into account, given your preferences, goals and objectives.\n\nThe plan recommendations are based on your current situation, your resources, and your goals. In addition, they are based on our current expectations of the behavior of the accounts and products being\nrecommended. This is a hypothetical example only and is not intended to predict the actual performance of any specific product. The returns have been shown to continue unchanged for all years of the plan but\nthis is not likely to occur and actual results may be more or less favorable. All investments have risks associated with them and future loss is possible.\n\nCRITERIA AND METHODOLOGY – The income plan may contain investment accounts, annuities, life policies, incomes like pensions and social security, income tax estimates, and detailed development of your\nretirement annual income target. The objective is to give you and your advisor the ability to show how these various pieces of a retirement puzzle can be brought together and structured to optimize income,\nminimize taxes and provide effective wealth transfer. The real power of the tool is creating scenarios which can be tested to see how these elements may be impacted under different conditions or using different\nplanning concepts. The hypothetical variables include account growth, inflation, tax obligation, and the desired annual income target. Changing any of them will greatly impact the plan results.\n\nFor assets allocated to investment accounts, growth will be estimated using an average fixed rate which is hypothetical and not meant to indicate historical or future results. The plan may also show income\ndistributions representing the amount of money to be withdrawn from the account. These income dollars may or may not be guaranteed and are subject to change. This illustrated income could represent a\ndistribution of principal and/or interest depending on investment performance. The growth rates illustrated on this proposed income plan are for illustrative purposes only and are not guaranteed. These rates will\nchange on a daily basis and also could be negative. Past performance is not an indication of future results.\nFor assets allocated to insurance contracts, the contract and any guarantees therein are subject to the claims paying ability of the carrier. Annuity projected growth rates may show income benefit base growth\nand not the market value of the annuity. Annuity distributions may be subject to withdrawal charges, premium bonus recapture charges and market value adjustments (if applicable) and may result in a loss of\nprincipal. Insurance company product recommendations must be accompanied by approved illustrations and/or brochures. Other investment recommendations must be accompanied by an approved prospectus.\n\nIf there are any insurance products or annuities within the plan presentation, the National Association of Insurance Commissioners has specifically required that the consumer be given an illustration disclosing all\naspects of how that product works and what the minimum guarantees are. This plan does not generate the required illustration and that must be furnished separately. All Income projections are hypothetical and\nshould not be considered indicative of actual income. The income portion of this analysis does not take into account any taxes unless otherwise noted.\n\nLIMITATIONS AND RISKS - The information contained in this report is not guaranteed to be accurate, complete or timely. Neither your advisor nor anyone who helped your advisor create or populate this report,\nincluding, but not limited to, any software or information provider, shall be liable for any damages or losses related to your use of the information contained in it. The information contained in the plan is to be used\nfor informational purposes only. The income plan does not provide tax advice. The tax calculations and tax projections shown in this plan are approximate and not intended to be accurate. An appropriate tax\nprofessional should be consulted prior to implementation of any strategy. The information provided in the plan is not intended to be used, nor can it be used for the purpose of avoiding U.S. Federal, state, or local\ntax penalties. Potential Social Security Benefits shown in the plan are for informational purposes only. Potential Cost of Living increases are shown at a fixed rate. This is not likely to happen. Actual Social\nSecurity Benefits may be impacted by a number of different factors related to your personal situation. You should refer to the Social Security Administration for information on your future benefit. We are not\naffiliated with the Social Security Administration or any other government agency. \n\nImportant: This report is a hypothetical illustration based on information provided by you the client with respect to your income, expenses, and asset holdings. The assumptions regarding investment returns, contract growth, cost of living increases, and/or inflation are hypothetical in nature, do not reflect actual investment results, and are not guarantees of future results and should be carefully considered. If any specific investment or contract is included in the plan, it must be accompanied by separate appropriate disclosures. This report is not complete without all pages."

const getTextForSlide = async (slideName) => {
  try {
    const slide = await SlideText.findOne({ slideName });
    if (!slide) {
      return { error: 'Slide not found' };
    }
    return slide;
  } catch (error) {
    return { error: error.message };
  }
};


module.exports = { getYearAndMonth, formatCurrency, querySIPS, generateAgeArray, replaceNegativeWithZero, isSlideValid, convertXMLtoJS, alternateTextSlide44, getTextForSlide };
