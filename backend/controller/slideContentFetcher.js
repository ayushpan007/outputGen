const { convertXMLtoJS, querySIPS, getYearAndMonth, formatCurrency } = require("../utils/utils")

const dataForSlide5 = async (xmlDataPath, res) => {

  const getSavingsAssets = (monetaryAssets) => {
    return monetaryAssets.map((asset) => ({
      description: asset.description,
      balance: formatCurrency(asset.balance),
    }));
  };

  const calculateTotalSavings = (savingsAssets) => {
    return formatCurrency(savingsAssets.reduce((total, asset) => {
      const balanceValue = parseFloat(asset.balance.replace(/[^\d.-]/g, "")) || 0;
      return total + balanceValue;
    }, 0));
  };

  const calculateTaxValues = (monetaryAssets) => {
    let sumTaxDeferred = 0;
    let sumTaxableSaving = 0;
    let sumTaxFreeRoth = 0;
    monetaryAssets.forEach((asset) => {
      const { description, balance } = asset;
      const balanceValue = parseFloat(balance);

      if (
        description === "Fidelity 401k" ||
        description === "Fidelity Rollover IRA"
      ) {
        sumTaxDeferred += balanceValue;
      } else if (description === "Individual TOD") {
        sumTaxableSaving += balanceValue;
      }
    });

    return {
      sumTaxDeferred: formatCurrency(sumTaxDeferred),
      sumTaxableSaving: formatCurrency(sumTaxableSaving),
      sumTaxFreeRoth: formatCurrency(sumTaxFreeRoth),
    };
  };

  const getLiabilities = (debts) => {
    return debts.map((debt) => ({
      initialAmount: debt.initialAmount,
      title: debt.title,
      rate: debt.expectedReturnPercent,
    }));
  };

  try {
    const json = await convertXMLtoJS(xmlDataPath);
    const monetaryAssets = (await querySIPS("case.monetaryAsset", json)) || [];
    const savingsAssets = getSavingsAssets(monetaryAssets);
    const totalSavings = calculateTotalSavings(savingsAssets);

    const { sumTaxDeferred, sumTaxableSaving, sumTaxFreeRoth } = calculateTaxValues(monetaryAssets);

    const scenarioSc09 = await querySIPS(
      "case.scenario[title='SC09 Income Planning Scenario']",
      json
    );
    const debts = await querySIPS(
      "account[$number(initialAmount)<0][]",
      scenarioSc09
    );

    const scenarioSc01 =
      (await querySIPS("case.scenario[title='SC01 As-Is Scenario']", json)) ||
      (await querySIPS("case.scenario[title='SC01 As Is Scenario']", json));

    let totalIncome = scenarioSc01.target.initialAmount || 0;
    totalIncome = formatCurrency(totalIncome);

    let c1Income =
      (await querySIPS("case.client.client1WagesCurrent", json)) || 0;
    c1Income = formatCurrency(c1Income);

    let c2Income =
      (await querySIPS("case.client.client2WagesCurrent", json)) || 0;
    c2Income = formatCurrency(c2Income);

    let c1SSAmt =
      (await querySIPS("case.client.client1SSProjectedAmount62", json)) || 0;
    c1SSAmt = formatCurrency(c1SSAmt);

    let c2SSAmt =
      (await querySIPS("case.client.client2SSProjectedAmount62", json)) || 0;
    c2SSAmt = formatCurrency(c2SSAmt);

    let c1SSAmtCurrently =
      (await querySIPS("case.client.client1SSCurrentAmount", json)) || 0;
    c1SSAmtCurrently = formatCurrency(c1SSAmtCurrently);

    let c2SSAmtCurrently =
      (await querySIPS("case.client.client2SSCurrentAmount", json)) || 0;
    c2SSAmtCurrently = formatCurrency(c2SSAmtCurrently);

    let c1PensionAmt =
      (await querySIPS("case.client.client1PensionCurrentAmount", json)) || 0;
    c1PensionAmt = formatCurrency(c1PensionAmt);

    let c2PensionAmt =
      (await querySIPS("case.client.client2PensionCurrentAmount", json)) || 0;
    c2PensionAmt = formatCurrency(c2PensionAmt);

    const discoveryScenario = await querySIPS(
      "case.scenario[title='Discovery Scenario']"
    );
    const marginalTaxRate = discoveryScenario?.year?.[0]?.taxRate ?? 0;
    const federalTaxRate = 0;
    let liabilities = [];
    let debtTotal = 0;

    try {
      liabilities = getLiabilities(debts);
      debtTotal = liabilities.reduce((total, debt) => total + parseInt(debt.initialAmount) || 0, 0);
    } catch {
      console.log("No liabilities found in this extract.");
    }

    const beneficiaries = await querySIPS("case.beneficiary", json) || [];
    const beneficiaryNames = beneficiaries.map(beneficiary => beneficiary.name);

    return {
      c1FirstName: await querySIPS("case.client.client1FirstName", json),
      c1LastName: await querySIPS("case.client.client1LastName", json),
      c2FirstName: await querySIPS("case.client.client2FirstName", json),
      c2LastName: await querySIPS("case.client.client2LastName", json),
      c1Age: (await querySIPS("case.client.client1InitialPlanAge", json)) || 0,
      c2Age: (await querySIPS("case.client.client2InitialPlanAge", json)) || 0,
      dependents: ["None"],
      beneficiaries: beneficiaryNames,
      c1RetAge: (await querySIPS("case.client.client1RetireAge", json)) || 0,
      c2RetAge: (await querySIPS("case.client.client2RetireAge", json)) || 0,

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
      liabilitiesTotal: debtTotal,
      totalIncome,

      fedRate: await querySIPS("tax.expectedReturnPercent", scenarioSc09),
      stateRes: await querySIPS("case.client.address2", json),
      year: (await getYearAndMonth(json)).planYear,
      month: (await getYearAndMonth(json)).updateMonth,
    };
  } catch (err) {
    console.log("Error processing XML for slide 5:", err);
    res.status(500).json({ message: 'Error in fetching data for slide 5', error: err.message });
  }
};

module.exports = { dataForSlide5 }