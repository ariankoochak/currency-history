const axios = require("axios");
const ExcelJS = require("exceljs");

const fetchExchangeRates = async (baseUrl, startDate, endDate) => {
    const rates = [];
    let currentDate = new Date(startDate);

    while (currentDate <= new Date(endDate)) {
        console.log(currentDate);
        const dateStr = currentDate.toISOString().split("T")[0];
        try {
            const response = await axios.get(baseUrl);
            if (
                response.data &&
                response.data.conversion_rates &&
                response.data.conversion_rates.IRR
            ) {
                rates.push({
                    Date: dateStr,
                    Rate: response.data.conversion_rates.IRR,
                });
            }
        } catch (error) {
            console.log(`Error fetching data for ${dateStr}:`, error.message);
        }

        currentDate.setDate(currentDate.getDate() + 1);
    }

    return rates;
};

const saveToExcel = async (data, fileName) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Exchange Rates");
    worksheet.columns = [
        { header: "Date", key: "Date", width: 15 },
        { header: "Rate", key: "Rate", width: 15 },
    ];
    data.forEach((rate) => {
        worksheet.addRow(rate);
    });
    await workbook.xlsx.writeFile(fileName);
    console.log(`File saved as ${fileName}`);
};

(async () => {
    const baseUrl =
        "https://v6.exchangerate-api.com/v6/<APIKEY>/latest/USD";
    const startDate = "2020-01-01";
    const endDate = "2025-01-10";
    const fileName = "exchange_rates.xlsx";

    console.log("Fetching exchange rates...");
    const rates = await fetchExchangeRates(baseUrl, startDate, endDate);

    console.log("Saving data to Excel...");
    await saveToExcel(rates, fileName);

    console.log("Done!");
})();
