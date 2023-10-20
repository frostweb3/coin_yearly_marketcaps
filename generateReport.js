const { readFileSync } = require("fs");
const XLSX = require("xlsx");
const PER_REQUEST_DELAY_MS = 2000; // 2 seconds

async function writeDailyData(sheet, date, colIdx, modeSymbolRender) {
  const justDate = date.toISOString().split('T')[0];
  var apiUrl = `https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/historical?date=${justDate}&limit=200&sort=market_cap&sort_dir=desc`;
  var apiKey = "";
  console.log(apiUrl);

   XLSX.utils.sheet_add_aoa(sheet, [[justDate.toString()]], { origin: { r: 0, c: colIdx } });
// console.log(sheet);

   return fetch(
    apiUrl, {
      headers: {
        "X-CMC_PRO_API_KEY": apiKey
      }
    }
  ).then((resp) => resp.json()).then((content) => {
    const { data } = content;
    // if(colIdx == 5) {
    //   console.log(content.data.slice(0, 10).map((e) => e.quote.USD));
    // }


    // In case the API returns something that does not have the format we want
    if (!content.data) {
      console.log(content);
    }

    data.forEach((coin, rowIdx) => {
      const { symbol } = coin;

      if (modeSymbolRender) {
        XLSX.utils.sheet_add_aoa(sheet, [[symbol]], { origin: { r: rowIdx + 1, c: colIdx - 1  } });
        return;
      }

      XLSX.utils.sheet_add_aoa(sheet, [[coin.quote.USD.market_cap]], { origin: { r: rowIdx + 1, c: colIdx } });
    });
  });
}

async function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms))
}

async function main() {
  const { read } = XLSX;

  const buf = readFileSync("out.xlsx");

  const workbook = read(buf);
  const sheet = workbook.Sheets["Marketcaps"];

  // Render only symbol column
  let yesterday = new Date();

  yesterday = new Date(yesterday.setDate(yesterday.getDate() - 1));
  await writeDailyData(sheet, yesterday, 1, true);

  for (var i = 1; i <= 30; i++) {
    // Subtract i days
    let currentDate = new Date();
    currentDate = new Date(currentDate.setDate(currentDate.getDate() - i));

    console.log(`Processing ${currentDate.toISOString()}...`);
    await writeDailyData(sheet, currentDate, i, false);
    await wait(PER_REQUEST_DELAY_MS);
  }

//   console.log(sheet);
  XLSX.writeFile(workbook, "out.xlsx");
}
main();
