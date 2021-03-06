
// hook
function onOpen() {
  const menu = [{name: 'ランキング', functionName: 'scrapingMinkabuRanking'}]
  for (let i=0; i<fundsSheetMax; i++) {
    menu.push({name: 'ファンド' + i, functionName: 'scrapingMinkabuFunds' + i})
  }
  menu.push({name: 'ファンドスコア', functionName: 'calcMinkabuFundsScore'})
  SpreadsheetApp.getActiveSpreadsheet().addMenu("みんかぶ", menu)
}

function scrapingMinkabuRanking() {
  (new MinkabuRankingScraper()).scraping()
}
 
function scrapingMinkabuFunds0() {
  (new MinkabuFundsScraper(0)).scraping()
}

function scrapingMinkabuFunds1() {
  (new MinkabuFundsScraper(1)).scraping()
}

function scrapingMinkabuFunds2() {
  (new MinkabuFundsScraper(2)).scraping()
}

function scrapingMinkabuFunds3() {
  (new MinkabuFundsScraper(3)).scraping()
}

function scrapingMinkabuFunds4() {
  (new MinkabuFundsScraper(4)).scraping()
}

function scrapingMinkabuFunds5() {
  (new MinkabuFundsScraper(5)).scraping()
}

function scrapingMinkabuFunds6() {
  (new MinkabuFundsScraper(6)).scraping()
}

function scrapingMinkabuFunds7() {
  (new MinkabuFundsScraper(7)).scraping()
}

function scrapingMinkabuFunds8() {
  (new MinkabuFundsScraper(8)).scraping()
}

function scrapingMinkabuFunds9() {
  (new MinkabuFundsScraper(9)).scraping()
}

function scrapingMinkabuFunds10() {
  (new MinkabuFundsScraper(10)).scraping()
}

function scrapingMinkabuFunds11() {
  (new MinkabuFundsScraper(11)).scraping()
}

function calcMinkabuFundsScore() {
  (new MinkabuFundsScoreCalculator).calc()
}

