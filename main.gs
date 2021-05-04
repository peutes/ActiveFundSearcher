
// hook
function onOpen() {
  const menu = [{name: 'ランキング', functionName: 'scrapingMinkabuRanking'}]
  for (let i=0; i<fundsSheetMax; i++) {
    menu.push({name: 'ファンド' + i, functionName: 'scrapingMinkabuFunds' + i})
  }
  menu.push({name: 'スコア計算', functionName: 'calcMinkabuFundsScore'})
  menu.push({name: 'スコアランキング計算', functionName: 'calcMinkabuFundsScoreRanking'})
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

function scrapingMinkabuFunds12() {
  (new MinkabuFundsScraper(12)).scraping()
}

function scrapingMinkabuFunds13() {
  (new MinkabuFundsScraper(13)).scraping()
}

function scrapingMinkabuFunds14() {
  (new MinkabuFundsScraper(14)).scraping()
}

function scrapingMinkabuFunds15() {
  (new MinkabuFundsScraper(15)).scraping()
}

function scrapingMinkabuFunds16() {
  (new MinkabuFundsScraper(16)).scraping()
}

function scrapingMinkabuFunds17() {
  (new MinkabuFundsScraper(17)).scraping()
}

function scrapingMinkabuFunds18() {
  (new MinkabuFundsScraper(18)).scraping()
}

function scrapingMinkabuFunds19() {
  (new MinkabuFundsScraper(19)).scraping()
}

function scrapingMinkabuFunds20() {
  (new MinkabuFundsScraper(20)).scraping()
}

function calcMinkabuFundsScore() {
  (new MinkabuFundsScoreCalculator).calc()
}

function calcMinkabuFundsScoreRanking() {
  (new MinkabuFundsScoreRanking).calc()
}

