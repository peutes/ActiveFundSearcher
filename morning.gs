
class MorningRankingScraper {
  constructor() {
    this._sheetInfo = new SheetInfo()
    this._funds = new Map()
    this.baseLink = 'http://www.morningstar.co.jp/FundData/'
  }
  
  scraping() {
    this._fetchFunds()
    this._getIdecoFunds()

    const data = []
    this._funds.forEach(fund => {
      data.push([fund.link, fund.isIdeco])
    })
  
    const sheet = this._sheetInfo.getSheet(this._sheetInfo.linkSheetName)
    sheet.clear()
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
  }

  _fetchFunds() {
    const rankingLink = this.baseLink + 'DetailSearchResult.do?mode=2'
//    const rankingLink = this.baseLink + 'DetailSearchResult.do?mode=2&selectedDcSmaEtfSearchKbn=on' // DCとETF系を基本除外 あまり多すぎて辛かったらこっちにする
    const fundReturnLink = this.baseLink + 'Return.do?fnc='
    
    const html0 = UrlFetchApp.fetch(rankingLink).getContentText('Shift-JIS')
    const count = Parser.data(html0).from('<span class="lltxt fcred"><span class="plus">').to('</span></span>').build()      
    const pageNum = Math.ceil(count / 50)
    
    for (let p=0; p<pageNum; p++) {
      const link = rankingLink + '&pageNo=' + (p + 1)
      const html = UrlFetchApp.fetch(link).getContentText('Shift-JIS')
      const table = Parser.data(html).from('<table class="table1f">').to("</table>").build()
      Parser.data(table).from('<td>').to("</td>").iterate()
        .map(td => /<a href="SnapShot\.do\?fnc=(.*)" target=.*<\/a>/g.exec(td))
        .filter(Boolean)
        .forEach(a => this._funds.set(a[1], new Fund(fundReturnLink + a[1], false)))

      if (p % 10 === 0) {
        console.log(p, pageNum, this._funds.size)
      }
    }

    console.log('_fetchFunds', this._funds.size)
  }
  
  _getIdecoFunds() {
    const idecoOriginalIds = [
      "2014120201", "2004022705", "2016042102", "1998040101", "2000082201", "2016100304", "201609300C", "1998102201", "2002052801", "200011300A", "1998082101", "2004012801",
      "2016102105", "2015042701", "2011120901", "2016042103", "2005083106", "2000032406", "2013051303", "2013051301", "2017070502", "201609080E", "2002012507", "201503310B",
      "2002121002", "2006013101", "2011102805", "2013051302", "2009121101", "2002012501", "2012102901", "2002093001", "2016063005", "2013051304", "2002121004", "2002110801",
      "2002040101", "2010102907", "2012022801", "2001113011", "2002121005", "2002121006", "2002121007", "2002121008", "2016090815", "2008010901", "2016033001", "2016033002",
      "2016033003", "2016033004", "2016033005", "2012012302", "2012012303", "2012012304", "2012012305", "2004123001", "2013051305", "2007081301", "2016102106", "2016092308",
      "2008071603", "2011020701", "2008062701"
    ]
    const idecoSelectIds = [
      "2016100304", "2018061103", "2016042102", "2018103104", "2017022705", "2016112105", "2013051303", "2007031506", "2017070502", "2005083106", "2000032406", "2017120601",
      "2018031901", "2017022703", "2013121001", "2018070301", "201609080E", "2002121002", "2011102805", "2017073108", "2017022704", "2018100401", "2017022702", "2002121004",
      "2016090812", "2007031505", "2017050902", "2018083105", "2018100402", "2012012302", "2012012303", "2012012304", "2012012305", "201306280D", "2016092308", "2011020701"
    ]
    const idecoIds = idecoOriginalIds.concat(idecoSelectIds)
    
    const fundReturnLink = this.baseLink + 'Return.do?fnc='
    idecoIds.forEach(id => this._funds.set(id, new Fund(fundReturnLink + id, true)))

    console.log('_getIdecoFunds', this._funds.size)
  }
}


function scrapingMorningRanking() {
  (new MorningRankingScraper()).scraping()
}
