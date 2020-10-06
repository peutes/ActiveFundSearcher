class Fund {
  constructor(link, isIdeco) {
    this.link = link
    this.isIdeco = isIdeco
    
    this.date = null
    this.category = null
    this.isRakuten = null
    this.rate = null
    this.name = null
    this.ignore = false
    this.returns = new Array(minkabuTermSize).fill(null)
    this.risks = new Array(minkabuTermSize).fill(null)
    this.sharps = new Array(minkabuTermSize).fill(null)

    this.scores = new Array(scoresSize)
    for (let i=0; i<scoresSize; i++) {
      this.scores[i] = new Array(minkabuTermSize).fill(null)
    }
    this.totalScores = new Array(scoresSize).fill(0)
  }
}

class MinkabuSheetInfo {
  constructor() {
    this.linkSheetName = 'Link'
    this.fundsSheetNames = [];
    for (let i=0; i<fundsSheetMax; i++) {
      this.fundsSheetNames.push('Funds' + i)
    }
    this.downsideRiskSheetName = 'Downside'
    this.infoSheetName = 'Info'
    this.logSheetName = 'Log'
    
    this.minkabuSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  getScoreSheetName() {
    const prefix = 'スコア：'
    return prefix + (new Date().toLocaleString("ja")) 
  }
  
  getSheet(name) {
    return this.minkabuSpreadSheet.getSheetByName(name)
  }

  insertSheet(name) {
    return this.minkabuSpreadSheet.insertSheet(name, 0)
  }
}

class MinkabuRankingScraper {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo()
    this._funds = new Map()
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
    const baseLink = 'https://itf.minkabu.jp'
    const pageNum = 190 // とりあえず3800件まで対応。

    let i=0;
    for (let p=1; p<=pageNum; p++) {
      const link = baseLink + '/ranking/return?page=' + p
      const html = UrlFetchApp.fetch(link).getContentText()
      const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
      const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
      if (aList.length === 0 || aList[0] === '<!DOCTYPE htm') {
        console.log("ERROR:", p, table, aList)
        return
      }
    
      aList.forEach(a => {
        const result = a.match(/href="(.*)">(.*)/)
        const pass = result[1]
        const link = baseLink + pass + '/risk_cont'
        this._funds.set(pass, new Fund(link, false))
        i++
      })
      
      if (p%20 === 0) {
        console.log(p, this._funds.size, i)
      }
    }
    console.log('_fetchFunds', this._funds.size, i)
  }
  
  _getIdecoFunds() {
    const ids = ['48315184']
    ids.forEach(id => {
      const link = 'https://itf.minkabu.jp/fund/' + id + '/risk_cont'
      this._funds.set('/fund/' + id, new Fund(link, false))
    })
    
    // SBI Select版
    const idecoIds = [
      '2931316B', '03317172', '9C31116A', '0131C18A', '04316186', '89311164', '8931217C', '03316183', '89313135', '96312073', 
      '03319172', '2931113C', '03311187', '0431Q169', '0231202C', '25311177', '65311058', '68311003', '0331C177', '8931111A', 
      '03318172', '0331A172', '0231402C', 'AN31118A', '0431U169', '29314136', '79314169', '03312175', '04316188', '8931118A',
      '96311073', '03311112', '89312121', '89313121', '89314121', '89315121',
    ]
    
    idecoIds.forEach(id => {
      const link = 'https://itf.minkabu.jp/fund/' + id + '/risk_cont'
      this._funds.set('/fund/' + id, new Fund(link, true))
    })
    console.log('_getIdecoFunds', this._funds.size)
  }
}

class MinkabuFundsScraper {
  constructor(fundsSheetNum) {
    this._fundsSheetNum = fundsSheetNum
    this._sheetInfo = new MinkabuSheetInfo()
    this._fundSheetName = this._sheetInfo.fundsSheetNames[fundsSheetNum]
    this._funds = new Map()
  }

  scraping() {
    this._fetchLinks()
    this._fetchDetail()
    this._output()
  }

  _fetchLinks() {
    const sheet = this._sheetInfo.getSheet(this._sheetInfo.linkSheetName)
    const values = sheet.getDataRange().getValues()
    const n = 300
    const end = Math.min(values.length, n * (this._fundsSheetNum + 1))
    for (let i=n * this._fundsSheetNum; i<end; i++) {
      this._funds.set(values[i][0], new Fund(values[i][0], values[i][1]))
    }
    console.log('_fetchLinks')
  }

  _fetchDetail() {
    console.log('getDetailFromMinkabu:' + this._funds.size)
    
    let i = 0
    this._funds.forEach(fund => {
      const html = UrlFetchApp.fetch(fund.link).getContentText()
      const table = Parser.data(html).from('<table class="md_table">').to('</table>').build()
      const spanList = Parser.data(table).from('<span>').to('</span>').iterate()
    
      // 3ヶ月のデータはスキップする
      fund.name = this._toHankaku(Parser.data(html).from('<p class="stock_name">').to('</p>').build())
      for (let i=1; i<minkabuTermSize + 1; i++) {
        const result1 = spanList[i].replace(/%/, '')
        const result2 = spanList[(minkabuTermSize + 1) + i].replace(/%/, '')
        const result3 = spanList[2 * (minkabuTermSize + 1) + i].replace(/%/, '')
        fund.returns[i - 1] = result1 != '-' ? Number(result1) : null
        fund.risks[i - 1] = result2 != '-' ? Number(result2) : null
        fund.sharps[i - 1] = result3 != '-' ? Number(result3) : null
      }
      fund.date = Parser.data(html).from('<span class="fsm">（').to('）</span>').build()
    
      i++
      if (i%50 === 0) {
        console.log(i)
      }
    })
    console.log('_fetchDetail')
  }

  _output() {
    const sheet = this._sheetInfo.getSheet(this._fundSheetName)
    sheet.clear()

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.date, fund.link, fund.isIdeco, fund.name, '']
      fund.returns.forEach((r, i) => {
        row.push(r, fund.risks[i], fund.sharps[i], '')
      })
      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    sheet.autoResizeColumns(2, 4)
    console.log('_output')
  }
  
  _toHankaku(str) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９！-～]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    }).replace(/”/g, "\"").replace(/’/g, "'").replace(/‘/g, "`").replace(/￥/g, "\\").replace(/　/g, " ").replace(/〜/g, "~")
  }
}

class MinkabuInfoScraper {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo()
  }

  scraping(fund) {
    if (fund.category !== null) {
      return
    }
    
    const link = /(.*)\/.*/.exec(fund.link)[1]
    const html = UrlFetchApp.fetch(link, {muteHttpExceptions:true}).getContentText()
    const tables = Parser.data(html).from('<table class="md_table md_table_vertical">').to('</table>').iterate()
    const tds = Parser.data(tables[1]).from('<td').to('</td>').iterate()
      .map(t => t.replace(' colspan="3"', '').replace('>', ''))

    fund.category = tds[1] // + '-' + tds[3
    
    const tds2 = Parser.data(tables[2]).from('<span class="gold_star">').to('</span>').iterate()
    const count = (4 - (tables[2].match(/-/g) || []).length)
    fund.isRakuten = count === 0 ? 0 : (tables[2].match(/★/g) || []).length / count
    
//    const link2 = link + '/sales_company'
//    const html2 = UrlFetchApp.fetch(link2).getContentText()
//    const table2 = Parser.data(html2).from('<table class="md_table">').to('</table>').build()
//    fund.isRakuten = table2.indexOf("楽天証券") !== -1
  }
}
