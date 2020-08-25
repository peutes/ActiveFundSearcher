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
    this.returns = new Array(termSize).fill(null)
    this.risks = new Array(termSize).fill(null)
    this.sharps = new Array(termSize).fill(null)

    this.scores = new Array(scoresSize)
    for (let i=0; i<scoresSize; i++) {
      this.scores[i] = new Array(termSize).fill(null)
    }
    this.totalScores = new Array(scoresSize).fill(0)
  }
}

class MinkabuRankingScraper {
  constructor() {
    this._sheetInfo = new SheetInfo()
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
    const pageNum = 180 // とりあえず3600件まで対応。

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
    
    const idecoIds = [      
      '8031114C', '64311042', '89311164', '32311984', '79313008', '9C31116A', 'AA311169', '4731198A', '89311025', '6831100B', '47311988', '29311041', '2931116A', '0131Q154', '7931211C', '93311164', 
      '65311058', '68311003', '89313135', '89311135', '25311177', '0431Q169', '64315021', '29316153', '0231202C', '20312061', '8931111A', '89312135', '0331109C', '01311021', '0331112A', '0331N029', 
      'AN311166', '89314135', '0231402C', '0131102B', '79312024', '0331110A', '0131F122', '2931201B', '0231502C', '0231602C', '0231702C', '0231802C', '0431X169', '64311081', '03312163', '03313163', 
      '03314163', '03315163', '03316163', '89313121', '89312121', '89314121', '89315121', '4731304C', '89315135', '29311078', '79314169', '01315087', '03311112', '04314086', 
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
    this._sheetInfo = new SheetInfo()
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
      for (let i=1; i<termSize + 1; i++) {
        const result1 = spanList[i].replace(/%/, '')
        const result2 = spanList[(termSize + 1) + i].replace(/%/, '')
        const result3 = spanList[2 * (termSize + 1) + i].replace(/%/, '')
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
    this._sheetInfo = new SheetInfo()
  }

  scraping(fund) {
    if (fund.category !== null) {
      return
    }
    
    const link = /(.*)\/.*/.exec(fund.link)[1]
    const html = UrlFetchApp.fetch(link).getContentText()
    const tables = Parser.data(html).from('<table class="md_table md_table_vertical">').to('</table>').iterate()
    const tds = Parser.data(tables[1]).from('<td').to('</td>').iterate().map(t => t.replace(' colspan="3"', '').replace('>', ''))
    fund.category = tds[1] // + '-' + tds[3

//    const link2 = link + '/sales_company'
//    const html2 = UrlFetchApp.fetch(link2).getContentText()
//    const table2 = Parser.data(html2).from('<table class="md_table">').to('</table>').build()
//    fund.isRakuten = table2.indexOf("楽天証券") !== -1
  }
}
