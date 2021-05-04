function scrapingMinkabuFunds() {
  (new MinkabuFundsScraper(0)).scraping()
}

class Fund {
  constructor(link, isIdeco) {
    this.link = link
    this.isIdeco = isIdeco
    
    this.name = null // 名前
    this.management = null // 資産運用会社

    this.rating = null // レーティング
    this.assets = null // 純資産額
    this.salesCommission = null // 販売手数料
    this.fee = null // 信託報酬
    this.redemptionFee = null // 信託財産留保額
    this.category = null // 分類
    this.launchDate = null // 設定年月日
    this.commissionPeriod = null // 信託期間
    
    this.returns = new Array(minkabuTermSize).fill(null)
    this.risks = new Array(minkabuTermSize).fill(null)
    this.sharps = new Array(minkabuTermSize).fill(null)

    this.isSaledRakuten = null // 楽天証券
    this.isSaledSbi = null // SBI証券
    this.isSaledAu = null // auカブコム証券
    this.isSaledMonex = null // マネックス証券
    this.isSaledMatsui = null // 松井証券
    this.isSaledNomura = null // 野村證券
    this.isSaledDaiwa = null // 大和証券
    this.isSaledSmbc = null // SMBC日興証券
    this.isSaledLine = null // LINE証券
    this.isSaledOkasan = null // 岡三オンライン証券
    this.isSaledPayPay = null // PayPay銀行・証券
    this.isSaledMitsui = null // 三井住友銀行・証券
    this.isSaledUfj = null // 三菱UFJ信託銀行・証券
    this.isSaledYutyo = null // ゆうちょ銀行・証券
    this.isSaledMizuho = null // みずほ銀行・証券

    this.ignore = false
    this.policy = new Array(minkabuTermSize).fill(null)
    this.scores = new Array(minkabuTermSize).fill(null)
    this.totalScoresOf10531 = null
    this.totalScoresOf531 = null
    this.totalScoresOf53 = null
    this.totalScoresOf31 = null
  }
}

class MinkabuSheetInfo {
  constructor() {
    this.linkSheetName = 'Link'
    this.fundsSheetNames = [];
    for (let i=0; i<fundsSheetMax; i++) {
      this.fundsSheetNames.push('Funds' + i)
    }
    this.logSheetName = 'Log'
    
    this.scoreSheetsName = ['2020年9月', '2020年10月', '2020年11月', '2020年12月', '2021年1月', '2021年2月', '2021年3月', '2021年4月']
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

    let i=0;
    for (let p=1; p<=minkabuRankPageNum; p++) {
      const link = baseLink + '/ranking/return?page=' + p
      const html = UrlFetchApp.fetch(link).getContentText()
      const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
      const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
      if (aList.length === 0 || aList[0].indexOf('<meta') != -1) {
        return
      }
    
      aList.forEach(a => {
        const result = a.match(/href="(.*)">(.*)/)
        const pass = result[1]
        const link = baseLink + pass
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
    // SBI Select版
    const idecoIds = [
      '2931316B', '03317172', '9C31116A', '0131C18A', '04316186', '89311164', '8931217C', '03316183', '89313135', '96312073', 
      '03319172', '2931113C', '03311187', '0431Q169', '0231202C', '25311177', '65311058', '68311003', '0331C177', '8931111A', 
      '03318172', '0331A172', '0231402C', 'AN31118A', '0431U169', '29314136', '79314169', '03312175', '04316188', '8931118A',
      '96311073', '03311112', '89312121', '89313121', '89314121', '89315121',
    ]
    
    idecoIds.forEach(id => {
      const link = 'https://itf.minkabu.jp/fund/' + id
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
    const end = Math.min(values.length, minkabuScrapingNum * (this._fundsSheetNum + 1))
    for (let i=minkabuScrapingNum * this._fundsSheetNum; i<end; i++) {
      this._funds.set(values[i][0], new Fund(values[i][0], values[i][1]))
    }
    console.log('_fetchLinks')
  }

  _fetchDetail() {
    console.log('getDetailFromMinkabu:' + this._funds.size)
    
    let i = 0
    const c = 2
    this._funds.forEach(fund => {
      const html = UrlFetchApp.fetch(fund.link).getContentText()
      const tables = Parser.data(html).from('<table class="md_table">').to('</table>').iterate()

      // 名前
      const fundDiv = this._toHankaku(Parser.data(html).from('<div class="fund_name">').to('</div>').build())
      fund.name = this._toHankaku(Parser.data(html).from('<h1>').to('</h1>').build())
      fund.management = this._toHankaku(Parser.data(fundDiv).from('<span>').to('</span>').build())

      const tdList0 = Parser.data(tables[0]).from('<td class="tar">').to('</td>').iterate()
      const tdList1 = Parser.data(tables[1]).from('<td class="tar">').to('</td>').iterate()

      fund.rating = Parser.data(tdList0[0]).from('<span class="gold_star">').to('</span>').build() // レーティング

      // 純資産額
      const amountList = ['兆', '億', '万'].map(s => {
        const regexp = new RegExp('([0-9]*)' + s)
        const res = tdList0[2].match(regexp)
        return res === null ? 0 : res[1]
      })
      fund.assets = 100000000 * Number(amountList[0]) + 10000 * Number(amountList[1]) + Number(amountList[2])

      fund.salesCommission = tdList1[0] // 販売手数料
      fund.fee = tdList1[1] // 信託報酬
      fund.redemptionFee = tdList1[2] // 信託財産留保額

      const tdList2 = Parser.data(tables[3]).from('<td>').to('</td>').iterate()

      fund.category = tdList2[1] // 分類
      fund.launchDate = tdList2[4] // 設定年月日
      fund.commissionPeriod = tdList2[5] // 信託期間

      // 1年3年5年10年のデータを取得する。3ヶ月と6ヶ月は無視。

      // なぜか目論見書が無いときにテーブルがずれるので調整
      let spanList = Parser.data(tables[4]).from('<span>').to('</span>').iterate()
      if(spanList.length < 2) {
        spanList = Parser.data(tables[3]).from('<span>').to('</span>').iterate()
      }

      for (let i=c; i<minkabuTermSize + c; i++) {
        const result1 = spanList[i].replace(/%/, '')
        const result2 = spanList[(minkabuTermSize + c) + i].replace(/%/, '')
        const result3 = spanList[2 * (minkabuTermSize + c) + i].replace(/%/, '')　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
        fund.returns[i - c] = result1 != '-' ? Number(result1) : null
        fund.risks[i - c] = result2 != '-' ? Number(result2) : null
        fund.sharps[i - c] = result3 != '-' ? Number(result3) : null
      }

      const salsLink = fund.link + '/sales_company'
      const html2 = UrlFetchApp.fetch(salsLink).getContentText()
      const tables2 = Parser.data(html2).from('<table class="md_table">').to('</table>').iterate()
      fund.isSaledRakuten = tables2[0].includes("楽天") // 楽天証券
      fund.isSaledSbi = tables2[0].includes("ＳＢＩ") // SBI証券
      fund.isSaledAu = tables2[0].includes("ａｕカブコム") // auカブコム証券
      fund.isSaledMonex = tables2[0].includes("マネックス") // マネックス証券
      fund.isSaledMatsui = tables2[0].includes("松井") // 松井証券
      fund.isSaledNomura = tables2[0].includes("野村") // 野村證券
      fund.isSaledDaiwa = tables2[0].includes("大和") // 大和証券
      fund.isSaledSmbc = tables2[0].includes("ＳＭＢＣ") // SMBC日興証券
      fund.isSaledLine = tables2[0].includes("ＬＩＮＥ") // LINE証券
      fund.isSaledOkasan = tables2[0].includes("岡三") // 岡三オンライン証券
      fund.isSaledPayPay = tables2[0].includes("ＰａｙＰａｙ") // PayPay銀行・証券
      fund.isSaledMitsui = tables2[0].includes("三井住友") // 三井住友銀行・証券
      fund.isSaledUfj = tables2[0].includes("三菱ＵＦＪ") // 三菱UFJ信託銀行・証券
      fund.isSaledYutyo = tables2[0].includes("ゆうちょ") // ゆうちょ銀行・証券
      fund.isSaledMizuho = tables2[0].includes("みずほ") // みずほ銀行・証券


      i++
      if (i%20 === 0) {
        console.log(i)
      }
    })
    console.log('_fetchDetail')
  }

  _output() {
    const data = []
    this._funds.forEach(fund => {
      const row = []
      fund.returns.forEach((r, i) => {
        row.push(r, fund.risks[i], fund.sharps[i], '')
      })
      row.push(fund.link, fund.isIdeco, fund.name, fund.management, fund.rating, fund.assets, fund.salesCommission, fund.fee, fund.redemptionFee, fund.category, fund.launchDate, 
        fund.commissionPeriod, '', fund.isSaledRakuten, fund.isSaledSbi, fund.isSaledAu, fund.isSaledMonex, fund.isSaledMatsui, fund.isSaledNomura, fund.isSaledDaiwa, fund.isSaledSmbc,
        fund.isSaledLine, fund.isSaledOkasan, fund.isSaledPayPay, fund.isSaledMitsui, fund.isSaledUfj, fund.isSaledYutyo, fund.isSaledMizuho)
 
      data.push(row)
    })

    const sheet = this._sheetInfo.getSheet(this._fundSheetName)
    sheet.clear()
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
//    sheet.autoResizeColumns(1, 25) // めちゃくちゃ重い
    console.log('_output')
  }
  
  _toHankaku(str) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９！-～]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    }).replace(/”/g, "\"").replace(/’/g, "'").replace(/‘/g, "`").replace(/￥/g, "\\").replace(/　/g, " ").replace(/〜/g, "~")
  }
}
