
const termSize = 6
const scoresSize = 3

class SheetInfo {
  constructor() {
    this.fundsSheetNum = 10
    this.linkSheetName = 'Link'
    this.fundsSheetNames = [];
    for (let i=0; i<this.fundsSheetNum; i++) {
      this.fundsSheetNames.push('Funds' + i)
    }
  }
  
  getScoreSheetName() {
    const prefix = 'スコア：'
    return prefix + (new Date().toLocaleString("ja")) 
  }
}

class Fund {
  constructor(link, isIdeco) {
    this.link = link
    this.isIdeco = isIdeco
    
    this.date = null
    this.name = null
    this.ignore = false
    this.returns = new Array(termSize).fill(null)
    this.sharps = new Array(termSize).fill(null)

    this.scores = new Array(scoresSize)
    for (let i=0; i<scoresSize; i++) {
      this.scores[i] = new Array(termSize).fill(null)
    }
    this.totalScores = new Array(scoresSize).fill(0)
  }
}

class RankingScraper {
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
  
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this._sheetInfo.linkSheetName)
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
    const idecoIds = [
      '03311112', '25311177', '93311164', '32311984', '0131F122', '0131102B', '8031114C', 'AA311169', '0231802C', '0131Q154', '89313121', '89315121', '65311058', '89312121', '0231402C',
      '89314121', '0231602C', '68311003', '0231502C', '0231202C', '0331109C', '0231702C', '8931111A', '89311025', '2931116A', '0331110A', '64315021', '64311081', '0431Q169', '89311135',
      '89313135', '0331112A', '89311164', '9C31116A', '7931211C', '79314169', '2931216A', '89312135', '0431X169', '29316153', '8931217C', '8931118A', 'AN31118A', '96311073', '96312073',
      '0431U169', '04316188', '04316186', '2931113C', '29314136', '2931316B', '0131C18A', '03317172', '03318172', '0331C177', '03319172', '0331A172', '03316183', '03312175', '03311187',
    ]
    
    idecoIds.forEach(id => {
      const link = 'https://itf.minkabu.jp/fund/' + id + '/risk_cont'
      this._funds.set('/fund/' + id, new Fund(link, true))
    })
    console.log('_getIdecoFunds', this._funds.size)
  }
}

class FundsScraper {
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this._sheetInfo.linkSheetName)
    const values = sheet.getDataRange().getValues()
    const n = 360
    const end = Math.min(values.length, n * (this._fundsSheetNum + 1))
    for (let i=n * this._fundsSheetNum; i<end; i++) {
      this._funds.set(values[i][0], new Fund(values[i][0], values[i][1]))
    }
    console.log('_fetchLinks')
  }

  _fetchDetail() {
    const sharpNum = 12 // 12番目からのspanがシャープレシオ
    console.log('getDetailFromMinkabu:' + this._funds.size)
    
    let i = 0
    this._funds.forEach(fund => {
      const html = UrlFetchApp.fetch(fund.link).getContentText()
      const table = Parser.data(html).from('<table class="md_table">').to('</table>').build()
      const spanList = Parser.data(table).from('<span>').to('</span>').iterate()
    
      fund.name = Parser.data(html).from('<p class="stock_name">').to('</p>').build()
      for (let i=0; i<termSize; i++) {
        const result1 = spanList[i].replace(/%/, '')
        const result2 = spanList[sharpNum + i].replace(/%/, '')
        fund.returns[i] = result1 != '-' ? Number(result1) : null
        fund.sharps[i] = result2 != '-' ? Number(result2) : null
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this._fundSheetName)
    sheet.clear()

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.date, fund.link, fund.isIdeco, fund.name, '']
      fund.returns.forEach((r, i) => {
        row.push(r, fund.sharps[i], '')
      })
      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    sheet.autoResizeColumns(2, 4)
    console.log('_output')
  }
}

class FundsScoreCalculator {
  constructor() {
    this._sheetInfo = new SheetInfo() 
    this._funds = new Map()
    this._ignoreList = [
      'ＤＩＡＭ新興市場日本株ファンド', 'ＦＡＮＧ＋インデックス・オープン', 'グローバル・プロスペクティブ・ファンド（イノベーティブ・フューチャー）', 'ダイワ／バリュー・パートナーズ・チャイナ・イノベーター・ファンド',
      '野村世界業種別投資シリーズ（世界半導体株投資）', '東京海上Ｒｏｇｇｅニッポン海外債券ファンド（為替ヘッジあり）', '三菱ＵＦＪ先進国高金利債券ファンド（毎月決算型）（グローバル・トップ）',
    ]
    this._blockList = ['公社債投信.*月号', '野村・第.*回公社債投資信託', 'ＭＨＡＭ・公社債投信.*月号']
  }
      
  calc() {
    this._fetchFunds()
    this._calcScores()
    this._output()
  }

  _fetchFunds() {
    this._sheetInfo.fundsSheetNames.forEach(sheetName => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
      const values = sheet.getDataRange().getValues()
      values.forEach(value => {
        if (value.length === 1) {
          return
        }
    
        const fund = new Fund(value[1], value[2])
        fund.date = value[0]
        fund.name = value[3]
        fund.ignore = this._ignoreList.some(i => i === fund.name)
        if (this._blockList.some(b => fund.name.match(b) !== null)) {
          return
        }
      
        for (let i=0; i<termSize; i++) {
          const r = value[3*i + termSize - 1]
          const s = value[3*i + termSize]
          if (r !== '') {
            fund.returns[i] = r
          }
          if (s !== '') {
            fund.sharps[i] = s
          }
        }
        this._funds.set(fund.link, fund)
      })
    })
  }
  
  _calcScores() {
    this._funds.forEach(fund => {
      fund.returns.forEach((r, i) => {
        if (r === null || fund.sharps[i] === null) {
          return
        }
        fund.scores[0][i] = Math.abs(fund.returns[i]) * fund.sharps[i]
        fund.scores[1][i] = fund.sharps[i]
        fund.scores[2][i] = Math.abs(fund.returns[i]) * fund.sharps[i]
      })
    })

    for (let n=0; n<scoresSize; n++) {
      const minList1 = this._getScoresList(n).map(s => Math.min(...s))
      // バグ埋め込みやすいので消すな。score === null のワナ
      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map((score, i) => score === null ? null : score - minList1[i])
      })
      
      const ignoreNum = Number.MIN_VALUE
      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map((score, i) => score === null ? null : Math.log(Math.sqrt(score) + Math.E) - 1)
      })
      
//      // 下位を外れ値として切り捨てる。正規化が安定する。試行錯誤の上、100で固定。
//      const minList2 = this._getScoresList(n).map(scores => (scores.sort((a, b) => a - b))[parseInt(scores.length/100)])
//      this._funds.forEach(fund => {
//        fund.scores[n] = fund.scores[n].map((score, i) => score === null ? null : score <= minList2[i] ? minList2[i] : score)
//      })
      
      this._normalizeAndInit(n, 100, 0)
      
      this._funds.forEach(fund => {
        fund.totalScores[n] = fund.scores[n].reduce((acc, score) => acc + score)
      })
    }
  }
    
  _getScoresList(n) {
    const scoresList = []
    this._funds.forEach(fund => scoresList.push(fund.scores[n]))
    return scoresList[0].map((_, i) => scoresList.map(r => r[i]).filter(Boolean)) // transpose
  }

  _analysis(scoresList) {
    const sumList = scoresList.map(scores => scores.reduce((acc, v) => acc + v))
    const aveList = sumList.map((sum, i) => sum/scoresList[i].length)
    const srdList = scoresList.map((scores, i) => {
      let sum = 0
      scores.forEach(score => {
        sum += Math.pow(score - aveList[i], 2)
      })
      return Math.sqrt(sum / (scores.length - 1))
    })

    //  http://www.gaoshukai.com/20/19/0001/
    const z = 2.576	// 99.0004935369%
    const initList = scoresList.map((_, i) => aveList[i] + z * srdList[i])

    // iDeCo用
    const initList2 = scoresList.map((_, i) => aveList[i])
    
    const maxList = scoresList.map(s => Math.max(...s))
    const minList = scoresList.map(s => Math.min(...s))
    
    console.log(aveList, srdList, initList, initList2, maxList, minList)
    return [aveList, srdList, initList, initList2, maxList, minList]
  }

  _normalizeAndInit(n, max, min) {
    console.log('normalizeAndInit')

    // 先に各期間の差を計算する。あえて差をつけるために反映は標準化したあとにする
    const [aveList, srdList, initList, initList2] = this._analysis(this._getScoresList(n))
    console.log('aveList', aveList)
    this._funds.forEach(fund => {
      const usedInitList = n === 2 ? initList2 : initList
      fund.scores[n] = fund.scores[n].map((score, i) => {
        return ((score === null ? usedInitList[i] : score) - aveList[i]) / srdList[i] / aveList[i] / (i === 0 ? 2 : 1)
      })
    })
    
    const ignoreNum = -100
    this._funds.forEach(fund => {
      if (n === 2) {
        fund.scores[n] = fund.scores[n].map(s => fund.isIdeco ? s : ignoreNum)
      }
    })    
  }
  
  _output() { 
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = spreadsheet.insertSheet(this._sheetInfo.getScoreSheetName(), 0)

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.link, fund.date, fund.name, fund.isIdeco]
      for (let i=0; i<scoresSize; i++) {
        row.push(fund.totalScores[i], ...(fund.scores[i]), '')
      }
      fund.returns.forEach((r, i) => {
        row.push('', r, fund.sharps[i])
      })
      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

    let n = 1
    const nameRow = 3
    const isIdecoRow = 4
    const totalScoreRow = 5
    this._funds.forEach(fund => {
      if (fund.ignore) {
        sheet.getRange(n, nameRow).setBackground('gray')
      }
      if (fund.isIdeco) {
        sheet.getRange(n, isIdecoRow).setBackground('yellow')
      }
      n++
    })
  
    this._setColors(sheet, totalScoreRow, nameRow)
    for (let i=0; i<scoresSize; i++) {
      sheet.getRange(1, totalScoreRow + i * (termSize + 2), sheet.getLastRow()).setFontWeight("bold")
    }
    sheet.autoResizeColumn(nameRow)
  }

  _setColors(sheet, totalScoreRow, nameRow) {
    const colors = ['cyan', 'lime', 'yellow', 'orange', 'pink', 'silver']
    for (let i=totalScoreRow; i < totalScoreRow + scoresSize * (2 + termSize) - 1; i++) {
      let c = false
      for (let j=1; j<scoresSize; j++) {
        if (i === totalScoreRow + j * (termSize + 2) - 1) {
          c = true
        }
      }
      if (c) {
        continue
      }
    
      sheet.getDataRange().sort({column: i, ascending: false})
      colors.forEach((color, m) => {
        sheet.getRange(1 + 5 * m, i, 5).setBackground(color)
      })
    }
  
    const allRange = sheet.getDataRange()
    allRange.sort({column: totalScoreRow, ascending: false})
    const nameRange = sheet.getRange(1, nameRow, sheet.getLastRow())
    this._setHighRankColor(10, nameRange)
  }

  _setHighRankColor(max, range) {
    const white = '#ffffff' // needs RGB color
    const aqua = 'aqua'

    let i = 0
    const rgbs = range.getBackgrounds().map(rows => {
      return rows.map(rgb => {
        if (i >= max || rgb !== white) {
          return rgb
        }
        i++;
        return aqua
      })
    })
    range.setBackgrounds(rgbs)
  }
}

// hook
function onOpen() {
  const sheetInfo = new SheetInfo()
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const menu = [{name: 'ランキングスクレイピング', functionName: 'scrapingRanking'}]
  for (let i=0; i<sheetInfo.fundsSheetNum; i++) {
    menu.push({name: 'ファンドスクレイピング' + i, functionName: 'scrapingFunds' + i})
  }
  menu.push({name: 'ファンドスコア計算', functionName: 'calcFundsScore'})
  sheet.addMenu("Google App Script", menu)
}

function scrapingRanking() {
  (new RankingScraper()).scraping()
}
 
function scrapingFunds0() {
  (new FundsScraper(0)).scraping()
}

function scrapingFunds1() {
  (new FundsScraper(1)).scraping()
}

function scrapingFunds2() {
  (new FundsScraper(2)).scraping()
}

function scrapingFunds3() {
  (new FundsScraper(3)).scraping()
}

function scrapingFunds4() {
  (new FundsScraper(4)).scraping()
}

function scrapingFunds5() {
  (new FundsScraper(5)).scraping()
}

function scrapingFunds6() {
  (new FundsScraper(6)).scraping()
}

function scrapingFunds7() {
  (new FundsScraper(7)).scraping()
}

function scrapingFunds8() {
  (new FundsScraper(8)).scraping()
}

function scrapingFunds9() {
  (new FundsScraper(9)).scraping()
}

function calcFundsScore() {
  (new FundsScoreCalculator).calc()
}

//function deleteAutoSheet() {
//  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
//  const ignoreSheet= ['過去メモ']
//  ignoreSheet.forEach(sheetName => {
//    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName(sheetName))
//    spreadsheet.moveActiveSheet(1)
//  })
//  
//  spreadsheet.getSheets().forEach((sheet, i) => {
//    if (i >= ignoreSheet.length) {
//      spreadsheet.deleteSheet(sheet)
//    }
//  })
//}
