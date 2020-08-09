class SheetInfo {
  constructor() {
    this.fundsSheetNum = 10
    this.linkSheetName = 'Link'
    this.fundsSheetNames = [];
    for(let i=0; i<this.fundsSheetNum; i++) {
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

    this.termNum = 6
    this.date = null
    this.name = null
    this.ignore = false
    this.returns = new Array(this.termNum).fill(null)
    this.sharps = new Array(this.termNum).fill(null)
    this.scores1 = new Array(this.termNum).fill(null)
    this.scores2 = new Array(this.termNum).fill(null)
    this.scores3 = new Array(this.termNum).fill(null)
    this.totalScore1 = 0
    this.totalScore2 = 0
    this.totalScore3 = 0
  }
  
//  sqrtTarget() {
//    const powExp = [0.24, 0.89, 0.89, 1.19, 1.19, 1.74]
//    return this.returns.map((r, i) => {
//      return this._calcTarget(r, Math.sqrt(Math.abs(r)), i, powExp)
//    })
//  }
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
    for(let p=1; p<=pageNum; p++) {
      const link = baseLink + '/ranking/return?page=' + p
      const html = UrlFetchApp.fetch(link).getContentText()
      const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
      const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
      if(aList.length === 0 || aList[0] === '<!DOCTYPE htm') {
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
      
      if(p%20 === 0) {
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
    for(let i=n * this._fundsSheetNum; i<end; i++) {
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
      for(let i=0; i<fund.termNum; i++) {
        const result1 = spanList[i].replace(/%/, '')
        const result2 = spanList[sharpNum + i].replace(/%/, '')
        fund.returns[i] = result1 != '-' ? Number(result1) : null
        fund.sharps[i] = result2 != '-' ? Number(result2) : null
      }
      fund.date = Parser.data(html).from('<span class="fsm">（').to('）</span>').build()
    
      i++
      if(i%50 === 0) {
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
      'ＤＩＡＭ新興市場日本株ファンド', 'ＳＢＩ中小型成長株ファンドジェイネクスト（ｊｎｅｘｔ）', 'ＦＡＮＧ＋インデックス・オープン', 'ＳＢＩ中小型割安成長株ファンドジェイリバイブ（ｊｒｅｖｉｖｅ）',
      '野村世界業種別投資シリーズ（世界半導体株投資）', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ａコース', 'ＵＢＳ中国株式ファンド', 'ＵＢＳ中国Ａ株ファンド（年１回決算型）（桃源郷）',
      '野村ＳＮＳ関連株投資Ａコース', 'ダイワ／バリュー・パートナーズ・チャイナ・イノベーター・ファンド', 'グローバル全生物ゲノム株式ファンド（１年決算型）', 'テトラ・エクイティ', 'ブラックロック・ゴールド・メタル・オープンＡコース',
      'グローバル・プロスペクティブ・ファンド（イノベーティブ・フューチャー）', '野村米国ブランド株投資（円コース）毎月分配型', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ｂコース',
      '野村ＳＮＳ関連株投資Ｂコース', '野村米国ブランド株投資（円コース）年２回決算型', 'ＵＳテクノロジー・イノベーターズ・ファンド（為替ヘッジあり）', 'ＵＢＳ次世代テクノロジー・ファンド',
      'グローバル・モビリティ・サービス株式ファンド（１年決算型）（グローバルＭａａＳ（１年決算型））', 'ＵＳテクノロジー・イノベーターズ・ファンド', 'ＵＳテクノロジー・イノベーターズ・ファンド（為替ヘッジあり）',
      'グローバル・ハイクオリティ成長株式ファンド（年２回決算型）（限定為替ヘッジ）（未来の世界（年２回決算型））', '野村米国ブランド株投資（米ドルコース）毎月分配型'
    ]
    this._termNum = null
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
        if(value.length === 1) {
          return
        }
    
        const fund = new Fund(value[1], value[2])
        fund.date = value[0]
        fund.name = value[3]
        fund.ignore = this._ignoreList.some(i => i === fund.name)
      
        if (this._termNum === null) {
          this._termNum = fund.termNum
        }
        
        for(let i=0; i<this._termNum; i++) {
          const r = value[3*i + this._termNum - 1]
          const s = value[3*i + this._termNum]
          if(r !== '') {
            fund.returns[i] = r
          }
          if(s !== '') {
            fund.sharps[i] = s
          }
        }
        this._funds.set(fund.link, fund)
      })
    })
  }
  
  _calcScores() {
    const powExp = [0.21, 0.69, 0.69, 0.89, 0.89, 1.11]
//    const powExp = [1,1,1,1,1,1]
    this._funds.forEach(fund => {
      fund.scores1 = fund.scores1.map((_, i) => {
        if(fund.returns[i] === null || fund.sharps[i] === null) {
          return null
        }
      
        const base = fund.returns[i] * fund.sharps[i]
        return Math.sign(fund.returns[i]) * Math.pow(Math.abs(base), powExp[i])
      })
    })      
    
    let scoresList1 = []
    this._funds.forEach(fund => {
      scoresList1.push(fund.scores1)
    })      
    scoresList1 = scoresList1[0].map((_, i) => scoresList1.map(r => r[i]).filter(Boolean)) // transpose

    const [aveList, srdList, medList, medList2] = this._analysis(scoresList1)
    this._funds.forEach(fund => {
      fund.scores1 = this._standardize(fund.scores1, aveList, srdList, medList)
      fund.totalScore1 = fund.scores1.reduce((acc, score) => acc + score)
    })
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
  
    const medList = scoresList.map((scores, i) => {
      const n = i === 0 ? 1 : 2
      return aveList[i] + n*srdList[i]
//    scores = scores.sort((a, b) => a - b)
//    return scores[parseInt(scores.length*96/100)]
    })
  
    // 上位50%  iDeCo用 #TODO
    const medList2 = scoresList.map((scores, i) => {
      scores = scores.sort((a, b) => a - b)
      return scores[parseInt(scores.length/2)]
    })
    console.log(aveList, srdList, medList, medList2)
    return [aveList, srdList, medList, medList2]
  }

  _standardize(scores, aveList, srdList, medList) {
    return scores.map((score, i) => ((score || medList[i]) - aveList[i]) / srdList[i])
  }

  _output() { 
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = spreadsheet.insertSheet(this._sheetInfo.getScoreSheetName(), 0)

    const data = []
    const scoreNum = 1
    this._funds.forEach(fund => {
      const row = [fund.link, fund.date, fund.name, fund.isIdeco, fund.totalScore1, fund.scores1, ''].flat()
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
      if(fund.ignore) {
        sheet.getRange(n, nameRow).setBackground('gray')
      }
      if(fund.isIdeco) {
        sheet.getRange(n, isIdecoRow).setBackground('yellow')
      }
      n++
    })
  
    this._setColors(sheet, totalScoreRow, nameRow, scoreNum)
    sheet.getRange(1, totalScoreRow, sheet.getLastRow()).setFontWeight("bold")
    sheet.autoResizeColumn(nameRow)
  }

  _setColors(sheet, totalScoreRow, nameRow, scoreNum) {
    const colors = ['cyan', 'lime', 'yellow', 'orange', 'pink', 'silver']
    for(let i=totalScoreRow; i < totalScoreRow + scoreNum*(2 + this._termNum) - 1; i++) {
      let c = false
      for(let j=1; j<scoreNum; j++) {
        if(i === totalScoreRow + j*(this._termNum + 2) - 1) {
          c = true
        }
      }
      if (c) {
        continue
      }
    
      sheet.getDataRange().sort({column: i, ascending: false})
      colors.forEach((color, m) => {
        sheet.getRange(1 + 5*m, i, 5).setBackground(color)
      })
    }
  
    const allRange = sheet.getDataRange()
  
    const idecoRankRow = totalScoreRow + 2*(2 + this._termNum)
    allRange.sort({column: idecoRankRow, ascending: false})
    allRange.sort({column: 4, ascending: false})
    const idecoRange = sheet.getRange(1, idecoRankRow, sheet.getLastRow())
    this._setHighRankColor(5, idecoRange)

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
        if(i >= max || rgb !== white) {
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
  for(let i=0; i<sheetInfo.fundsSheetNum; i++) {
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
//    if(i >= ignoreSheet.length) {
//      spreadsheet.deleteSheet(sheet)
//    }
//  })
//}
