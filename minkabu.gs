
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
    const sheet = this._sheetInfo.getSheet(this._fundSheetName)
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

class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new SheetInfo() 
    this._funds = new Map()
    this._ignoreList = [
      'ＤＩＡＭ新興市場日本株ファンド', 'ＦＡＮＧ＋インデックス・オープン', 'グローバル・プロスペクティブ・ファンド（イノベーティブ・フューチャー）', 'ダイワ／バリュー・パートナーズ・チャイナ・イノベーター・ファンド',
      '野村世界業種別投資シリーズ（世界半導体株投資）', '東京海上Ｒｏｇｇｅニッポン海外債券ファンド（為替ヘッジあり）', '三菱ＵＦＪ先進国高金利債券ファンド（毎月決算型）（グローバル・トップ）',
      '三菱ＵＦＪ先進国高金利債券ファンド（年１回決算型）（グローバル・トップ年１）', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ａコース', '野村ＳＮＳ関連株投資Ａコース', 'ＵＳテクノロジー・イノベーターズ・ファンド（為替ヘッジあり）',
      '野村米国ブランド株投資（円コース）毎月分配型', 'ＵＳテクノロジー・イノベーターズ・ファンド', 'グローバル全生物ゲノム株式ファンド（１年決算型）', '野村米国ブランド株投資（円コース）年２回決算型',
      'グローバル・モビリティ・サービス株式ファンド（１年決算型）（グローバルＭａａＳ（１年決算型））', 'グローバル・モビリティ・サービス株式ファンド（１年決算型）（グローバルＭａａＳ（１年決算型））',
      'リスク抑制世界８資産バランスファンド（しあわせの一歩）', 'スパークス・ベスト・ピック・ファンド（ヘッジ型）', '世界８資産リスク分散バランスファンド（目標払出し型）（しあわせのしずく）',
      'グローバル・ハイクオリティ成長株式ファンド（年２回決算型）（限定為替ヘッジ）（未来の世界（年２回決算型））', 'ＪＰ日米バランスファンド（ＪＰ日米）', 'ＧＳフューチャー・テクノロジー・リーダーズＡコース（限定為替ヘッジ）（ｎｅｘｔＷＩＮ）',
      '野村世界業種別投資シリーズ（世界半導体株投資）', 'ＵＢＳ中国株式ファンド', '野村米国ブランド株投資（円コース）毎月分配型', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ｂコース',
      'ＵＢＳ中国Ａ株ファンド（年１回決算型）（桃源郷）', 'アライアンス・バーンスタイン・米国成長株投信Ｃコース毎月決算型（為替ヘッジあり）予想分配金提示型', 'アライアンス・バーンスタイン・米国成長株投信Ｄコース毎月決算型（為替ヘッジなし）予想分配金提示型',
      'グローバル・フィンテック株式ファンド（為替ヘッジあり・年２回決算型）', 'グローバル・フィンテック株式ファンド（年２回決算型）', 'グローバル・スマート・イノベーション・オープン（年２回決算型）（ｉシフト）',
      'グローバル・スマート・イノベーション・オープン（年２回決算型）為替ヘッジあり（ｉシフト（ヘッジあり））', '野村ＳＮＳ関連株投資Ｂコース',
    ]
    this._blockList = ['公社債投信.*月号', '野村・第.*回公社債投資信託']
  }
      
  calc() {
    this._fetchFunds()
    this._calcScores()
    this._output()
  }

  _fetchFunds() {
    this._sheetInfo.fundsSheetNames.forEach(sheetName => {
      const sheet = this._sheetInfo.getSheet(sheetName)
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
            fund.returns[i] = Number(r)
          }
          if (s !== '') {
            fund.sharps[i] = Number(s)
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
//        fund.scores[1][i] = fund.returns[i] !== 0 ? fund.sharps[i] / Math.abs(fund.returns[i]) : 0 // 最小分散ポートフォリオ戦略。思ったより微妙でがっかり。
        fund.scores[2][i] = Math.sqrt(Math.abs(fund.returns[i])) * fund.sharps[i] // iDeCo版
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
        fund.scores[n] = fund.scores[n].map((score, i) => score === null ? null : Math.sqrt(Math.log(score + Math.E)))
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
    
//    const maxList = scoresList.map(s => Math.max(...s))
//    const minList = scoresList.map(s => Math.min(...s))

    // 下位1%
    const lowList = scoresList.map(scores => {
      scores.sort((a, b) => a - b)
      return scores[parseInt(scores.length/100)]
    })
    console.log("aveList", aveList, "srdList", srdList, "lowList", lowList)
    return [aveList, srdList, lowList]
  }

  _normalizeAndInit(n, max, min) {
    console.log('normalizeAndInit')

    // スコア基本戦略：リターンxシャープレシオ→ログ化→ルート化→下位1%を切り捨て減算→正規化→中心極限定理の端っこの部分を抽出
    
    const [aveList, srdList, lowList] = this._analysis(this._getScoresList(n))
    this._funds.forEach(fund => {
      fund.scores[n] = fund.scores[n].map((score, i) => {
        if (score === null) {
          return null
        }

        const res = 10000000000 * (score - lowList[i]) // 歪みをボトムランクに移す。なぜか若干引き算するとうまくいく。最下位層のデータが悪さをしてるのかも？
        return Math.sign(res) * Math.pow(Math.abs(res), Math.pow(2, i === 0 ? -4 : -1))  // 3ヶ月の i === 0 のときのみ、より分散を小さくする。
      })
    })

    const [aveList2, srdList2] = this._analysis(this._getScoresList(n))
    this._funds.forEach(fund => {
      fund.scores[n] = fund.scores[n].map((score, i) => score === null ? null : (score - aveList2[i]) / srdList2[i])
    })

    const topPer = 100 // 100位がギリギリな雰囲気あるので、今後の展開によっては90位、80位でもよいかも？
    const initList = this._getScoresList(n).map((scores, i) => {
      scores.sort((a, b) => b - a)
      return scores[topPer]
    })
    
    this._funds.forEach(fund => {
      fund.scores[n] = fund.scores[n].map((score, i) => {
        return 5 * (score === null ? (n === 2 ? 0 : initList[i]) : score)
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
    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    const nameCol = 3
    const isIdecoCol = 4
    const totalScoreCol = 5

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

    sheet.autoResizeColumn(nameCol)
    for (let i=0; i<scoresSize; i++) {
      sheet.getRange(1, totalScoreCol + i * (termSize + 2), sheet.getLastRow()).setFontWeight("bold")
    }

    let n = 1
    this._funds.forEach(fund => {
      if (fund.ignore) {
        sheet.getRange(n, nameCol, 1, totalScoreCol + scoresSize * termSize - 1).setBackground('gray')
      }
      if (fund.isIdeco) {
        sheet.getRange(n, isIdecoCol).setBackground('yellow')
      }
      n++
    })

    const allRange = sheet.getDataRange()
    this._setColors(sheet, allRange, totalScoreCol, nameCol)
    allRange.sort({column: totalScoreCol, ascending: false})
  }

  _setColors(sheet, allRange, totalScoreCol, nameCol) {
    const white = '#ffffff' // needs RGB color
    const colors = ['cyan', 'lime', 'yellow', 'orange', 'pink', 'silver', ' white', ' white', ' white']

    allRange.sort({column: totalScoreCol, ascending: false})
    const nameRange = sheet.getRange(1, nameCol, sheet.getLastRow())
    this._setHighRankColor(10, nameRange)

    for (let i=totalScoreCol; i < totalScoreCol + scoresSize * (2 + termSize) - 1; i++) {
      let c = false
      for (let j=1; j<scoresSize; j++) {
        if (i === totalScoreCol + j * (termSize + 2) - 1) {
          c = true
        }
      }
      if (c) {
        continue
      }
    
      allRange.sort({column: i, ascending: false})
      
      let j = 0
      const range = sheet.getRange(1, i, 5 * colors.length)
      const rgbs = range.getBackgrounds().map(rows => {
        return rows.map(rgb => {
          if (rgb !== white) {
            return rgb
          }
          const result = colors[parseInt(j / 5)]
          j++;

          return result
        })
      })
      range.setBackgrounds(rgbs)
    }
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
