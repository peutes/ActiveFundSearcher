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

class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new SheetInfo() 
    this._infoScraper = new MinkabuInfoScraper()
    
    this._funds = new Map()
    this._ignoreList = [
      "DIAM新興市場日本株ファンド", "FANG+インデックス・オープン", "野村世界業種別投資シリーズ(世界半導体株投資)", "アライアンス・バーンスタイン・米国成長株投信Cコース毎月決算型(為替ヘッジあり)予想分配金提示型",
      "野村クラウドコンピューティング&スマートグリッド関連株投信Aコース", "野村クラウドコンピューティング&スマートグリッド関連株投信Bコース", "野村米国ブランド株投資(円コース)毎月分配型",
      "UBS中国株式ファンド", "ＵＢＳ中国Ａ株ファンド（年１回決算型）（桃源郷）", "野村SNS関連株投資Aコース", "野村米国ブランド株投資(円コース)年2回決算型", "ダイワ/バリュー・パートナーズ・チャイナ・イノベーター・ファンド",
      "USテクノロジー・イノベーターズ・ファンド(為替ヘッジあり)", "グローバル・プロスペクティブ・ファンド(イノベーティブ・フューチャー)", "グローバル・モビリティ・サービス株式ファンド(1年決算型)(グローバルMaaS(1年決算型))",
      "野村SNS関連株投資Bコース", "USテクノロジー・イノベーターズ・ファンド", "UBS中国A株ファンド(年1回決算型)(桃源郷)", "UBS次世代テクノロジー・ファンド",
      "グローバル・フィンテック株式ファンド(為替ヘッジあり・年2回決算型)", "グローバル・フィンテック株式ファンド(年2回決算型)", "東京海上Roggeニッポン海外債券ファンド(為替ヘッジあり)",
      "リスク抑制世界8資産バランスファンド(しあわせの一歩)", "三菱UFJ先進国高金利債券ファンド(年1回決算型)(グローバル・トップ年1)",
      "三菱UFJ先進国高金利債券ファンド(毎月決算型)(グローバル・トップ)", "アライアンス・バーンスタイン・米国成長株投信Dコース毎月決算型(為替ヘッジなし)予想分配金提示型",
      "三菱UFJグローバル・ボンド・オープン(毎月決算型)(花こよみ)", "グローバルAIファンド(予想分配金提示型)", "グローバル・スマート・イノベーション・オープン(年2回決算型)為替ヘッジあり(iシフト(ヘッジあり))",
      "GSフューチャー・テクノロジー・リーダーズAコース(限定為替ヘッジ)(nextWIN)", "グローバル全生物ゲノム株式ファンド(1年決算型)", "世界8資産リスク分散バランスファンド(目標払出し型)(しあわせのしずく)",
      "グローバル・ハイクオリティ成長株式ファンド(年2回決算型)(限定為替ヘッジ)(未来の世界(年2回決算型))", "グローバル・スマート・イノベーション・オープン(年2回決算型)(iシフト)",
      "野村米国ブランド株投資(米ドルコース)毎月分配型", "グローバルAIファンド(為替ヘッジあり予想分配金提示型)", "野村米国ブランド株投資(米ドルコース)年2回決算型", "新興国ハイクオリティ成長株式ファンド(未来の世界(新興国))",
      "US成長株オープン(円ヘッジありコース)", "米国IPOニューステージ・ファンド<為替ヘッジあり>(年2回決算型)", "米国IPOニューステージ・ファンド<為替ヘッジなし>(年2回決算型)",
      "野村エマージング債券投信(金コース)毎月分配型", "JPMグレーター・チャイナ・オープン", "T&Dダブルブル・ベア・シリーズ7(ナスダック100・ダブルブル7)", "スパークス・ベスト・ピック・ファンド(ヘッジ型)",
      "野村エマージング債券投信(金コース)年2回決算型", "野村米国ブランド株投資(アジア通貨コース)年2回決算型", "野村米国ブランド株投資(アジア通貨コース)毎月分配型", "あい・パワーファンド(iパワー)",
      "SBI中小型成長株ファンドジェイネクスト(jnext)", "UBS中国新時代株式ファンド(年2回決算型)", "ニッセイAI関連株式ファンド(年2回決算型・為替ヘッジあり)(AI革命(年2・為替ヘッジあり))",
      "ニッセイAI関連株式ファンド(年2回決算型・為替ヘッジなし)(AI革命(年2・為替ヘッジなし))", "グローバル・ハイクオリティ成長株式ファンド(年2回決算型)(為替ヘッジなし)(未来の世界(年2回決算型))",
      "新光日本小型株ファンド(風物語)", "GSフューチャー・テクノロジー・リーダーズBコース(為替ヘッジなし)(nextWIN)", "次世代通信関連アジア株式戦略ファンド(THE ASIA 5G)",
      "インデックスファンド海外債券ヘッジあり(DC専用)", "ゴールドマン・サックス・世界債券オープンCコース(毎月分配型、限定為替ヘッジ)", "ゴールドマン・サックス・世界債券オープンDコース(毎月分配型、為替ヘッジなし)",
      "JP日米バランスファンド(JP日米)", "ほくよう資産形成応援ファンド(ほくよう未来への翼)", "野村PIMCO米国投資適格債券戦略ファンド(為替ヘッジあり)毎月分配型",
      "野村PIMCO米国投資適格債券戦略ファンド(為替ヘッジあり)年2回決算型", "三菱UFJ/ピムコトータル・リターン・ファンド<米ドルヘッジ型>(毎月決算型)", "MHAM USインカムオープン毎月決算コース(為替ヘッジなし)(ドルBOX)",
      "ダイワ米国国債7-10年ラダー型ファンド(部分為替ヘッジあり)-USトライアングル-", "USストラテジック・インカム・アルファ毎月決算型", "ダイワ債券コア戦略ファンド(為替ヘッジあり)",
      "USストラテジック・インカム・アルファ年1回決算型", "ダイワDBモメンタム戦略ファンド(為替ヘッジあり)", "JPM USコア債券ファンド(為替ヘッジあり、年1回決算型)",
      "ダイワDBモメンタム戦略ファンド(為替ヘッジなし)", "野村外国債券インデックスファンド(確定拠出年金向け)", "野村未来トレンド発見ファンドCコース(為替ヘッジあり)予想分配金提示型(先見の明)",
      "三井住友・DC外国債券インデックスファンド", "米国地方債ファンド2016-07(為替ヘッジあり)(ドリームカントリー)", "マニュライフ・米国投資適格債券戦略ファンドAコース(為替ヘッジあり・毎月)",
      "マニュライフ・米国投資適格債券戦略ファンドCコース(為替ヘッジあり・年2回)", "先進国投資適格債券ファンド(為替ヘッジあり)(マイワルツ)", "ゴールドマン・サックス・世界債券オープンAコース(限定為替ヘッジ)",
      "ステート・ストリート先進国債券インデックス・オープン(為替ヘッジあり)", "ゴールドマン・サックス・世界債券オープンBコース(為替ヘッジなし)", "海外債券インデックスファンド(個人型年金向け)(ゆうちょDC海外債券インデックス)",
      "フィデリティ世界医療機器関連株ファンド(為替ヘッジあり)", "フィデリティ世界医療機器関連株ファンド(為替ヘッジなし)", "US成長株オープン(円ヘッジなしコース)",
      "ダイワ米国国債ファンド-ラダー10-(為替ヘッジあり)", "ダイワ米国国債ファンド-ラダー10-(為替ヘッジなし)", "海外国債ファンド(3ヵ月決算型)", "ロボット・テクノロジー関連株ファンド-ロボテック-(為替ヘッジあり)",
      "DIAM外国債券パッシブ・ファンド", "フィデリティ世界医療機器関連株ファンド(為替ヘッジなし)", "野村外国債券インデックスファンド", "三井住友・A株メインランド・チャイナ・オープン", 
      "マネックス・日本成長株ファンド(ザ・ファンド@マネックス)", "テトラ・エクイティ", "スパークス・ベスト・ピック・ファンドⅡ(日本・アジア)マーケットヘッジあり",
    ]
//    this._blockList = ['^公社債投信.*月号$', '^野村・第.*回公社債投資信託$', '^MHAM・公社債投信.*月号$', '^日興・公社債投信.*月号$', '^大和・公社債投信.*月号$]
    this._blockList = []
      
    this.logSheet = this._sheetInfo.getSheet(this._sheetInfo.logSheetName)
    this.logSheet.clear()
  }
      
  calc() {
    this._fetchFunds()
    this._decidePolicy()
    this._calcScores()
    this._fetchCategory()
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
          const return_ = value[4*i + termSize]
          const risk = value[4*i + termSize + 1]
          const sharp = value[4*i + termSize + 2]
          if (return_ !== '') {
            fund.returns[i] = Number(return_)
          }
          if (risk !== '') {
            fund.risks[i] = Number(risk)
          }
          if (sharp !== '') {
            fund.sharps[i] = Number(sharp)
          }
        }
        this._funds.set(fund.link, fund)
      })
    })
  }
  
  _decidePolicy() {
    this._funds.forEach(fund => {
      fund.returns.forEach((r, i) => {
        if (r === null || fund.risks[i] === null || fund.sharps[i] === null) {
          return
        }
      
        // 基本方針：平均分散モデルのシャープレシオ最大化における
        // 必要最低限のフィルタは必要だった。expフィルタよりも√フィルタのほうが優秀
        // リターンが小さすぎる債券ファンドをフィルタで消す
        const w = 0.3  // リターンとリスクがあまりにも小さすぎるのを除去。公社債投信が0以下になるようにする。
        const w2 = 0.2  // リターンとリスクがあまりにも小さすぎるのを除去。公社債投信が0以下になるようにする。
        fund.scores[0][i] = fund.sharps[i] * Math.sqrt(Math.abs(r) * fund.risks[i] / ((Math.abs(r) + w) * (fund.risks[i] + w)))
        fund.scores[1][i] = fund.sharps[i] * Math.sqrt(Math.abs(r) * fund.risks[i] / ((Math.abs(r) + w2) * (fund.risks[i] + w2)))
        fund.scores[2][i] = fund.sharps[i] * Math.sqrt(Math.abs(r) * fund.risks[i] / ((Math.abs(r) + w) * (fund.risks[i] + w)))
      })
    })
  }
  
  _calcScores() {
    console.log('_calcScores')

    for (let n=0; n<scoresSize; n++) {
      console.log("n", n)
      const isIdecoScores = n === 2

      // 各期間ごとのスコアのバランスを整えるために標準化
      const [aveList, srdList] = this._analysis(this._getScoresList(n))
      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map((s, i) => s === null ? null : (s - aveList[i]) / srdList[i])
      })

      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map(s => isIdecoScores && !fund.isIdeco ? null : s)
      })

      // 初期値を自動決定するのに各期間のスコアのランキングを使う
      const rank = this._calcRank(n, isIdecoScores)
      const initList = this._getInitList(n, rank)
      console.log("initList", initList)
      
      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map((s, i) => s || initList[i])
        fund.totalScores[n] = fund.scores[n].reduce((acc, score) => acc + score, 0)
      })
      
      // 正規分布の信頼区間 95%ゾーンのスコア20以上を購入するのが望ましい -> 購入数は50~60がいいんじゃないかという裏付け
      const totalScores = []
      this._funds.forEach(fund => {
        if (!isIdecoScores || (isIdecoScores && fund.isIdeco)) {
          totalScores.push(fund.totalScores[n])
        }
      })
      const ave = totalScores.reduce((acc, v) => acc + v, 0) / totalScores.length
      const srd = Math.sqrt(totalScores.reduce((acc, v) => acc + Math.pow(v - ave, 2), 0) / totalScores.length)
      this._funds.forEach(fund => {
        if (!isIdecoScores || (isIdecoScores && fund.isIdeco)) {
          fund.totalScores[n] = 10 * (fund.totalScores[n] - ave) / srd
        } else {
          fund.totalScores[n] = 0
        }
      })
    }
  }

  _getScoresList(n) {
    const scoresList = []
    this._funds.forEach(fund => scoresList.push(fund.scores[n]))
    return scoresList[0].map((_, i) => scoresList.map(r => r[i]).filter(Boolean)) // transpose
  }

  _analysis(scoresList) {
    const aveList = scoresList.map((scores, i) => {
      const sum = scores.reduce((acc, v) => acc + v, 0)
      return sum / scores.length
    })
    
    const srdList = scoresList.map((scores, i) => {
      const sum = scores.reduce((acc, v) => acc + Math.pow(v - aveList[i], 2), 0)
      return Math.sqrt(sum / scores.length)
    })
    
//    const maxList = scoresList.map(s => Math.max(...s))
//    const minList = scoresList.map(s => Math.min(...s))

//    // 下位1%
//    const lowList = scoresList.map(scores => {
//      scores.sort((a, b) => a - b)
//      return scores[parseInt(scores.length/100)]
//    })
    console.log("aveList", aveList, "srdList", srdList)
    return [aveList, srdList]
  }
  
  // useNum: トータルスコアをどこまで見るか？ 同時購入数を設定
  _calcRank(n, isIdecoScores) {
    const selectedNum = isIdecoScores ? idecoPurchaseNum : purchaseNum

    let max = 0, finalRank = 0, rankMax = this._funds.size / 2
    for (let rank=0; rank<rankMax; rank++) {
      const initList = this._getInitList(n, rank)
      let scoresList = []
      this._funds.forEach(fund => {
        if (!isIdecoScores || (isIdecoScores && fund.isIdeco)) {
          scoresList.push(fund.scores[n].map((score, i) => score === null ? initList[i] : score))
        }
      })

      // ラストはトータルスコア
      scoresList = scoresList.map(scores => scores.concat(scores.reduce((acc, v) => acc + v, 0))) // pushは元を上書きするので禁止
      
      for (let i=0; i<scoresList[0].length - 1; i++) {
        let k = selectedNum
        scoresList = scoresList.sort((s1, s2) => s2[i] - s1[i]).map(scores => {
          if (k > 0 && scores[i] !== initList[i]) {
            scores[i] = k // 1 だと、たまに偶然でぶっこわれるケースがあった。変動が激しい。
            k--
          } else {
            scores[i] = 0
          }
          return scores
        })
      }
      scoresList = scoresList.sort((s1, s2) => s2[s1.length - 1] - s1[s1.length - 1]).slice(0, selectedNum)

//      this.logSheet.getRange(1, 1, scoresList.length, scoresList[0].length).setValues(scoresList)

      let sum = 0
      const indicator = scoresList.map(scores => {
        for (let i=0; i<scoresList[0].length - 1; i++) {
          sum += scores[i]
        }
      })
      if (sum > max) {
        finalRank = rank
        max = sum
      }

      const l = [[rank, sum]]
      this.logSheet.getRange(rank + 1, l[0].length * n + 1, 1, l[0].length).setValues(l)
    }

    this.logSheet.getRange(rankMax + 2, 2 * n + 1).setValue(finalRank)
    console.log('_calcRank:rank', finalRank)    

    return finalRank
  }
  
  _getInitList(n, rank) {
    return this._getScoresList(n).map((scores, i) => {
      scores.sort((a, b) => b - a)
      return scores[Math.min(rank, scores.length - 1)]
    })
  }
  
  _fetchCategory() {
    for (let n=0; n<scoresSize; n++) {
      const isIdecoScores = n === 2
      if (isIdecoScores) {
        return
      }
      
      [...this._funds.values()].filter(f => !f.ignore).sort((a, b) => b.totalScores[n] - a.totalScores[n]).slice(0, purchaseNum).forEach(f => this._infoScraper.scraping(f))
    }
    console.log('_fetchCategory')
  }
  
  _output() { 
    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    const categoryCol = 2
    const nameCol = 3
    const isIdecoCol = 4
    const totalScoreCol = 5

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.link, fund.category, fund.name, fund.isIdeco]
      for (let i=0; i<scoresSize; i++) {
        row.push(fund.totalScores[i], ...(fund.scores[i]), '')
      }
      fund.returns.forEach((r, i) => {
        row.push('', r, fund.risks[i], fund.sharps[i])
      })
      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    const lastRow = sheet.getLastRow()

    sheet.autoResizeColumn(categoryCol)
    sheet.autoResizeColumn(nameCol)
    
    for (let i=0; i<scoresSize; i++) {
      sheet.getRange(1, totalScoreCol + i * (termSize + 2), lastRow).setFontWeight("bold")
    }

    let n = 1
    this._funds.forEach(fund => {
      if (fund.ignore) {
        sheet.getRange(n, 1, 1, isIdecoCol + (termSize + 2) * (scoresSize - 1)).setBackground('gray')
      }
      if (fund.isIdeco) {
        sheet.getRange(n, isIdecoCol).setBackground('yellow')
      }
      n++
    })

    const allRange = sheet.getDataRange()
    this._setColors(sheet, allRange, totalScoreCol, nameCol, lastRow)
    allRange.sort({column: totalScoreCol, ascending: false})
    
    sheet.insertRowBefore(1)
    const topRow = ['リンク', 'カテゴリ', '投資信託名称',  'iDeCo']
    for (let i=0; i<scoresSize; i++) {
      topRow.concat('トータルスコア', '6ヶ月', '1年', '3年', '5年', '10年', '')
    }
    const topRowRange = sheet.getRange(1, 1, 1, topRow.length)
    topRowRange.setValues([topRow])
    topRowRange.setBackgrounds([topRow.map(r => 'silver')])
  }

  _setColors(sheet, allRange, totalScoreCol, nameCol, lastRow) {
    const white = '#ffffff' // needs RGB color
    const colors = ['cyan', 'lime', 'yellow', 'orange', '#cfe2f3', '#cfe2f3', 'pink', 'pink', '#F7CEB9', '#F7CEB9', 'silver', 'silver'].concat(new Array(10).fill('white'))
    const colorNum = 5
    
    allRange.sort({column: totalScoreCol, ascending: false})
    const nameRange = sheet.getRange(1, nameCol, lastRow)
    this._setHighRankColor(nameRange)

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
      const range = sheet.getRange(1, i, colorNum * colors.length)
      const rgbs = range.getBackgrounds().map(rows => {
        return rows.map(rgb => {
          if (rgb !== white) {
            return rgb
          }
          const result = colors[parseInt(j / colorNum)]
          j++;

          return result
        })
      })
      range.setBackgrounds(rgbs)
    }
  }

  _setHighRankColor(range) {
    const white = '#ffffff' // needs RGB color
    const aqua = 'aqua'

    let i = 0
    const rgbs = range.getBackgrounds().map(rows => {
      return rows.map(rgb => {
        if (i >= purchaseNum || rgb !== white) {
          return rgb
        }
        i++;
        return aqua
      })
    })
    range.setBackgrounds(rgbs)
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
    fund.category = tds[1] // + '-' + tds[3]
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