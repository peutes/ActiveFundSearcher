class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo() 
    this._infoScraper = new MinkabuInfoScraper()
    
    this._funds = new Map()
    this._noRakutenIgnoreList = [
      "三菱UFJ先進国高金利債券ファンド(毎月決算型)(グローバル・トップ)", "野村世界業種別投資シリーズ(世界半導体株投資)", "野村クラウド関連株式投信Aコース(為替ヘッジあり)", "野村米国ブランド株投資(円コース)毎月分配型", "野村クラウド関連株式投信Bコース(為替ヘッジなし)", "ひふみ投信", "マネックス・日本成長株ファンド(ザ・ファンド@マネックス)", "シンプレクス・ジャパン・バリューアップ・ファンド", "野村米国ブランド株投資(円コース)年2回決算型",
      "UBS中国株式ファンド", "野村米国ブランド株投資(米ドルコース)毎月分配型", "UBS次世代テクノロジー・ファンド", "野村米国ブランド株投資(米ドルコース)年2回決算型", "新光日本小型株ファンド(風物語)", "東京海上Roggeニッポン海外債券ファンド(為替ヘッジあり)", "UBS中国A株ファンド(年1回決算型)(桃源郷)", "USテクノロジー・イノベーターズ・ファンド", "野村米国ブランド株投資(アジア通貨コース)毎月分配型",
      "インデックスファンド海外債券ヘッジあり(DC専用)", "USテクノロジー・イノベーターズ・ファンド(為替ヘッジあり)", "三菱UFJ先進国高金利債券ファンド(年1回決算型)(グローバル・トップ年1)", "US成長株オープン(円ヘッジありコース)", "ダイワ・セレクト日本", "野村SNS関連株投資Aコース", "野村米国ブランド株投資(アジア通貨コース)年2回決算型", "JPMグレーター・チャイナ・オープン", "US成長株オープン(円ヘッジなしコース)", 
      "野村SNS関連株投資Bコース", "ダイワ新興企業株ファンド", "結い2101", "DCインデックスバランス(株式20)", "生活基盤関連株式ファンド(ゆうゆう街道)", "厳選ジャパン", "USストラテジック・インカム・アルファ毎月決算型", "グローバル全生物ゲノム株式ファンド(1年決算型)", "JPM中小型株オープン", "リスク抑制世界8資産バランスファンド(しあわせの一歩)", "ワールド・ウォーター・ファンドAコース",
      "マイストーリー・株25", "三井住友・中国A株・香港株オープン", "ジャパン・アクティブ・グロース(資産成長型)", "ダイワ/バリュー・パートナーズ・チャイナ・イノベーター・ファンド", "年金積立アセット・ナビゲーション・ファンド(株式20)(DCAナビ20)", "ダイワ/ミレーアセット・グローバル・グレートコンシューマー株式ファンド(為替ヘッジあり)", "日本株リーダーズファンド", "DCニッセイワールドセレクトファンド(債券重視型)",
      "京都・滋賀インデックスファンド(京(みやこ)ファンド)", "FANG+インデックス・オープン", "グローバル・プロスペクティブ・ファンド(イノベーティブ・フューチャー)", "三井住友・DC外国債券インデックスファンド", "ジャパン・アクティブ・グロース(分配型)", "ハッピーライフファンド・株25", "浪花おふくろファンド(おふくろファンド)", "フィデリティ世界医療機器関連株ファンド(為替ヘッジなし)", "JPM日本中小型株ファンド",
      "フィデリティ世界医療機器関連株ファンド(為替ヘッジあり)", "先進国投資適格債券ファンド(為替ヘッジあり)(マイワルツ)", "野村アクア投資Aコース", "野村グローバルSRI100(野村世界社会的責任投資)", "三菱UFJライフプラン25(ゆとりずむ25)", "ダイワ債券コア戦略ファンド(為替ヘッジあり)", "米国製造業株式ファンド(USルネサンス)", "グローバル・モビリティ・サービス株式ファンド(1年決算型)(グローバルMaaS(1年決算型))",
      "インデックスファンド海外株式ヘッジあり(DC専用)", "アムンディ・グラン・チャイナ・ファンド", "野村エマージング債券投信(金コース)毎月分配型", "三菱UFJライフプラン50(ゆとりずむ50)", "ダイワ成長株オープン", "野村エマージング債券投信(金コース)年2回決算型", "いちよし公開ベンチャー・ファンド", "ダイワ米国国債7-10年ラダー型ファンド(部分為替ヘッジあり)-USトライアングル-", "ブラックロック・ヘルスサイエンス・ファンド(為替ヘッジあり)",
      "SBI中小型成長株ファンドジェイネクスト(jnext)" // スポットのみ
      // TODO 買える証券別で分ける
    ]
    this._otherIgnoreList = [
      "DIAM新興市場日本株ファンド", "三菱UFJグローバル・ボンド・オープン(毎月決算型)(花こよみ)", "アライアンス・バーンスタイン・米国成長株投信Cコース毎月決算型(為替ヘッジあり)予想分配金提示型", "アライアンス・バーンスタイン・米国成長株投信Dコース毎月決算型(為替ヘッジなし)予想分配金提示型", "MHAM USインカムオープン毎月決算コース(為替ヘッジなし)(ドルBOX)", "ゴールドマン・サックス・世界債券オープンCコース(毎月分配型、限定為替ヘッジ)",
      "GS日本フォーカス・グロース毎月決算コース", "グローバル・フィンテック株式ファンド(年2回決算型)", "グローバル・フィンテック株式ファンド(為替ヘッジあり・年2回決算型)", "グローバル・スマート・イノベーション・オープン(年2回決算型)(iシフト)", "グローバル・スマート・イノベーション・オープン(年2回決算型)為替ヘッジあり(iシフト(ヘッジあり))", "グローバル・ロボティクス株式ファンド(年2回決算型)",
      "あい・パワーファンド(iパワー)", // 業務停止命令で換金もできなくなり危険すぎるため。あと積立購入ができない。
    ]
    this._redemptionDeadlineIgnoreList = []
          
    this.logSheet = this._sheetInfo.getSheet(this._sheetInfo.logSheetName)
    this.logSheet.clear()
  }
      
  calc() {
    this._fetchFunds()
    this._decideScreeningPolicy()
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
        fund.assets = value[3]
        fund.name = value[4]
        fund.ignore = this._noRakutenIgnoreList.some(i => i === fund.name) || this._otherIgnoreList.some(i => i === fund.name) || this._redemptionDeadlineIgnoreList.some(i => i === fund.name)
      
        const n = 6
        for (let i=0; i<minkabuTermSize; i++) {
          const return_ = value[4*i + n + 0]
          const risk = value[4*i + n + 1]
          const sharp = value[4*i + n + 2]
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
  
  _decideScreeningPolicy() {
    console.log('_decideScreeningPolicy')
    this._funds.forEach(fund => {
      fund.returns.forEach((r, i) => {
        if (r === null || fund.risks[i] === null || fund.sharps[i] === null) {
          return
        }
    
//        // ルーキー枠制度 3年未満で1年のデータしか存在しないが優秀なファンドを救済する
//        const rookieFilter = (i === 1 && fund.returns[2] === null ? 1.1 : 1)

        // 純資産信頼性フィルタ（平均化すると邪魔なので外す）
        //const assetsPow = Math.pow(fund.assets, 3)
        //const assetsFilter = assetsPow / (assetsPow + Math.pow(50000, 3))

        // TODO: スクリーニングで早めにやったほうがいいリスト        
        // ・償還期間を見るべき。5年以内に償還するファンドは避けるべき。1年ではなく5年。
        // ・信託財産留保額を見るべき
        // ・NISA口座から外れたら、分配金フィルタを入れるべき。分配金の税金で損するので、25%*20%=5%はダウンするべき。95%計算

        // √を取りたくなるが、√すると絶対値1以下が逆転して、分布が歪になり正規分布でなくなるため使えない・・・
        fund.policy[0][i] = this._calcLowRiskFilter(fund, i, 100) * fund.sharps[i]
        fund.policy[1][i] = this._calcLowRiskFilter(fund, i, 2) * fund.sharps[i] // 比較用
        fund.policy[2][i] = fund.policy[0][i]
      })    
    })
  }

  // 公社債と弱小債券ファンドを除去するためのフィルタ。債券ファンドはリスクが低いため、無駄にシャープレシオが高くなりやすいため傾き補正。また、リターンが飛び抜けてるファンドのインパクトを下げる効果もある。
  _calcLowRiskFilter(fund, i, exp) {
    const f = Math.max(Math.pow(fund.risks[i], exp), Math.pow(fund.risks[i], 1/exp), Math.abs(fund.returns[i]))
    return fund.sharps[i] == 0 ? 0 : f / (f + Math.abs(fund.sharps[i])) // sqrt(sharp) にすると、1年のときはいいが3年5年10年で意味が反転するので良くないため辞める。
  }
  
  _calcMinusScores(fund, i) {
    return fund.returns[i] * fund.risks[i] // returnをlogや√をとると正規分布では無くなる
  }
  
  _calcScores() {
    console.log('_calcScores')
    
    // マイナスの時の正規分布作成用データ
    let rrScoresList = []
    this._funds.forEach(fund => rrScoresList.push(
      fund.returns.map((r, i) => this._calcMinusScores(fund, i))
    ))
    rrScoresList = this._transposeAndFilter(rrScoresList)
    const [_rrAveList, rrSrdList] = this._analysis(rrScoresList)

    for (let n=0; n<calculatedNum; n++) {
      console.log("n", n)
      
      this._funds.forEach(f => f.scores[n] = f.policy[n])
      
      // 各期間ごとのスコアのバランスを整えるために標準化してZスコアを使う
      const [aveList, srdList] = this._analysis(this._getScoresList(n))
      
      this._funds.forEach(fund => {
        fund.scores[n] = fund.scores[n].map((score, i) => {
          if (score === null) {
            return null
          }
          
          // マイナススコア評価：マイナス時はリスクの意味合いがかわり、シャープレシオが使えなくなるため、評価方法を変える。
          // 二つの正規分布を結合して、Zスコアを計算する
          const minusScores = this._calcMinusScores(fund, i)
          const plusRes = (score - aveList[i]) / srdList[i]
          const minusRes = minusScores / rrSrdList[i] - aveList[i] / srdList[i]
          const z = fund.returns[i] >= 0 ? plusRes : minusRes
          
          return i === 0 ? 0 : z
        })
      })

      // 初期値を自動決定するのに各期間のスコアのランキングを使う
//      const rank = this._calcRank(n)
//      const initList = this._getInitList(n, rank)
//      console.log("initList", initList)
      
      this._funds.forEach(fund => {
        // ignore は消さないで、もしも購入可能になったときのスコア表示は残すため、false
        if (this._isTargetFund(n, fund, false)) {
          fund.scores[n] = fund.scores[n].map((s, i) => s !== null ? 10 * s : 0)
          fund.totalScores[n] = fund.scores[n].reduce((acc, score) => acc + score, 0)
        } else {
          fund.scores[n] = fund.scores[n].map(s => null)
          fund.totalScores[n] = null
        }
      })
      
      // 正規分布の信頼区間 95%ゾーンのスコア20以上を購入するのが望ましい -> 購入数は60~70がいいんじゃないかという裏付け
      const totalScores = []
      this._funds.forEach(fund => {
        if (this._isTargetFund(n, fund, false)) {
          totalScores.push(fund.totalScores[n])
        }
      })
      const ave = totalScores.reduce((acc, v) => acc + v, 0) / totalScores.length
      const ave2 = totalScores.reduce((acc, v) => acc + Math.pow(v - ave, 2), 0) / totalScores.length
      const srd = Math.sqrt(ave2)
      this._funds.forEach(fund => {
        if (this._isTargetFund(n, fund, false)) {
          fund.totalScores[n] = 10 * (fund.totalScores[n] - ave) / srd
        }
      })
    }
  }
    
  _isTargetFund(n, fund, useIgnore) {
    const isIdecoScores = this._isIdecoScores(n)
    return (!isIdecoScores && !(useIgnore && fund.ignore)) || (isIdecoScores && fund.isIdeco)    // iDeCoは ignore無視
  }

  _getScoresList(n) {
    const scoresList = []
    this._funds.forEach(fund => scoresList.push(fund.scores[n]))
    return this._transposeAndFilter(scoresList)
  }

  _transposeAndFilter(scoresList) {
    return scoresList[0].map((_, i) => scoresList.map(r => r[i]).filter(Boolean)) // transpose
  }

  _analysis(scoresList) {
    const aveList = scoresList.map((scores, i) => {
      const sum = scores.reduce((acc, v) => acc + v, 0)
      return sum / scores.length
    })
    
    const srdList = scoresList.map((scores, i) => {
      const sum = scores.reduce((acc, v) => acc + Math.pow(v - aveList[i], 2), 0)
      return Math.sqrt(sum / (scores.length - 1))
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

// 研究の結果、デフォルト値は最も総合スコアを高める値は0に収束するという結論が出たためランキング利用は不要に

//  // useNum: トータルスコアをどこまで見るか？ 同時購入数を設定
//  _calcRank(n) {
//    const selectedNum = this._isIdecoScores(n) ? idecoPurchaseNum : purchaseNum
//
//    const rankMax = this._funds.size / 2
//    let max = 0, finalRank = 0
//    const initAdd = Math.pow(10, -10)
//    for (let rank=0; rank<rankMax; rank++) {
//      const initList = this._getInitList(n, rank).map(s => s + initAdd)
//      let scoresList = []
//      this._funds.forEach(fund => {
//        // 最終スコアの調整のため、全てのデータを使わず、ターゲットになってるスコアしか対象にしないため、true
//        // ignoreは初期値の変動に影響する。
//        if (this._isTargetFund(n, fund, true)) {
//          scoresList.push(
//            fund.scores[n].map((s, i) => s !== null ? s : initList[i])
//          )
//        }
//      })
//      
//      // ラストはトータルスコア
//      scoresList = scoresList.map(scores => scores.concat(
//        scores.reduce((acc, v) => acc + v, 0)
//      )) // pushは元を上書きするので禁止
//      
//      // 初期値を設定したときに、初期値「以外」のスコアが最大になるように設定する。
//      for (let i=0; i<scoresList[0].length - 1; i++) {
//        scoresList = scoresList
//          .sort((s1, s2) => s2[i] - s1[i])
//          .map(scores => {
//            if (scores[i] === initList[i]) {
//              scores[i] = 0
//            }
//            return scores
//          })
//      }
//      scoresList = scoresList
//        .sort((s1, s2) => s2[s1.length - 1] - s1[s1.length - 1])
//        .slice(0, selectedNum)
//
//      let sum = 0
//      const indicator = scoresList.map(scores => {
//        for (let i=0; i<scoresList[0].length - 1; i++) {
//          sum += scores[i]
//        }
//      })
//      if (sum > max) {
//        finalRank = rank
//        max = sum
//      }
//
//      const l = [[rank, sum]]
//      this.logSheet.getRange(rank + 1, l[0].length * n + 1, 1, l[0].length).setValues(l)
//    }
//
//    this.logSheet.getRange(rankMax + 2, 2 * n + 1).setValue(finalRank)
//    console.log('_calcRank:rank', finalRank)    
//
//    return finalRank
//  }
//  
//  _getInitList(n, rank) {
//    return this._getScoresList(n).map((scores, i) => {
//      if (scores.length === 0) {
//        return 0
//      }
//      scores.sort((a, b) => b - a)
//      return scores[Math.min(rank, scores.length - 1)]
//    })
//  }
  
  _fetchCategory() {
    for (let n=0; n<calculatedNum; n++) {
      if (n > 1) {
        return
      }
      
      [...this._funds.values()]
        .filter(f => !f.ignore)
        .sort((a, b) => b.totalScores[n] - a.totalScores[n])
        .slice(0, purchaseNum)
        .forEach(f => this._infoScraper.scraping(f))
    }
    console.log('_fetchCategory')
  }

  _isIdecoScores(n) {
    return n === 2
  }
  
  _output() { 
    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    const categoryCol = 4
    const nameCol = 6
    const isIdecoCol = 7
    const totalScoreCol = 8

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.link, fund.ignore, fund.isRakuten, fund.category, fund.assets, fund.name, fund.isIdeco]
      for (let i=0; i<calculatedNum; i++) {
        row.push(fund.totalScores[i], ...(fund.scores[i]), '')
      }
      fund.returns.forEach((r, i) => {
        row.push(r, fund.risks[i], fund.sharps[i], fund.policy[0][i], '')
      })
      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    const lastRow = sheet.getLastRow()

    sheet.autoResizeColumn(categoryCol)
    sheet.autoResizeColumn(nameCol)
    
    for (let i=0; i<calculatedNum; i++) {
      sheet.getRange(1, totalScoreCol + i * (minkabuTermSize + 2), lastRow).setFontWeight("bold")
    }

    let n = 1
    this._funds.forEach(fund => {
      if (fund.ignore) {
        sheet.getRange(n, 1, 1, isIdecoCol + (minkabuTermSize + 2) * (calculatedNum - 1)).setBackground('gray')
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
    const topRow = ['リンク', '無視', '楽天', 'カテゴリ', '純資産額', '投資信託名称',  'iDeCo']
    for (let i=0; i<calculatedNum; i++) {
      topRow.concat('トータルスコア', '6ヶ月', '1年', '3年', '5年', '10年', '')
    }
    const topRowRange = sheet.getRange(1, 1, 1, topRow.length)
    topRowRange.setValues([topRow])
    topRowRange.setBackgrounds([topRow.map(r => 'silver')])
  }

  _setColors(sheet, allRange, totalScoreCol, nameCol, lastRow) {
    const white = '#ffffff' // needs RGB color
    let colors = ['cyan', 'lime', 'yellow', 'orange', '#ea9999', '#ea9999', '#cfe2f3', '#cfe2f3', '#d9ead3', '#d9ead3']
      .concat(new Array(12).fill('white'))
    const colorNum = 5
    
    allRange.sort({column: totalScoreCol, ascending: false})
    const nameRange = sheet.getRange(1, nameCol, lastRow)
    this._setHighRankColor(nameRange)

    for (let i=totalScoreCol; i < totalScoreCol + calculatedNum * (2 + minkabuTermSize) - 1; i++) {
      let c = false
      for (let j=1; j<calculatedNum; j++) {
        if (i === totalScoreCol + j * (minkabuTermSize + 2) - 1) {
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

function calcMinkabuFundsScore() {
  (new MinkabuFundsScoreCalculator).calc()
}