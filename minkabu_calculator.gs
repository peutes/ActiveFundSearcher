class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo() 
    
    this._funds = []
    this._otherIgnoreList = [
      // 定期購入ができない系
      "DIAM新興市場日本株ファンド",
      "あい・パワーファンド(iパワー)", // 業務停止命令で換金もできなくなり危険すぎるため。あと積立購入ができない
      "SBI中小型成長株ファンドジェイネクスト(jnext)", // スポット購入オンリー
      "SBI中小型割安成長株ファンドジェイリバイブ(jrevive)", // スポット購入オンリー

      // 分配金系
      "三菱UFJグローバル・ボンド・オープン(毎月決算型)(花こよみ)",
      "アライアンス・バーンスタイン・米国成長株投信Cコース毎月決算型(為替ヘッジあり)予想分配金提示型",
      "アライアンス・バーンスタイン・米国成長株投信Dコース毎月決算型(為替ヘッジなし)予想分配金提示型",
      "グローバル・フィンテック株式ファンド(年2回決算型)",
      "グローバル・フィンテック株式ファンド(為替ヘッジあり・年2回決算型)",
      "グローバル・スマート・イノベーション・オープン(年2回決算型)(iシフト)",
      "グローバル・スマート・イノベーション・オープン(年2回決算型)為替ヘッジあり(iシフト(ヘッジあり))",
    ]
          
    this.logSheet = this._sheetInfo.getSheet(this._sheetInfo.logSheetName)
    this.logSheet.clear()
  }
      
  calc() {
    this._fetchFunds()
    this._decideScreeningPolicy()
    this._calsScores()
    this._calcTotalScores()
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

        const fund = new Fund(null, null)
        const n = 4 // return + risk + sharpRatio + 空白 の1セット
        for (let i=0; i<minkabuTermSize; i++) {
          const return_ = value[n*i]
          const risk = value[n*i + 1]
          const sharp = value[n*i + 2]
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

        let i = n * minkabuTermSize;
        fund.link = value[i++]
        fund.isIdeco = value[i++]
        fund.name = value[i++]
        fund.management = value[i++]
        fund.rating = value[i++]
        fund.assets = value[i++]
        fund.salesCommission = value[i++]
        fund.fee = value[i++]
        fund.redemptionFee = value[i++]
        fund.category = value[i++]
        fund.launchDate = value[i++]
        fund.commissionPeriod = value[i++]
        i++;

        fund.isSaledRakuten = value[i++]
        fund.isSaledSbi = value[i++]
        fund.isSaledAu = value[i++]
        fund.isSaledMonex = value[i++]
        fund.isSaledMatsui = value[i++]
        fund.isSaledNomura = value[i++]
        fund.isSaledDaiwa = value[i++]
        fund.isSaledSmbc = value[i++]
        fund.isSaledLine = value[i++]
        fund.isSaledOkasan = value[i++]
        fund.isSaledPayPay = value[i++]
        fund.isSaledMitsui = value[i++]
        fund.isSaledUfj = value[i++]
        fund.isSaledYutyo = value[i++]
        fund.isSaledMizuho = value[i++]

        fund.ignore = this._otherIgnoreList.some(i => i === fund.name)
      
        this._funds.push(fund)
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
        fund.policy[i] = this._calcLowRiskFilter(fund, i, 100) * fund.sharps[i]
      })    
    })
  }

  // 公社債と弱小債券ファンドを除去するためのフィルタ。債券ファンドはリスクが低いため、無駄にシャープレシオが高くなりやすいため傾き補正。また、リターンが飛び抜けてるファンドのインパクトを下げる効果もある。
  _calcLowRiskFilter(fund, i, exp) {
    const f = Math.max(Math.pow(fund.risks[i], exp), Math.pow(fund.risks[i], 1/exp), Math.abs(fund.returns[i]))
    return fund.sharps[i] == 0 ? 0 : f / (f + Math.abs(fund.sharps[i])) // sqrt(sharp) にすると、1年のときはいいが3年5年10年で意味が反転するので良くないため辞める。
  }
  
  _calsScores() {
    console.log('_calsScores')

    // マイナスの時の正規分布作成用データ
    const rrScoresList = this._funds.map(f => f.returns.map((_, i) => this._calcMinusScores(f, i)))
    const rrSrdList = this._transposeAndFilter(rrScoresList).map(s => this._srd(s, this._ave(s)))
    
    // 各期間ごとのスコアのバランスを整えるために標準化してZスコアを使う
    const scoresList = this._transposeAndFilter(this._funds.map(f => f.policy))
    const aveList = scoresList.map(this._ave)
    const srdList = scoresList.map((s, i) => this._srd(s, aveList[i]))
    console.log("aveList", aveList, "srdList", srdList)
    
    this._funds.forEach(fund => {
      fund.scores = fund.policy.map((policy, i) => {
        if (policy === null) {
          return null
        }
        
        // マイナススコア評価：マイナス時はリスクの意味合いがかわり、シャープレシオが使えなくなるため、評価方法を変える。
        // 二つの正規分布を結合して、Zスコアを計算する
        const minusScores = this._calcMinusScores(fund, i)
        const plusRes = (policy - aveList[i]) / srdList[i]
        const minusRes = minusScores / rrSrdList[i] - aveList[i] / srdList[i]
        const score = fund.returns[i] >= 0 ? plusRes : minusRes
        return 10 * score
      })
    })
  }

  _calcMinusScores(fund, i) {
    return fund.returns[i] * fund.risks[i] // returnをlogや√をとると正規分布では無くなる
  }

  _calcTotalScores() {
    console.log('_calcTotalScores')

    let totalScores = this._funds.map(f => this._calcTotalScoresByRange(f.scores, 0, 3))
    let ave = this._ave(totalScores)
    this._funds.forEach((f, i) => f.totalScoresOf10531 = 10 * (totalScores[i] - ave) / this._srd(totalScores, ave))

    totalScores = this._funds.map(f => this._calcTotalScoresByRange(f.scores, 0, 2))
    ave = this._ave(totalScores)
    this._funds.forEach((f, i) => f.totalScoresOf531 = 10 * (totalScores[i] - ave) / this._srd(totalScores, ave))

    totalScores = this._funds.map(f => this._calcTotalScoresByRange(f.scores, 1, 2))
    ave = this._ave(totalScores)
    this._funds.forEach((f, i) => f.totalScoresOf53 = 10 * (totalScores[i] - ave) / this._srd(totalScores, ave))

    totalScores = this._funds.map(f => this._calcTotalScoresByRange(f.scores, 0, 1))
    ave = this._ave(totalScores)
    this._funds.forEach((f, i) => f.totalScoresOf31 = 10 * (totalScores[i] - ave) / this._srd(totalScores, ave))
  }

  _calcTotalScoresByRange(scores, start, end) {
      return [...Array(end + 1).keys()].slice(start).reduce((acc, i) => acc + Number(scores[i]))
  }

  _transposeAndFilter(scoresList) {
    return scoresList[0].map((_, i) => scoresList.map(r => r[i]).filter(Boolean)) // transpose
  }

  _ave(scores) {
    return scores.reduce((acc, v) => acc + v, 0) / scores.length
  }

  _srd(scores, ave) {
    return Math.sqrt(scores.reduce((acc, v) => acc + Math.pow(v - ave, 2), 0) / (scores.length - 1))
  }
  
  _output() { 
    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    const data = []
    this._funds.forEach(fund => {
      const canBuy = fund.isSaledRakuten && !fund.ignore
      const row = [
        fund.link, fund.name, fund.totalScoresOf10531, ...fund.scores, '',
        fund.category, canBuy, fund.isIdeco, fund.management, fund.rating, fund.assets, fund.salesCommission, fund.fee, fund.redemptionFee, fund.commissionPeriod,
      ]

      fund.returns.forEach((r, i) => {
        row.push(
          r, fund.risks[i], fund.sharps[i], fund.policy[i])
      })

      row.push(
        !fund.ignore, fund.isSaledRakuten, fund.isSaledSbi, fund.isSaledAu, fund.isSaledMonex, fund.isSaledMatsui, fund.isSaledNomura, fund.isSaledDaiwa,
        fund.isSaledSmbc, fund.isSaledLine, fund.isSaledOkasan, fund.isSaledPayPay, fund.isSaledMitsui, fund.isSaledUfj, fund.isSaledYutyo, fund.isSaledMizuho,
      )

      data.push(row)
    })
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    sheet.getDataRange().sort({column: 3, ascending: false})

    sheet.insertRowBefore(1)
    const topRow = [
      'リンク', 'ファンド名', '10531トータル', '1年スコア', '3年スコア', '5年スコア', '10年スコア', '',
      'カテゴリ', '購入可', 'iDeCo', '運営会社', 'レーティング', '資産', '販売手数料', '信託報酬', '信託財産留保額', '信託期間',
      '1年リターン', '1年リスク', '1年シャープ', '1年ポリシー', '3年リターン', '3年リスク', '3年シャープ', '3年ポリシー',
      '5年リターン', '5年リスク', '5年シャープ', '5年ポリシー', '10年リターン', '10年リスク', '10年シャープ', '10年ポリシー',
      'NO無視', '楽天', 'SBI', 'au', 'マネックス', '松井', '野村', '大和', 'SMBC', 'LINE', '岡三', 'PayPay', '三井', 'UFJ', 'ゆうちょ', 'みずほ',
    ]
    const topRowRange = sheet.getRange(1, 1, 1, topRow.length)
    topRowRange.setValues([topRow])
    topRowRange.setBackgrounds([topRow.map(_ => 'silver')])

    sheet.autoResizeColumn(1)
    sheet.autoResizeColumn(2)
    sheet.getDataRange().createFilter()
  }

  // _setColors(sheet, allRange, totalScoreCol, nameCol, lastRow) {
  //   const white = '#ffffff' // needs RGB color
  //   let colors = ['cyan', 'lime', 'yellow', 'orange', '#ea9999', '#ea9999', '#cfe2f3', '#cfe2f3', '#d9ead3', '#d9ead3']
  //     .concat(new Array(12).fill('white'))
  //   const colorNum = 5
    
  //   allRange.sort({column: totalScoreCol, ascending: false})
  //   const nameRange = sheet.getRange(1, nameCol, lastRow)

  //   for (let i=totalScoreCol; i < totalScoreCol + calculatedNum * (2 + minkabuTermSize) - 1; i++) {
  //     let c = false
  //     for (let j=1; j<calculatedNum; j++) {
  //       if (i === totalScoreCol + j * (minkabuTermSize + 2) - 1) {
  //         c = true
  //       }
  //     }
  //     if (c) {
  //       continue
  //     }
    
  //     allRange.sort({column: i, ascending: false})
      
  //     let j = 0
  //     const range = sheet.getRange(1, i, colorNum * colors.length)
  //     const rgbs = range.getBackgrounds().map(rows => {
  //       return rows.map(rgb => {
  //         if (rgb !== white) {
  //           return rgb
  //         }
  //         const result = colors[parseInt(j / colorNum)]
  //         j++;

  //         return result
  //       })
  //     })
  //     range.setBackgrounds(rgbs)
  //   }
  // }
}

function calcMinkabuFundsScore() {
  (new MinkabuFundsScoreCalculator).calc()
}