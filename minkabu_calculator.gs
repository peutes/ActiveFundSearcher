class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo() 
    
    this._funds = []
    this._otherIgnoreList = [
      "アライアンス・バーンスタイン・米国成長株投信Dコース毎月決算型(為替ヘッジなし)予想分配金提示型",
      "ブラックロック・ヘルスサイエンス・ファンド(為替ヘッジなし)",
      "ブラックロック・ヘルスサイエンス・ファンド(為替ヘッジなし/年4回決算型)"
    ]
  }
      
  calc() {
    console.log("debug1")
    this._fetchFunds()
    console.log("debug2")
    this._decideScreeningPolicy()
    console.log("debug3")
    this._calcScores()
    console.log("debug4")
    this._calcTotalScores()
    console.log("debug5")
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
        fund.salesCommission = Number(value[i++])
        fund.fee = Number(value[i++])
        fund.redemptionFee = Number(value[i++])
        fund.category = value[i++]
        fund.launchDate = value[i++]
        fund.commissionPeriod = value[i++]
        i++;

        fund.isSaledRakuten = value[i++] // ノーロード
        fund.isSaledSbi = value[i++]　// ノーロード
        fund.isSaledAu = value[i++]　// ノーロード
        fund.isSaledMonex = value[i++]　// ノーロード
        fund.isSaledMatsui = value[i++]　// ノーロード
        fund.isSaledNomura = value[i++] // ごくわずかしかノーロードは無いので意味無し
        fund.isSaledDaiwa = value[i++] // 一部ノーロード
        fund.isSaledSmbc = value[i++]  // 一部ノーロード
        fund.isSaledLine = value[i++]　// ノーロード
        fund.isSaledOkasan = value[i++]　// ノーロード
        fund.isSaledPayPay = value[i++]　// ノーロード
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
        if (r === null || fund.risks[i] === 0 || fund.risks[i] === null || fund.sharps[i] === null) {
          return
        }
    
        // 純資産信頼性フィルタ（平均化すると邪魔なので外す）
        //const assetsPow = Math.pow(fund.assets, 3)
        //const assetsFilter = assetsPow / (assetsPow + Math.pow(50000, 3))

        // TODO: スクリーニングで早めにやったほうがいいリスト        
        // ・NISA口座から外れたら、分配金フィルタを入れるべき。分配金の税金で損するので、25%*20%=5%はダウンするべき。95%計算

        // √を取りたくなるが、√すると絶対値1以下が逆転して、分布が歪になり正規分布でなくなるため使えない・・・
        const sharp = fund.sharps[i] - fund.redemptionFee / fund.risks[i]
       
        fund.policy[i] = this._calcLowRiskFilter(fund, i) * sharp // 100は公社債投信が除去できなかった
      })
    })
  }

  // 公社債と弱小債券ファンドを除去するためのフィルタ。債券ファンドはリスクが低いため、無駄にシャープレシオが高くなりやすいため傾き補正。また、リターンが飛び抜けてるファンドのインパクトを下げる効果もある。
  _calcLowRiskFilter(fund, i) {
    //let f = fund.risks[i] * 0.2 // リターンが低い投資信託を下げる。
    let f = Math.max(fund.risks[i] - minusRisk, 0) // リターンが低い投資信託を下げる。
    
    return fund.sharps[i] == 0 ? 0 : f / (f + Math.abs(fund.sharps[i])) // sqrt(sharp) にすると、1年のときはいいが3年5年10年で意味が反転するので良くないため辞める。
  }
  
  _calcScores() {
    console.log('_calcScores')

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
    if(scores.length > 0) {
      return scores.reduce((acc, v) => acc + v, 0) / scores.length
    } else {
      return 0
    }
  }

  _srd(scores, ave) {
    if (scores.length > 0) {
      return Math.sqrt(scores.reduce((acc, v) => acc + Math.pow(v - ave, 2), 0) / (scores.length - 1))
    } else {
      return 0
    }
  }
  
  _output() {
    console.log('_output')

    const data = []
    this._funds.forEach(fund => {
      console.log(fund.name, fund.totalScoresOf53, fund.scores)
      
      const row = [
        fund.link, fund.name, fund.totalScoresOf53, ...fund.scores, '',
        fund.category, fund.isIdeco, fund.management, fund.rating, fund.assets, fund.salesCommission, fund.fee, fund.redemptionFee, fund.commissionPeriod,
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
    console.log(data.length, data[0].length)
    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    sheet.insertRows(1, 3000)
    sheet.insertColumnsAfter(1, 20);
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

    // 5年スコアでソートすることに決定！
    sheet.getDataRange().sort({column: 6, ascending: false})

    sheet.insertRowBefore(1)
    const topRow = [
      'リンク', 'ファンド名', '53トータル', '1年スコア', '3年スコア', '5年スコア', '10年スコア', '',
      'カテゴリ', 'iDeCo', '運営会社', 'レーティング', '資産', '販売手数料', '信託報酬', '信託財産留保額', '信託期間',
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
}

function calcMinkabuFundsScore() {
  (new MinkabuFundsScoreCalculator).calc()
}
