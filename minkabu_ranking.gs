function calcMinkabuFundsScoreRanking() {
  (new MinkabuFundsScoreRanking).calc()
}

class MinkabuFundsScoreRanking {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo()

    this._names = new Map()
    this._links = new Map()
    this._categories = new Map()
    this._isIDeCos = new Map()
    this._noIgnores = new Map()
    this._buys = new Map()

    this._scores1 = new Map()
    this._scores3 = new Map()
    this._scores5 = new Map()
    this._scores10 = new Map()

    this._totalScores13510 = new Map()
    this._totalScores135 = new Map()
    this._totalScores1310 = new Map()
    this._totalScores1510 = new Map()
    this._totalScores3510 = new Map()
    this._totalScores13 = new Map()
    this._totalScores15 = new Map()
    this._totalScores110 = new Map()
    this._totalScores35 = new Map()
    this._totalScores310 = new Map()
    this._totalScores510 = new Map()

    this._lightIgnores = [
      '0131200B', '39312065', '3531299B', '03312968',
      '89312121', '89313121', '89314121', '89315121', '0331A172', '96311073',, // iDeCo
    ]
    this._scoreKey = 'score'
  }

  calc() {
    this._sheetInfo.scoreSheetsName.forEach(sheetName => this._getBySheetName(sheetName))
    this._setInfo()

    this._calcScores(this._scores1)
    this._calcScores(this._scores3)
    this._calcScores(this._scores5)
    this._calcScores(this._scores10)
    this._calcTotalScores()

    // 単一期間のスコアは独自の評価が多すぎるため外す。最低でも2つ以上の期間を見る必要がありそう。
    const upperByFirst = 1
    this._calcBuys(this._scores1, upperByFirst)
    this._calcBuys(this._scores3, upperByFirst)
    this._calcBuys(this._scores5, upperByFirst)
    this._calcBuys(this._scores10, upperByFirst)

    const upperBySecond = 1 // 3は多すぎた。
    this._calcBuys(this._totalScores13, upperBySecond)
    this._calcBuys(this._totalScores15, upperBySecond)
    this._calcBuys(this._totalScores110, upperBySecond)
    this._calcBuys(this._totalScores35, upperBySecond)
    this._calcBuys(this._totalScores310, upperBySecond)
    this._calcBuys(this._totalScores510, upperBySecond)

    const upperByThird = 2
    this._calcBuys(this._totalScores13510, upperByThird)
    this._calcBuys(this._totalScores1310, upperByThird)
    this._calcBuys(this._totalScores1510, upperByThird)
    this._calcBuys(this._totalScores3510, upperByThird)

    // まずは3位までで検証
    const upperByForth = 3
    this._calcBuys(this._totalScores135, upperByForth)

    this._setBuysByCategory()
    this._output()
  }

  _getBySheetName(sheetName) {
    const sheet = this._sheetInfo.getSheet(sheetName)
    sheet.getDataRange().getValues().forEach(value => {
      let i=0
      const link = value[i++]
      const name = value[i++]
      i++
      const score1 = value[i++]
      const score3 = value[i++]
      const score5 = value[i++]
      const score10 = value[i++]

      const id = this._getIdByLink(link)
      if (id === null) {
        return
      }

      this._names.set(id, name)
      this._links.set(id, link)
      this._scores1.set(id, this._scores1.get(id) || new Map())
      this._scores3.set(id, this._scores3.get(id) || new Map())
      this._scores5.set(id, this._scores5.get(id) || new Map())
      this._scores10.set(id, this._scores10.get(id) || new Map())

      this._scores1.get(id).set(sheetName, Number(score1))
      this._scores3.get(id).set(sheetName, Number(score3))
      this._scores5.get(id).set(sheetName, Number(score5))
      this._scores10.get(id).set(sheetName, Number(score10))
    })
  }

  _setInfo() {
    const sheetName = this._sheetInfo.scoreSheetsName.slice(-1)[0]
    const sheet = this._sheetInfo.getSheet(sheetName)
    sheet.getDataRange().getValues().forEach(value => {
      let i=0
      const link = value[i]
      i += 8
      const category = value[i++]
      const isIdeco = value[i]
      i += 24
      const noIgnore = value[i++]
      const isRakuten = value[i]
//      console.log(isIdeco, noIgnore, isRakuten)

      const id = this._getIdByLink(link)
      if (id === null) {
        return
      }

      this._categories.set(id, category)
      this._noIgnores.set(id, (noIgnore && isRakuten) || isIdeco)
      this._isIDeCos.set(id, isIdeco)
    })
  }

  _getIdByLink(link) {
    const match = link.match(/https:\/\/itf\.minkabu\.jp\/fund\/(.*)\/risk_cont|https:\/\/itf\.minkabu\.jp\/fund\/(.*)/)
    if (!match) {
      return null
    }
    return match[1] || match[2]
  }

  _calcScores(targetScores) {
    const sums = []
    targetScores.forEach(scores => {
      const sum = Array.from(scores.values()).reduce((acc, s) => acc + s)
      scores.set(this._scoreKey, sum)
      sums.push(sum)
    })

    const ave = this._ave(sums)
    targetScores.forEach(scores => {
      scores.set(this._scoreKey, 10 * (scores.get(this._scoreKey) - ave) / this._srd(sums, ave))
    })
  }

  _calcTotalScores() {
    this._names.forEach((_, id) => {
      this._totalScores13510.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores3.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores135.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores3.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey))
      this._totalScores1310.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores3.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores1510.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores3510.set(id, this._scores3.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores13.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores3.get(id).get(this._scoreKey))
      this._totalScores15.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey))
      this._totalScores110.set(id, this._scores1.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores35.set(id, this._scores3.get(id).get(this._scoreKey) + this._scores5.get(id).get(this._scoreKey))
      this._totalScores310.set(id, this._scores3.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
      this._totalScores510.set(id,this._scores5.get(id).get(this._scoreKey) + this._scores10.get(id).get(this._scoreKey))
    })

    this._calcTotalScore(this._totalScores13510)
    this._calcTotalScore(this._totalScores135)
    this._calcTotalScore(this._totalScores1310)
    this._calcTotalScore(this._totalScores1510)
    this._calcTotalScore(this._totalScores3510)
    this._calcTotalScore(this._totalScores13)
    this._calcTotalScore(this._totalScores15)
    this._calcTotalScore(this._totalScores110)
    this._calcTotalScore(this._totalScores35)
    this._calcTotalScore(this._totalScores310)
    this._calcTotalScore(this._totalScores510)
  }

  _calcTotalScore(targetTotalScores) {
    const totalScores = Array.from(targetTotalScores.values())
    const ave = this._ave(totalScores)
    targetTotalScores.forEach((score, id) => {
      targetTotalScores.set(id, new Map([[this._scoreKey, 10 * (score - ave) / this._srd(totalScores, ave)]]))
    })
  }

  _calcBuys(targetScores, upper) {
    upper--
    const sortScores = this._changeSortedScores(targetScores)
    sortScores.forEach((e, i) => {
      if (this._lightIgnores.includes(e['id'])){
        upper++
      }
      if (i <= upper) {
        this._setBuys(e['id'])
      }
    })
  }

  _setBuysByCategory() {
    let sortedScores
    const asiaCategories = ['国際株式型-アジアオセアニア株式型', '国際株式型-エマージング株式型']
    const goldCategory = '国際株式型-国際資源関連株式型'
    let asiaSelectedNum = 2
    let goldSelectedNum = 1
    let bondSelectedNum = 2

    sortedScores = this._changeSortedScores(this._totalScores135)
    sortedScores.forEach(e => {
      if (asiaSelectedNum > 0 && asiaCategories.includes(this._categories.get(e['id']))) {
        this._setBuys(e['id'])
        asiaSelectedNum--
        return
      }

      if (goldSelectedNum > 0 && this._categories.get(e['id']) === goldCategory) {
        this._setBuys(e['id'])
        goldSelectedNum--
        return
      }

      if (bondSelectedNum > 0 && this._categories.get(e['id']).indexOf('株式型') == -1) {
        this._setBuys(e['id'])
        bondSelectedNum--
        return
      }
    })
  }

  _changeSortedScores(targetScores) {
    // Map型からObject型に変換する
    return Array.from(targetScores.entries())
      .map(e => {
        return {'id': e[0], [this._scoreKey]: e[1].get(this._scoreKey)}
      })
      .filter(e => this._noIgnores.get(e['id']))
      .sort((a, b) => b[this._scoreKey] - a[this._scoreKey])
  }

  _setBuys(id) {
    this._buys.set(id, (this._buys.get(id) || 0 ) + 1)
  }

  _ave(scores) {
    return scores.reduce((acc, v) => acc + v, 0) / scores.length
  }

  _srd(scores, ave) {
    return Math.sqrt(scores.reduce((acc, v) => acc + Math.pow(v - ave, 2), 0) / (scores.length - 1))
  }

  _output() {
    const data = []
    const lightIgnoreCols = []
    let col = 0
    this._names.forEach((name, id) => {
      if (this._noIgnores.get(id)) {
        data.push([
          id,
          this._categories.get(id),
          name,
          this._buys.get(id),
          this._buys.get(id) > 0 ? '◯' : '',
          this._isIDeCos.get(id),
          this._totalScores135.get(id).get(this._scoreKey),
          this._scores1.get(id).get(this._scoreKey),
          this._scores3.get(id).get(this._scoreKey),
          this._scores5.get(id).get(this._scoreKey),
          this._scores10.get(id).get(this._scoreKey),
          this._totalScores13510.get(id).get(this._scoreKey),
          this._totalScores1310.get(id).get(this._scoreKey),
          this._totalScores1510.get(id).get(this._scoreKey),
          this._totalScores3510.get(id).get(this._scoreKey),
          this._totalScores13.get(id).get(this._scoreKey),
          this._totalScores15.get(id).get(this._scoreKey),
          this._totalScores110.get(id).get(this._scoreKey),
          this._totalScores35.get(id).get(this._scoreKey),
          this._totalScores310.get(id).get(this._scoreKey),
          this._totalScores510.get(id).get(this._scoreKey),
        ])

        col++
        if (this._lightIgnores.includes(id)){
          lightIgnoreCols.push(col)
        }
      }
    })

    const sheet = this._sheetInfo.insertSheet(this._sheetInfo.getScoreSheetName())
    sheet.clear()
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
    lightIgnoreCols.forEach(col => sheet.getRange(col, 1, 1, 4).setBackground('gray'))

    sheet.insertRowBefore(1)
    const topRow = [
      'ID', 'カテゴリ', 'ファンド名', '購入', 'セレクト', 'iDeCo', '135トータル', '1年スコア', '3年スコア', '5年スコア', '10年スコア', '13510トータル', '1310トータル', '1510トータル', '3510トータル', '13トータル', '15トータル', '110トータル', '35トータル', '310トータル', '510トータル',
    ]
    const topRowRange = sheet.getRange(1, 1, 1, topRow.length)
    topRowRange.setValues([topRow])
    topRowRange.setBackgrounds([topRow.map(_ => 'silver')])

    sheet.autoResizeColumn(2)
    sheet.autoResizeColumn(3)
    const allDataRange = sheet.getDataRange()
    allDataRange.createFilter()
    allDataRange.sort({column: 7, ascending: false})
    console.log('_output')
  }
}