
class FundsDownsideRiskCalculator {
  constructor() {
    this.sheetInfo = new SheetInfo()
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetInfo.downsideRiskSheetName)
    this.sheet.clear()
  }
  
  calc() {
    const link = "https://toushin-lib.fwg.ne.jp/FdsWeb/FDST030000/csv-file-download?isinCd=JP90C000ATW8&associFundCd=39311149"
    const allPrices = this._calcAllPrices(link)
    const risks = this._calcRisks(allPrices)

    // 最終的にrisk1を採用するか、risk2を採用するかは要検討
    risks.forEach((pp, i) => this.output(1, [risks]))
    //allPrices.forEach((pp, i) => this.output((5 * i + 1), pp))

  }
  
  _calcAllPrices(link) {
    const csv = UrlFetchApp
      .fetch(link)
      .getContentText('Shift-JIS')
      .split('\r\n')
      .map(row => row.split(','))
      .filter(a => a.length > 1 && a[0] !== '')

    // 分配金込のリターン計算1
    // TODO: 他に合わせて、月末からのカウントにする

    const now = new Date()
    const endDate = new Date(now.getFullYear(), now.getMonth(), 1)
    endDate.setSeconds(endDate.getSeconds() -1)
    const startDates = (new Array(termSize)).fill().map(_ => new Date(now.getFullYear(), now.getMonth(), 1))
    startDates[0].setMonth(startDates[0].getMonth() - 3)
    startDates[1].setMonth(startDates[1].getMonth() - 6)
    startDates[2].setFullYear(startDates[2].getFullYear() - 1)
    startDates[3].setFullYear(startDates[3].getFullYear() - 3)
    startDates[4].setFullYear(startDates[4].getFullYear() - 5)
    startDates[5].setFullYear(startDates[5].getFullYear() - 10)
    
    console.log(endDate.toLocaleString())
    startDates.forEach(d => console.log(d.toLocaleString()))
      
    const baseColNum = 2
    const prices = (new Array(termSize)).fill().map(_ => [])
    csv.slice(1)
      .map(d => {
        const matched = /(\d+)年(\d+)月(\d+)日/.exec(d[0]);
        const date = new Date(matched[1] - 0, matched[2] - 1, matched[3] - 0)
        return [date, Number(d[1]), Number(d[3])]
      }).filter(d => d[0] <= endDate)
      .forEach((d, i) => {
        startDates.forEach((s, i) => {
          if (s <= d[0] && d[0] <= endDate) {
            prices[i].push([d[2], d[1], d[0]]) // [分配金, 価格, 日付]
          }
        })
      })

      
    const prices2 = prices.map((pp, ii) => {
      let beforeAll = 0
      return pp.map((p, i) => {
        const div = p[0]
        const price = p[1]
        const date = p[2]
        const beforePrice = i === 0 ? price : pp[i - 1][1]
        beforeAll = i === 0 ? price : beforeAll

        const rate = (price + div) / beforePrice // rateの計算に分配金が使われることに注意
        const allPrice = beforeAll * rate
      
        beforeAll = allPrice
        return [allPrice, div, price, date]
      })
    })
    
    return prices2
  }
  
  _calcRisks(allPrices) {
    const risks = allPrices.map(pp => {
      let plus2Sum = 0, minus2Sum = 0
      pp.forEach((p, i) => {
        const diff = i === 0 ? 0 : p[0] - pp[i - 1][0]
        plus2Sum += diff > 0 ? diff * diff : 0
        minus2Sum += diff < 0 ? diff * diff : 0
      })
    
      const upsidePotential = Math.sqrt(plus2Sum / (pp.length - 1))          // アップサイドポテンシャル
      const downsideRisk = Math.sqrt(minus2Sum / (pp.length - 1))            // ダウンサイドリスク
      const risk1 = upsidePotential / (upsidePotential + downsideRisk)       // 0<=n<=1 1が優秀。0が劣等。
      const risk2 = downsideRisk === 0 ? 0 : upsidePotential / downsideRisk  // 大きいほうが優勢。だがこの指標はインフレするかも
      return [upsidePotential, downsideRisk, risk1, risk2, '']
    }).flat()
    
    console.log(risks)
    return risks
  }
  
  output(colNum, data) {
    this.sheet.getRange(1, colNum, data.length, data[0].length).setValues(data)    
  }
}
