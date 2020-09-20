class MinkabuFundsScoreCalculator {
  constructor() {
    this._sheetInfo = new MinkabuSheetInfo() 
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
      "コーポレート・ボンド・インカム(為替ヘッジ型)(泰平航路)", "パインブリッジ世界国債インカムオープン「毎月タイプ」(スーパーシート)", "ニッセイ欧州株式厳選ファンドリスクコントロールコース",
      "三井住友・中国A株・香港株オープン", "ワールド・ゲノムテクノロジー・オープンBコース", "ワールド・ゲノムテクノロジー・オープンAコース", "リビング・アース戦略ファンド(年2回決算コース)",
      "ブラックロック・ゴールド・メタル・オープンAコース", "ブラックロック・ゴールド・メタル・オープンBコース", "ダイワ投信倶楽部外国債券インデックス", "グローバル・ソブリン・オープン（毎月決算型）",
      "リビング・アース戦略ファンド(年2回決算コース)", "グローバル・セキュリティ株式ファンド(3ヵ月決算型)", "世界分散投資戦略ファンド(グローバル・ビュー)", "野村グローバルAI関連株式ファンドAコース",
      "野村グローバルAI関連株式ファンドBコース", "DCニッセイワールドセレクトファンド(安定型)", "ブラックロック・ヘルスサイエンス・ファンド(為替ヘッジなし)", "ブラックロック・ヘルスサイエンス・ファンド(為替ヘッジあり)",
      "シンプレクス・ジャパン・バリューアップ・ファンド", "ファイン・ブレンド(毎月分配型)", "ワールド・ウォーター・ファンドAコース", "ワールド・ウォーター・ファンドBコース",
      "三井住友・NYダウ・ジョーンズ指数オープン(為替ヘッジあり)", "三井住友・NYダウ・ジョーンズ・インデックスファンド(為替ヘッジ型)(NYドリーム)", "三井住友・NYダウ・ジョーンズ指数オープン(為替ヘッジなし)",
      "三井住友・NYダウ・ジョーンズ・インデックスファンド(為替ノーヘッジ型)(NYドリーム)", "野村アクア投資Aコース", "野村アクア投資Bコース", "ひふみ投信", "三菱UFJ国内債券インデックスファンド(確定拠出年金)",
      "シンプレクス・ジャパン・バリューアップ・ファンド", "スマート・ファイブ(1年決算型)", "スマート・ファイブ(毎月決算型)", "大和住銀日本小型株ファンド", "DCインデックスバランス(株式20)",
      "DCインデックスバランス(株式40)", "DCインデックスバランス(株式60)", "DCインデックスバランス(株式80)", "シンプレクス・ジャパン・バリューアップ・ファンド",
      "年金積立アセット・ナビゲーション・ファンド(株式20)(DCAナビ20)", "年金積立アセット・ナビゲーション・ファンド(株式40)(DCAナビ40)", "年金積立アセット・ナビゲーション・ファンド(株式60)(DCAナビ60)",
      "年金積立アセット・ナビゲーション・ファンド(株式80)(DCAナビ80)", "野村世界業種別投資シリーズ(世界ヘルスケア株投資)", "野村世界業種別投資シリーズ(世界金融株投資)",
      "野村世界6資産分散投信(安定コース)", "野村世界6資産分散投信(分配コース)", "野村世界6資産分散投信(成長コース)", "野村世界6資産分散投信(配分変更コース)", "いちよし公開ベンチャー・ファンド",
      "マイストーリー・株25", "マイストーリー・株50", "マイストーリー・株75", "マイストーリー・株100", "ハッピーライフファンド・株25", "ハッピーライフファンド・株50", "ハッピーライフファンド・株100",
      "かいたくファンド", "三菱UFJライフプラン25(ゆとりずむ25)", "三菱UFJライフプラン50(ゆとりずむ50)", "三菱UFJライフプラン75(ゆとりずむ75)", "MHAMライフナビゲーションインカム(ライフナビインカム)",
      "米国優先証券オープン", "バランスセレクト30", "バランスセレクト50", "バランスセレクト70", "ダイワライフスタイル25", "ダイワライフスタイル50", "ダイワライフスタイル75",
      "ワールド・ウォーター・ファンドAコース", "ワールド・ウォーター・ファンドBコース", "結い2101", "ピムコ世界債券戦略ファンド(毎月決算型)Aコース(為替ヘッジあり)", 
      "ピムコ世界債券戦略ファンド(毎月決算型)Bコース(為替ヘッジなし)", "米国国債ファンド為替ヘッジなし(毎月決算型)", "ロボット・テクノロジー関連株ファンド-ロボテック-",
      "ダイワ債券コア戦略ファンド(為替ヘッジなし)", "ダイワ・セレクト日本", "アムンディ・中国株ファンド(悟空)", "大和住銀中国株式ファンド", "BNYメロン・米国株式ダイナミック戦略ファンド(亜米利加)",
      "先進国ハイクオリティ成長株式ファンド(為替ヘッジあり)(未来の世界(先進国))", "先進国ハイクオリティ成長株式ファンド(為替ヘッジなし)(未来の世界(先進国))", "厳選ジャパン", "日興AM中国A株ファンド2(黄河2)",
      "ニッセイSDGsグローバルセレクトファンド(年2回決算型・為替ヘッジあり)", "ティー・ロウ・プライス世界厳選成長株式ファンドAコース(資産成長型・為替ヘッジあり)",
      "ティー・ロウ・プライス世界厳選成長株式ファンドCコース(分配重視型・為替ヘッジあり)", "ティー・ロウ・プライス世界厳選成長株式ファンドDコース(分配重視型・為替ヘッジなし)",
      "ティー・ロウ・プライス世界厳選成長株式ファンドBコース(資産成長型・為替ヘッジなし)", "UBS中国A株ファンド(年4回決算型)(桃源郷・年4)", "米国製造業株式ファンド(USルネサンス)",
      "米国厳選成長株集中投資ファンドBコース(為替ヘッジなし)(新世紀アメリカ~Yes,We Can!~)", "ニッセイ次世代医療ファンド", "ダイワ/ミレーアセット・グローバル・グレートコンシューマー株式ファンド(為替ヘッジあり)",
      "野村グローバルSRI100(野村世界社会的責任投資)", "ヘッジ付先進国株式インデックスオープン", "インデックスファンド海外株式ヘッジあり(DC専用)", "アジアオープン", "ダイワ新興企業株ファンド",
      "きらめきジャパン(きらめき)", "アムンディ・りそなグローバル・ブランド・ファンド(ティアラ)", "J-REITオープン(年4回決算型)", "JPM日本中小型株ファンド", "6資産バランスファンド(成長型)(ダブルウイング)",
      "グローバル・ボンド・ポート毎月決算コース(為替ヘッジなし)", "DIAMグローバル・ボンド・ポート毎月決算コース2(ぶんぱいくん)", "生活基盤関連株式ファンド(ゆうゆう街道)",
      "アムンディ・グラン・チャイナ・ファンド", "浪花おふくろファンド(おふくろファンド)", 
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
      
        for (let i=0; i<minkabuTermSize; i++) {
          const return_ = value[4*i + minkabuTermSize]
          const risk = value[4*i + minkabuTermSize + 1]
          const sharp = value[4*i + minkabuTermSize + 2]
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

        // 公社債と弱小債券ファンドを除去するためのフィルタ。債券ファンドはリスクが低いため、無駄にシャープレシオが高くなりやすい。また、リターンが飛び抜けてるファンドのインパクトを下げる効果もある。債券ファンドはデフォルトだとゴミがあまりにも多すぎる。。。
        // このルートを取ると結構強力に債券ファンドが一掃される。
//        const filter = fund.risks[i] === 0 ? 0 : Math.sqrt(fund.risks[i] / (fund.risks[i] + Math.abs(fund.sharps[i]))) * Math.sqrt(fund.risks[i] / (fund.risks[i] + Math.abs(r))) // シャープレシオは1以下も多いため、sqrt(fund.sharps[i])すると逆に増加するので注意
//        const filter2 = fund.risks[i] === 0 ? 0 : Math.sqrt(fund.risks[i] / (fund.risks[i] + Math.abs(fund.sharps[i]))) * Math.sqrt(fund.risks[i] * fund.risks[i] / (fund.risks[i] * fund.risks[i] +  Math.abs(r) * Math.abs(r))) // シャープレシオは1以下も多いため、sqrt(fund.sharps[i])すると逆に増加するので注意
//        const filter = fund.risks[i] === 0 ? 0 : Math.sqrt(fund.risks[i] / (fund.risks[i] + Math.sqrt(Math.abs(fund.sharps[i]))))

        const w = 0.3  // シャープレシオの副作用。リターンとリスクがあまりにも小さすぎるのを除去。公社債投信をランク外へ排除。
        const filter = Math.sqrt(
          Math.abs(r) * fund.risks[i] / ((Math.abs(r) + w) * (fund.risks[i] + w))
        ) // グラフの形状的にフィルタとしてlogより√が適任
        
        // √を取りたくなるが、√すると絶対値1以下が逆転して、分布が歪になり正規分布でなくなるため使えない・・・
        fund.scores[0][i] = fund.sharps[i] * filter
        fund.scores[1][i] = fund.sharps[i] * filter
        fund.scores[2][i] = fund.scores[0][i]
      })
    })
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

    for (let n=0; n<scoresSize; n++) {
      console.log("n", n)

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
          
          // ルーキー枠制度。6ヶ月のデータを重み付け加算
//          if (n === 1) {
//            const y = fund.scores[n][2] === null ? 1 : (fund.scores[n][3] === null ? 3 : (fund.scores[n][4] === null ? 5 : 10))
//            return i === 0 ? z / (2 * y) : z
//          }
          return i === 0 ? 0 : z      
        })
      })

      // 初期値を自動決定するのに各期間のスコアのランキングを使う
      const rank = this._calcRank(n)
      const initList = this._getInitList(n, rank)
      console.log("initList", initList)
      
      this._funds.forEach(fund => {
        // ignore は消さないで、もしも購入可能になったときのスコア表示は残すため、false
        if (this._isTargetFund(n, fund, false)) {
          fund.scores[n] = fund.scores[n].map((s, i) => 10 * (s !== null ? s : initList[i]))
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
  
  // useNum: トータルスコアをどこまで見るか？ 同時購入数を設定
  _calcRank(n) {
    const selectedNum = this._isIdecoScores(n) ? idecoPurchaseNum : purchaseNum

    const rankMax = this._funds.size / 2
    let max = 0, finalRank = 0
    const initAdd = Math.pow(10, -10)
    for (let rank=0; rank<rankMax; rank++) {
      const initList = this._getInitList(n, rank).map(s => s + initAdd)
      let scoresList = []
      this._funds.forEach(fund => {
        // 最終スコアの調整のため、全てのデータを使わず、ターゲットになってるスコアしか対象にしないため、true
        // ignoreは初期値の変動に影響する。
        if (this._isTargetFund(n, fund, true)) {
          scoresList.push(
            fund.scores[n].map((s, i) => s !== null ? s : initList[i])
          )
        }
      })
      
      // ラストはトータルスコア
      scoresList = scoresList.map(scores => scores.concat(
        scores.reduce((acc, v) => acc + v, 0)
      )) // pushは元を上書きするので禁止
      
      // 初期値を設定したときに、初期値「以外」のスコアが最大になるように設定する。
      for (let i=0; i<scoresList[0].length - 1; i++) {
        scoresList = scoresList
          .sort((s1, s2) => s2[i] - s1[i])
          .map(scores => {
            if (scores[i] === initList[i]) {
              scores[i] = 0
            }
            return scores
          })
      }
      scoresList = scoresList
        .sort((s1, s2) => s2[s1.length - 1] - s1[s1.length - 1])
        .slice(0, selectedNum)

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
      if (scores.length === 0) {
        return 0
      }
      scores.sort((a, b) => b - a)
      return scores[Math.min(rank, scores.length - 1)]
    })
  }
  
  _fetchCategory() {
    for (let n=0; n<scoresSize; n++) {
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
    const nameCol = 5
    const isIdecoCol = 6
    const totalScoreCol = 7

    const data = []
    this._funds.forEach(fund => {
      const row = [fund.link, fund.ignore, fund.isRakuten, fund.category, fund.name, fund.isIdeco]
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
      sheet.getRange(1, totalScoreCol + i * (minkabuTermSize + 2), lastRow).setFontWeight("bold")
    }

    let n = 1
    this._funds.forEach(fund => {
      if (fund.ignore) {
        sheet.getRange(n, 1, 1, isIdecoCol + (minkabuTermSize + 2) * (scoresSize - 1)).setBackground('gray')
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
    const topRow = ['リンク', '無視', '楽天', 'カテゴリ', '投資信託名称',  'iDeCo']
    for (let i=0; i<scoresSize; i++) {
      topRow.concat('トータルスコア', '6ヶ月', '1年', '3年', '5年', '10年', '')
    }
    const topRowRange = sheet.getRange(1, 1, 1, topRow.length)
    topRowRange.setValues([topRow])
    topRowRange.setBackgrounds([topRow.map(r => 'silver')])
  }

  _setColors(sheet, allRange, totalScoreCol, nameCol, lastRow) {
    const white = '#ffffff' // needs RGB color
    let colors = ['cyan', 'lime', 'yellow', 'orange', '#ea9999', '#ea9999', '#cfe2f3', '#cfe2f3', '#d9ead3', '#d9ead3', '#fff2cc', '#fff2cc', '#f4cccc', '#f4cccc', 'silver']
      .concat(new Array(12).fill('white'))
    const colorNum = 5
    
    allRange.sort({column: totalScoreCol, ascending: false})
    const nameRange = sheet.getRange(1, nameCol, lastRow)
    this._setHighRankColor(nameRange)

    for (let i=totalScoreCol; i < totalScoreCol + scoresSize * (2 + minkabuTermSize) - 1; i++) {
      let c = false
      for (let j=1; j<scoresSize; j++) {
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