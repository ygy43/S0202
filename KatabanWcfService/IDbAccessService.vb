Imports S0202.Models

<ServiceContract>
Public Interface IDbAccessService

#Region "ログイン画面関連"

    <OperationContract>
    Function SelectUserMstByUserIdAndPassword(userId As String, password As String) As IEnumerable(Of UserInfo)

    <OperationContract>
    Function UpdateUserMstPassword(userId As String, newPassword As String) As Integer

#End Region

#Region "メニュー画面関連"

    <OperationContract>
    Function SelectInformationByLanguage(language As String) As IEnumerable(Of UpdateHistory)

#End Region

#Region "機種選択画面"

    <OperationContract>
    Function SelectSeriesInfoBySeries(input As String,
                                      country As String,
                                      language As String,
                                      page As Integer,
                                      pageSize As Integer) As IEnumerable(Of SeriesInfo)

    <OperationContract>
    Function SelectSeriesInfoCountBySeries(input As String,
                                           country As String,
                                           language As String) As Integer

    <OperationContract>
    Function SelectSeriesInfoWithKeyBySeries(series As String,
                                             keyKataban As String,
                                             country As String,
                                             language As String) As SeriesInfo

    <OperationContract>
    Function SelectSeriesInfoByFullKataban(input As String,
                                           country As String,
                                           language As String,
                                           page As Integer,
                                           pageSize As Integer) As IEnumerable(Of SeriesInfoFullKataban)

    <OperationContract>
    Function SelectSeriesInfoCountByFullKataban(input As String,
                                                country As String,
                                                language As String) As Integer 

    <OperationContract>
    Function SelectSeriesInfoWithKeyByFullKataban(series As String,
                                                  currency As String,
                                                  country As String,
                                                  language As String) As SeriesInfoFullKataban

    <OperationContract>
    Function SelectSeriesInfoByShiire() As IEnumerable(Of SeriesInfo)

    <OperationContract>
    Function SelectSeriesInfoWithKeyByShiire() As SeriesInfo

#End Region

#Region "オプション選択画面"

    <OperationContract>
    Function SelectKatabanStructure(series As String, keyKataban As String, language As String) _
        As IEnumerable(Of KatabanStructureInfo)

    <OperationContract>
    Function SelectKatabanStructureOptions(series As String, keyKataban As string) As List(Of KatabanStructureOptionInfoAllSeqNo)

    <OperationContract>
    Function SelectKatabanStructureOptionsBySeqNo(series As String, keyKataban As String, seqNo As Integer, language As String) _
        As IEnumerable(Of KatabanStructureOptionInfo)

    <OperationContract>
    Function SelectElePatternInfoAll(series As String, keyKataban As String, seqNo As Integer) _
        As IEnumerable(Of ElePatternInfo)

    <OperationContract>
    Function SelectElePatternInfoPlural(series As String, keyKataban As String, seqNo As Integer) _
        As IEnumerable(Of ElePatternInfo)
    
    <OperationContract>
    Function SelectRodEndInfo(series As String, keyKataban As String) As IEnumerable(Of RodEndInfo)

    <OperationContract>
    Function SelectRodEndExternalFormInfo(series As String, keyKataban As String, boreSize As Integer) As IEnumerable(of RodEndExternalFormInfo)

    <OperationContract>
    Function SelectWfMaxValue(series As String, keyKataban As String, boreSize As Integer) As String
#End Region

#Region "製品情報"

    <OperationContract>
    Function SelectFullKatabanPriceInfo(kataban As String, currency As String) As PriceInfo

    <OperationContract>
    Function SelectAccumulatePriceInfo(kataban As String, currency As String) As PriceInfo

    <OperationContract>
    Function SelectScrewPriceInfo(kataban As String) As ScrewPriceInfo

    <OperationContract>
    Function AddPriceInfo(price1 As PriceInfo, price2 As PriceInfo) As PriceInfo

    <OperationContract>
    Function MinusPriceInfo(price1 As PriceInfo, price2 As PriceInfo) As PriceInfo

    <OperationContract>
    Function MultiplePriceInfo(price As PriceInfo, multiple As List(Of Decimal)) As PriceInfo

    <OperationContract>
    Function SelectExchangeRate(katabanCurrency As String, userCurrency As String) As Decimal

    <OperationContract>
    Function SelectMathTypeLocalPrice(userCountry As String, katabanFirstHyphen As String) As MathTypeInfoLocalPrice

    <OperationContract>
    Function SelectMathTypeFobPrice(userCountry As String, katabanFirstHyphen As String, shipPlace As String) _
        As MathTypeInfoFobPrice

    <OperationContract>
    Function SelectQuantityUnit(kataban As String, language As String) As QuantityUnitInfo

    <OperationContract>
    Function CheckEl(kataban As String, elFlag As String) As Boolean

    <OperationContract>
    Function SelectStock(kataban As String, language As String, shipPlace As String) As StockInfo

    <OperationContract>
    Function SelectStroke(series As String, keyKataban As String, boreSize As Integer, country As String) _
        As List(Of StrokeInfo)

#End Region

#Region "価格キー関連"

    <OperationContract>
    Function SelectVoltageInfo(series As String,
                               keyKataban As String,
                               portSize As String,
                               coil As String,
                               voltageDiv As String,
                               voltage As Integer) As List(Of VoltageInfo)

#End Region
End Interface