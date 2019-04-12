
Imports KatabanBusinessLogic.Models

Namespace ViewModels.Prices
    Public Class PricesIndexViewModel
        Public Sub New(katabanInfo As KatabanInfo)

            '表示名称
            Me.DisplayName = katabanInfo.DisplayName

            'フル形番
            Me.FullKataban = katabanInfo.FullKataban

            '価格情報
            Me.ListPrice = katabanInfo.ListPrice
            Me.RegisterPrice = katabanInfo.RegisterPrice
            Me.SsPrice = katabanInfo.SsPrice
            Me.BsPrice = katabanInfo.BsPrice
            Me.GsPrice = katabanInfo.GsPrice
            Me.PsPrice = katabanInfo.PsPrice

            '通貨
            Me.Currency = katabanInfo.Currency

            '現地定価
            Me.LocalPrice = katabanInfo.LocalPrice

            '購入価格
            Me.FobPrice = katabanInfo.FobPrice

            'チェック区分
            Me.CheckDiv = katabanInfo.CheckDiv

            '標準納期
            Me.StandardNouki = katabanInfo.StandardNouki

            '適用個数
            Me.Kosuu = katabanInfo.Kosuu

            '販売数量単位
            Me.QuantityPerSalesUnit = katabanInfo.QuantityPerSalesUnit

            'EL情報
            Me.ElDiv = katabanInfo.ElFlag

            '在庫情報
            Me.Stock = katabanInfo.Stock

            '出荷場所
            Me.SelectedShipPlace = String.Empty
            Me.ShipPlaces = katabanInfo.ShipPlaces

            '検索した価格情報
            Me.SelectedPrice = New SelectedPriceInfo
        End Sub

        ''' <summary>
        '''     表示名称
        ''' </summary>
        Public Property DisplayName As String

        ''' <summary>
        '''     フル形番
        ''' </summary>
        Public Property FullKataban As String

        ''' <summary>
        '''     チェック区分
        ''' </summary>
        Public Property CheckDiv As String

        ''' <summary>
        '''     標準納期
        ''' </summary>
        Public Property StandardNouki As String

        ''' <summary>
        '''     適用個数
        ''' </summary>
        Public Property Kosuu As String

        ''' <summary>
        '''     販売数量
        ''' </summary>
        Public Property QuantityPerSalesUnit As String

        ''' <summary>
        '''     販売数量単位
        ''' </summary>
        ''' <returns></returns>
        Public Property SalesUnit As String

        ''' <summary>
        '''     EL区分
        ''' </summary>
        Public Property ElDiv As String

        ''' <summary>
        '''     在庫情報
        ''' </summary>
        Public Property Stock As String

        ''' <summary>
        '''     選択した出荷場所
        ''' </summary>
        Public Property SelectedShipPlace As String

        ''' <summary>
        '''     出荷場所候補
        ''' </summary>
        Public Property ShipPlaces As List(Of String)

        ''' <summary>
        '''     定価
        ''' </summary>
        Public Property ListPrice As String

        ''' <summary>
        '''     登録店価格
        ''' </summary>
        Public Property RegisterPrice As String

        ''' <summary>
        '''     SS店価格
        ''' </summary>
        Public Property SsPrice As Decimal

        ''' <summary>
        '''     BS店価格
        ''' </summary>
        Public Property BsPrice As Decimal

        ''' <summary>
        '''     GS店価格
        ''' </summary>
        Public Property GsPrice As Decimal

        ''' <summary>
        '''     PS店価格
        ''' </summary>
        Public Property PsPrice As Decimal

        ''' <summary>
        '''     通貨
        ''' </summary>
        Public Property Currency As String

        ''' <summary>
        '''     現地定価
        ''' </summary>
        Public Property LocalPrice As String

        ''' <summary>
        '''     購入価格
        ''' </summary>
        Public Property FobPrice As String

        ''' <summary>
        '''     選択した価格情報
        ''' </summary>
        Public Property SelectedPrice As SelectedPriceInfo
    End Class
End Namespace