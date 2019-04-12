Namespace Models
    ''' <summary>
    '''     在庫情報
    ''' </summary>
    Public Class StockInfo
        Public Sub New()
            stock_place_cd = String.Empty
            stock_qty = String.Empty
            shipment_qty = string.Empty
            stock_content = string.Empty
        End Sub

        Public Property stock_place_cd As String
        Public Property stock_qty As String
        Public Property shipment_qty As String
        Public Property stock_content As String
    End Class
End NameSpace