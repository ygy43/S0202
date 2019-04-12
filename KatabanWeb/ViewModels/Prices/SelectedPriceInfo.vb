Namespace ViewModels.Prices
    ''' <summary>
    '''     単価画面選択した価格情報
    ''' </summary>
    Public Class SelectedPriceInfo
        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        Public Sub New()
            Rate = "1.0000"
            TotalWithoutTax = String.Empty
            Price = String.Empty
            Tax = String.Empty
            Amount = 1
            TotalWithTax = String.Empty
        End Sub

        ''' <summary>
        '''     掛率
        ''' </summary>
        Public Property Rate As String
        
        ''' <summary>
        '''     金額
        ''' </summary>
        Public Property TotalWithoutTax As String

        ''' <summary>
        '''     単価
        ''' </summary>
        Public Property Price As String

        ''' <summary>
        '''     消費税
        ''' </summary>
        Public Property Tax As String

        ''' <summary>
        '''     数量
        ''' </summary>
        Public Property Amount As Integer
        
        ''' <summary>
        '''     合計
        ''' </summary>
        Public Property TotalWithTax as String

    End Class
End NameSpace