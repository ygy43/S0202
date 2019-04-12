Namespace Models
    ''' <summary>
    '''     価格キー情報
    ''' </summary>
    Public Class PriceKeyInfo
        Public Sub New()
            PriceKey = String.Empty
            Amount = 0
            PriceKeyDiv = string.Empty
        End Sub

        ''' <summary>
        '''     価格キー
        ''' </summary>
        Public Property PriceKey As String

        ''' <summary>
        '''     価格キー数量
        ''' </summary>
        Public Property Amount As Decimal

        ''' <summary>
        '''     価格キー区分
        ''' </summary>
        Public Property PriceKeyDiv As String
    End Class
End NameSpace