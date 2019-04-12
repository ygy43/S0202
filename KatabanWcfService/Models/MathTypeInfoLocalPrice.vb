Namespace Models
    Public Class MathTypeInfoLocalPrice
        Public Sub New()
            list_price_rate1 = 0
            list_price_rate2 = 0
            mathType = string.Empty
            mathPosition = String.Empty
        End Sub

        ''' <summary>
        '''     掛率１
        ''' </summary>
        ''' <returns></returns>
        Public Property list_price_rate1 As Decimal

        ''' <summary>
        '''     掛率２
        ''' </summary>
        ''' <returns></returns>
        Public Property list_price_rate2 As Decimal

        ''' <summary>
        '''     端数処理方法
        ''' </summary>
        ''' <returns></returns>
        Public Property mathType As String

        ''' <summary>
        '''     端数処理小数点
        ''' </summary>
        ''' <returns></returns>
        Public Property mathPosition As String
    End Class
End NameSpace