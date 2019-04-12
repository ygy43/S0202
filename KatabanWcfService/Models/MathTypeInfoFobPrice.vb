Namespace Models
    Public Class MathTypeInfoFobPrice
        Public Sub New()
            fob_rate = 0
            mathType = string.Empty
            mathPosition = String.Empty
            currency_cd = string.Empty
            authorization_no = String.Empty
        End Sub

        ''' <summary>
        '''     掛率
        ''' </summary>
        ''' <returns></returns>
        Public Property fob_rate As Decimal

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
        
        ''' <summary>
        '''     通貨
        ''' </summary>
        ''' <returns></returns>
        Public Property currency_cd As String
        
        ''' <summary>
        '''     端数処理小数点
        ''' </summary>
        ''' <returns></returns>
        Public Property authorization_no As String
    End Class
End NameSpace