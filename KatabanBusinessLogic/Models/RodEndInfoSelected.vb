Namespace Models
    ''' <summary>
    '''     ロッド先端情報
    ''' </summary>
    Public Class RodEndInfoSelected
        Public Sub New()
            RodEndOption = string.Empty 
            RodEndWFStdVal = string.Empty
        End Sub

        ''' <summary>
        '''     ロッド先端特注WF標準寸法
        ''' </summary>
        Public Property RodEndOption As String

        ''' <summary>
        '''     ロッド先端特注WF標準寸法
        ''' </summary>
        Public Property RodEndWFStdVal As String
    End Class
End NameSpace