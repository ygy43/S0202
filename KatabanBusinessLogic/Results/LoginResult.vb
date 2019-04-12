Imports KatabanBusinessLogic.KatabanWcfService

Namespace Results
    ''' <summary>
    '''     ログイン結果
    ''' </summary>
    Public Class LoginResult
        Inherits ProcessResult

        Public Sub New()
            MyBase.New()
            Me.User = New UserInfo()
        End Sub

        Sub New(errors As IEnumerable(Of String))
            MyBase.New(errors)
        End Sub

#Region "プロパティ"

        ''' <summary>
        '''     処理結果
        ''' </summary>
        Public Property User As UserInfo

#End Region
    End Class
End Namespace