Namespace Results
    Public Class ProcessResult

#Region "コンストラクタ"

        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        Sub New()
            Me.IsSucceed = True
            Me.Errors = New List(Of String)
        End Sub

        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        ''' <param name="errors">エラーメッセージ</param>
        Sub New(errors As IEnumerable(Of String))
            Me.IsSucceed = False
            Me.Errors = errors
        End Sub

#End Region

#Region "プロパティ"

        ''' <summary>
        '''     処理結果
        ''' </summary>
        Public Property IsSucceed As Boolean

        ''' <summary>
        '''     エラーメッセージ
        ''' </summary>
        Public Property Errors As IEnumerable(Of String)

#End Region
    End Class
End NameSpace