Namespace Classes
    Public Class ProcessResult

#Region "コンストラクタ"

        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        Sub New()
            Me.IsSucceed = True
            Me.Values = nothing
            Me.Errors = New List(Of String)
        End Sub

        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        ''' <param name="isSucceed">処理結果</param>
        ''' <param name="errors">エラーメッセージ</param>
        Sub New(isSucceed As Boolean, errors As IEnumerable(Of String))
            Me.IsSucceed = isSucceed
            Me.Errors = errors
        End Sub

#End Region

#Region "プロパティ"

        ''' <summary>
        '''     処理結果
        ''' </summary>
        Public Property IsSucceed As Boolean

        ''' <summary>
        '''     成功する場合の情報
        ''' </summary>
        Public Property Values As Object

        ''' <summary>
        '''     エラーメッセージ
        ''' </summary>
        Public Property Errors As IEnumerable(Of String)

#End Region
    End Class
End NameSpace