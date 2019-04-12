Namespace Results
    Public Class RodEndCheckResult
        Inherits ProcessResult

        Public Sub New()
            MyBase.New()
            Me.ErrorSeqNo = 0
        End Sub

        Sub New(errorSeqNo As Integer, errors As IEnumerable(Of String))
            MyBase.New(errors)
            Me.ErrorSeqNo = errorSeqNo
        End Sub

#Region "プロパティ"

        ''' <summary>
        '''     エラー構成番号
        ''' </summary>
        Public Property ErrorSeqNo As Integer

#End Region
    End Class
End NameSpace