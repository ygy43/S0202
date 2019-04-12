Namespace Results
    ''' <summary>
    '''     ELEPattern検証結果
    ''' </summary>
    Public Class ElePatternCheckResult
        Public Sub New(mark As String, result As Boolean)
            Me.ConditionMark = mark
            Me.Result = result
        End Sub

        Public Property ConditionMark As String

        Public Property Result As Boolean
    End Class
End Namespace