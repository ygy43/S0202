Namespace Utility
    Public Class ListMethods
        ''' <summary>
        '''     固定長の文字リストを作成
        ''' </summary>
        ''' <param name="number">アイテム数</param>
        ''' <returns></returns>
        Public Shared Function SetEmptyList(number As Integer) As List(Of String)
            Dim result As New List(Of String)

            For i = 0 To number - 1
                result.Add(String.Empty)
            Next

            Return result
        End Function

        ''' <summary>
        '''     固定長の文字リストを作成
        ''' </summary>
        ''' <param name="number">Item数</param>
        ''' <param name="originalList">元リスト</param>
        ''' <returns></returns>
        Public Shared Function SetEmptyList(number As Integer,
                                            originalList As List(Of String)) As List(Of String)

            While originalList.Count < number
                originalList.Add(String.Empty)
            End While

            Return originalList

        End Function
    End Class
End Namespace