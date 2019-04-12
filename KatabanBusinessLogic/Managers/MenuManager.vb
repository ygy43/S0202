Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanCommon.Constants

Namespace Managers
    ''' <summary>
    '''     メニュー画面ビジネスロジック
    ''' </summary>
    Public Class MenuManager
        ''' <summary>
        '''     更新履歴を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetUpdateHistories(language As String) As List(Of String)

            Dim result As New List(Of String)

            Using client As New DbAccessServiceClient
                Dim updateHistories As List(Of UpdateHistory) = client.SelectInformationByLanguage(language)

                If updateHistories.Count = 0 Then
                    '対応言語の更新履歴が存在しない場合は、デフォルト言語を取得
                    Dim updateHistoriesDefault As List(Of UpdateHistory) =
                            client.SelectInformationByLanguage(Divisions.LanguageDiv.DefaultLang)

                    result.AddRange(updateHistoriesDefault.Select(Function(u) u.message).ToList())
                Else
                    '対応言語の更新履歴が存在する場合
                    result.AddRange(updateHistories.Select(Function(u) u.message).ToList())
                End If
            End Using

            Return result
        End Function
    End Class
End NameSpace