Imports System.Globalization
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.My.Resources
Imports KatabanBusinessLogic.Results

Namespace Managers
    ''' <summary>
    '''     ログイン画面ビジネスロジック
    ''' </summary>
    Public Class UserManager
        ''' <summary>
        '''     ログイン
        ''' </summary>
        ''' <param name="userId">ユーザーID</param>
        ''' <param name="password">パスワード</param>
        ''' <param name="language">言語</param>
        ''' <returns></returns>
        Public Shared Function PasswordSignIn(userId As String,
                                              password As String,
                                              language As String) As LoginResult

            Using client As New DbAccessServiceClient
                Dim userInfo As List(Of UserInfo) = client.SelectUserMstByUserIdAndPassword(userId, password)

                If userInfo.Count = 0 Then
                    '認証失敗
                    Errors.Culture = New CultureInfo(language)
                    Return New LoginResult(New List(Of String) From {Errors.E001})
                Else
                    '認証成功
                    Dim result As New LoginResult

                    result.User = userInfo.First
                    Return result
                End If
            End Using
        End Function

        ''' <summary>
        '''     パスワード更新
        ''' </summary>
        ''' <param name="userId">ユーザーID</param>
        ''' <param name="currentPassword">更新前パスワード</param>
        ''' <param name="newPassword">更新後パスワード</param>
        ''' <returns></returns>
        Public Shared Function ResetPassword(userId As String,
                                             currentPassword As String,
                                             newPassword As String,
                                             language As String) As ProcessResult
            Using client As New DbAccessServiceClient
                Dim userInfo = client.SelectUserMstByUserIdAndPassword(userId, currentPassword)

                If userInfo Is Nothing Then
                    '認証失敗
                    Errors.Culture = New CultureInfo(language)
                    Return New ProcessResult(New List(Of String) From {Errors.E001})
                Else
                    '認証成功の場合は、パスワード更新
                    Dim affectedRows = client.UpdateUserMstPassword(userId, newPassword)

                    If affectedRows = 1 Then
                        Return New ProcessResult
                    Else
                        Return New ProcessResult(New List(Of String) From {Errors.E002})
                    End If
                End If
            End Using
        End Function
    End Class
End Namespace