Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanCommon.Constants

Namespace MyHelpers
    Public Class UserHelper
        ''' <summary>
        '''     ログインユーザー情報
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property User As UserInfo
            Get
                If HttpContext.Current.User.Identity.IsAuthenticated Then
                    Dim userInfo = CType(HttpContext.Current.User, MyPrincipal).User

                    If HttpContext.Current.Session(SessionKeys.Language) IsNot Nothing Then
                        userInfo.language_cd = HttpContext.Current.Session(SessionKeys.Language)
                    End If

                    Return userInfo
                Else
                    Return Nothing
                End If
            End Get
        End Property

        ''' <summary>
        '''     ログオフ
        ''' </summary>
        ''' <param name="session"></param>
        ''' <param name="response"></param>
        Friend Shared Sub Logoff(session As HttpSessionStateBase, response As HttpResponseBase)
            'Delete the user details from cache.
            session.Abandon()

            'Delete the authentication ticket and sign out.
            FormsAuthentication.SignOut()

            'Clear authentication cookie.
            Dim cookie As New HttpCookie(FormsAuthentication.FormsCookieName, "")
            cookie.Expires = DateTime.Now.AddYears(-1)
            response.Cookies.Add(cookie)
        End Sub
    End Class
End Namespace