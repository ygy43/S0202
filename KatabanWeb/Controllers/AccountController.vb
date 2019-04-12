Imports System.Globalization
Imports System.IO
Imports System.Xml.Serialization
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Results
Imports KatabanCommon.Constants
Imports Newtonsoft.Json.Linq
Imports S0202.MyHelpers
Imports S0202.ViewModels.Account

Namespace Controllers
    Public Class AccountController
        Inherits Controller

#Region "イベント"

        ' GET: /Account/Login
        Public Function Login() As ActionResult
            Return View()
        End Function

        ' POST: /Account/Login
        <HttpPost>
        Public Function Login(model As LoginViewModel) As ActionResult
            If Not ModelState.IsValid Then
                Return View(model)
            End If

            'ユーザー認証
            Dim result = UserManager.PasswordSignIn(model.UserId, model.Password, "ja-JP")

            If result.IsSucceed Then
                '認証成功情報を保存
                Dim userInfo As UserInfo = result.User
                Dim serializer As New XmlSerializer(GetType(UserInfo))

                Using sw As New StringWriter
                    serializer.Serialize(sw, userInfo)

                    Dim userData = sw.ToString()
                    Dim ticket = New FormsAuthenticationTicket(1,
                                                           model.UserId,
                                                           Now,
                                                           Now.AddMinutes(2880),
                                                           False,
                                                           userData,
                                                           FormsAuthentication.FormsCookiePath)

                    Dim hash As String = FormsAuthentication.Encrypt(ticket)
                    Dim cookie As New HttpCookie(FormsAuthentication.FormsCookieName, hash)
                    If ticket.IsPersistent Then
                        cookie.Expires = ticket.Expiration
                    End If
                    Response.Cookies.Add(cookie)
                End Using

                'メニュー画面へ遷移
                Return RedirectToAction("Index", "Menu")

            End If

            'エラー情報を設定
            AddErrors(result)
            Return View()
        End Function

        ' GET: /Account/ResetPassword
        Public Function ResetPassword() As ActionResult
            Return View()
        End Function

        ' POST: /Account/ResetPassword
        <HttpPost>
        Public Function ResetPassword(model As ResetPasswordViewModel, language As String) As ActionResult
            If Not ModelState.IsValid Then
                Return View(model)
            End If

            If Session("UserInfo") IsNot Nothing Then
                Dim result = UserManager.ResetPassword(Session("UserInfo")("user_id"),
                                                       model.CurrentPassword,
                                                       model.NewPassword,
                                                       language)

                If result.IsSucceed Then
                    Return RedirectToAction("ResetPasswordConfirmation", "Account")
                End If

                AddErrors(result)
            End If
            Return View()
        End Function

        ' GET: /Account/ResetPasswordConfirmation
        Public Function ResetPasswordConfirmation() As ActionResult
            Return View()
        End Function

        ' GET: /Account/ChangeCulture
        Public Function ChangeCulture(language As String, returnUrl As String) As ActionResult

            'Cookieに追加
            Session.Remove(SessionKeys.Language)
            Session.Add(SessionKeys.Language, language)

            Return Redirect(Request.UrlReferrer.ToString())

        End Function

        ' GET: /Account/LogOff
        Public Function LogOff() As ActionResult

            UserHelper.Logoff(Session, Response)

            Return RedirectToAction("Login", "Account")
        End Function

#End Region

#Region "メソッド"

        ''' <summary>
        '''     エラー情報を追加
        ''' </summary>
        ''' <param name="result"></param>
        Private Sub AddErrors(result As ProcessResult)
            For Each [error] As String In result.Errors
                ModelState.AddModelError("", [error])
            Next
        End Sub

#End Region
    End Class
End Namespace