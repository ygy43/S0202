Imports System.Globalization
Imports System.IO
Imports System.Security.Principal
Imports System.Web.Optimization
Imports System.Xml.Serialization
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanCommon.Constants
Imports S0202.MyHelpers
Imports S0202.ViewModels.Options

Public Class MvcApplication
    Inherits HttpApplication

    ''' <summary>
    '''     アプリケーション起動
    ''' </summary>
    Sub Application_Start()
        AreaRegistration.RegisterAllAreas()
        RegisterGlobalFilters(GlobalFilters.Filters)
        RegisterRoutes(RouteTable.Routes)
        RegisterBundles(BundleTable.Bundles)

        'ロッド先端情報を種類ごとにバインドできるように
        ModelBinders.Binders.Add(GetType(RodEndUnitViewModel), New RodEndUnitViewModelModelBinder)
    End Sub

    ''' <summary>
    '''     リクエスト発生
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub Application_AuthenticateRequest(sender As Object, e As EventArgs)
        If HttpContext.Current.User IsNot Nothing Then
            If HttpContext.Current.User.Identity.IsAuthenticated Then
                Dim id = TryCast(HttpContext.Current.User.Identity, FormsIdentity)
                If (id IsNot Nothing) Then
                    Dim ticket As FormsAuthenticationTicket = id.Ticket
                    Dim identity As New GenericIdentity(ticket.Name, "Forms")
                    Dim principal = New MyPrincipal(identity)
                    Dim userData As String = ticket.UserData
                    Dim deserializer = New XmlSerializer(GetType(UserInfo))

                    Using tr As TextReader = New StringReader(userData)
                        'ディシリアル化
                        principal.User = CType(deserializer.Deserialize(tr), UserInfo)
                        'ユーザー情報を保存
                        HttpContext.Current.User = principal
                    End Using
                End If
            End If
        End If
    End Sub

    Sub Application_AcquireRequestState(sender As Object, e As EventArgs)

        Dim languageSession = Session(SessionKeys.Language)
        Dim language = LanguageDiv.DefaultLang

        If languageSession IsNot Nothing Then
            If Not String.IsNullOrEmpty(languageSession) Then
                language = languageSession
            End If
        End If

        Dim cultureInfo = New CultureInfo(language)

        Threading.Thread.CurrentThread.CurrentCulture = cultureInfo
        Threading.Thread.CurrentThread.CurrentUICulture = cultureInfo
    End Sub
End Class
