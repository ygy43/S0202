Imports KatabanBusinessLogic
Imports KatabanBusinessLogic.Managers
Imports KatabanCommon.Constants
Imports S0202.Filters
Imports S0202.MyHelpers
Imports S0202.ViewModels.Menu

Namespace Controllers
    <AuthorizeFilter>
    Public Class MenuController
        Inherits Controller

        ' GET: /Menu/Index
        Function Index() As ActionResult

            Dim viewmodel As New MenuIndexViewModel
            If UserHelper.User.user_class >= Levels.UserClassLevel.DmSalesOffice Then

                viewmodel.Messages = MenuManager.GetUpdateHistories("ja")
            Else
                viewmodel.Messages = New List(Of String)
            End If
            Return View(viewmodel)
        End Function
    End Class
End Namespace