Imports KatabanBusinessLogic.Models
Imports S0202.Filters
Imports S0202.MyHelpers
Imports S0202.ViewModels.Prices

Namespace Controllers
    Public Class PricesController
        Inherits Controller

        ' GET: Prices
        <AuthorizeFilter>
        Function Index() As ActionResult

            '選択情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()
            Dim katabanInfo As New KatabanInfo(selectedData, UserHelper.User)
            Dim model As New PricesIndexViewModel(katabanInfo)

            Return View(model)
        End Function
    End Class
End Namespace