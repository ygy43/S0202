Namespace Filters
    Public Class AuthorizeFilter
        Inherits ActionFilterAttribute
        Implements IActionFilter

        Public Overrides Sub OnActionExecuting(filterContext As ActionExecutingContext)
            If Not HttpContext.Current.User.Identity.IsAuthenticated Then
                '認証されない場合はログイン画面に遷移
                filterContext.Result = New RedirectToRouteResult(New RouteValueDictionary _
                                                                    From {{"Controller", "Account"}, {"Action", "Login"}})
            End If
            MyBase.OnActionExecuting(filterContext)
        End Sub
    End Class
End Namespace