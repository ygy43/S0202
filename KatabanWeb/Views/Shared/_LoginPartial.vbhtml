@Imports Microsoft.AspNet.Identity
@Imports S0202.MyHelpers

@If Request.IsAuthenticated
    @Html.AntiForgeryToken()
    @<ul class="nav navbar-nav navbar-right">
        <li>
            @Html.ActionLink("Hello " + UserHelper.User.user_id + "!", "Index", "Manage", routeValues:=Nothing, htmlAttributes:=New With {.title = "Manage"})
        </li>
        <li>
            @Html.ActionLink(S0202.My.Resources.RLayout.LogOff, "LogOff", "Account", routeValues:=Nothing, htmlAttributes:=New With {.title = "LogOff"})
        </li>
    </ul>
Else
    @*@<ul class="nav navbar-nav navbar-right">
            <li>@Html.ActionLink(S0202.My.Resources.RLayout.Register, "Register", "Account", routeValues:=Nothing, htmlAttributes:=New With {.id = "registerLink"})</li>
            <li>@Html.ActionLink(Resources.RLayout.LogIn, "Login", "Account", routeValues:=Nothing, htmlAttributes:=New With {.id = "loginLink"})</li>
        </ul>*@
End If