@ModelType S0202.ViewModels.Account.LoginViewModel
@Code
    ViewBag.Title = S0202.My.Resources.RLogin.Login
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h2>@ViewBag.Title</h2>
<div class="row">
    <div>
        <section id="loginForm">
            @Using Html.BeginForm("Login", "Account", New With {.ReturnUrl = ViewBag.ReturnUrl}, FormMethod.Post, New With {.class = "form-signin", .role = "form"})
                @Html.AntiForgeryToken()
                @<text>
                    <h4>Use a local account to log in.</h4>
                    <hr />
                    @Html.ValidationSummary(True, "", New With {.class = "text-danger"})
                    <div class="form-group">
                        <div class="col-md-12">
                            @Html.TextBoxFor(Function(m) m.UserId, New With {.class = "form-control", .placeholder = S0202.My.Resources.RLogin.UserId, .autofocus = "autofocus"})
                            @Html.ValidationMessageFor(Function(m) m.UserId, "", New With {.class = "text-danger"})
                        </div>
                    </div>
                    <div class="form-group">
                        <div class="col-md-12">
                            @Html.PasswordFor(Function(m) m.Password, New With {.class = "form-control", .placeholder = S0202.My.Resources.RLogin.Password})
                            @Html.ValidationMessageFor(Function(m) m.Password, "", New With {.class = "text-danger"})
                        </div>
                    </div>
                    <div class="form-group">
                        <div class="col-md-12">
                            <input type="submit" value=@S0202.My.Resources.RLogin.Login class="btn btn-lg btn-primary btn-block" />
                        </div>
                    </div>
                </text>
            End Using
        </section>
    </div>
</div>
@Section Scripts
    @Scripts.Render("~/bundles/jqueryval")
End Section
