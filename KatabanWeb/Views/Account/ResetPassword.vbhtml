@ModelType S0202.ViewModels.Account.ResetPasswordViewModel
@Code
    ViewBag.Title = "ResetPassword"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h2>@ViewBag.Title</h2>
<div class="row">
    <div class="col-md-8">
        <section id="loginForm">
            @Using Html.BeginForm("ResetPassword", "Account", New With {.ReturnUrl = ViewBag.ReturnUrl}, FormMethod.Post, New With {.class = "form-horizontal", .role = "form"})
                @Html.AntiForgeryToken()
                @<text>
                    <h4>Use a local account to log in.</h4>
                    <hr />
                    @Html.ValidationSummary(True, "", New With {.class = "text-danger"})
                    <div class="form-group">
                        @Html.LabelFor(Function(m) m.CurrentPassword, New With {.class = "col-md-2 control-label"})
                        <div class="col-md-10">
                            @Html.TextBoxFor(Function(m) m.CurrentPassword, New With {.class = "form-control"})
                            @Html.ValidationMessageFor(Function(m) m.CurrentPassword, "", New With {.class = "text-danger"})
                        </div>
                    </div>
                    <div class="form-group">
                        @Html.LabelFor(Function(m) m.NewPassword, New With {.class = "col-md-2 control-label"})
                        <div class="col-md-10">
                            @Html.PasswordFor(Function(m) m.NewPassword, New With {.class = "form-control"})
                            @Html.ValidationMessageFor(Function(m) m.NewPassword, "", New With {.class = "text-danger"})
                        </div>
                    </div>
                    <div class="form-group">
                        <div class="col-md-offset-2 col-md-10">
                            <input type="submit" value="Log in" class="btn btn-default" />
                        </div>
                    </div>
                    @*<p>
                            @Html.ActionLink("Register as a new user", "Register")
                        </p>*@
                    @* Enable this once you have account confirmation enabled for password reset functionality
                        <p>
                            @Html.ActionLink("Forgot your password?", "ForgotPassword")
                        </p>*@
                </text>
            End Using
        </section>
    </div>
</div>
@Section Scripts
    @Scripts.Render("~/bundles/jqueryval")
End Section
