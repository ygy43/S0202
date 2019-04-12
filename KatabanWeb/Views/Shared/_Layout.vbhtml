<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>@ViewBag.Title - My ASP.NET Application</title>
    @Styles.Render("~/Content/css")
    @Scripts.Render("~/bundles/modernizr")
    @Scripts.Render("~/bundles/jquery")
</head>
<body class="text-center">
    <div class="navbar navbar-inverse navbar-fixed-top navbar-light">
        <div class="container">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                @Html.ActionLink(S0202.My.Resources.RLayout.Title, "Index", "Menu", New With {.area = ""}, New With {.class = "navbar-brand"})
            </div>

            <div class="navbar-collapse collapse">
                <ul class="nav navbar-nav">
                    <li>@Html.ActionLink(S0202.My.Resources.RLayout.ModelSearch, "Search", "Series")</li>
                    <li>@Html.ActionLink(S0202.My.Resources.RLayout.PartNoSearch, "Select", "Options")</li>
                </ul>

                @Html.Partial("_LoginPartial")
                @Html.Partial("_Languages")
            </div>
         </div>
    </div>
    <div class="container body-content">
        @RenderBody()
        <hr />
        <footer>
            <p>Microsoft Internet Explorer 9 or higher ver. are recommended. </p>
            <p>Netscape Navigator is not recommended. </p>
        </footer>
    </div>

    @Scripts.Render("~/bundles/bootstrap")
    @Scripts.Render("~/bundles/custom")
    <script>

        //言語を選択するイベント
        function changeLanguage(returnUrl) {
            var language = $("#Language option:selected").val();
            var redirectURL = '@Url.Action("ChangeCulture", "Account")';
            window.location.href = redirectURL + "?language=" + language + "&returnUrl=" + returnUrl;
        }
    </script>
    @RenderSection("scripts", required:=False)
</body>
</html>
