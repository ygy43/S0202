@ModelType S0202.ViewModels.Menu.MenuIndexViewModel
@Code
    ViewData("Title") = "更新履歴"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h2>@ViewData("Title")</h2>

<div class="form-group">
    <select multiple class="form-control height-information">
        @For Each message In Model.Messages
            @<option>
                @message
            </option>
        Next
    </select>
</div>
