@ModelType S0202.ViewModels.Options.RodEndUnitOtherViewModel

<table>
    <tr>
        <td>
            @Model.TextTitle
        </td>
    </tr>

    <tr>
        <td>
            @Html.TextBoxFor(Function(model) model.CustomValue)
        </td>
    </tr>
</table>
@Html.HiddenFor(Function(m) m.PatternSymbol)
@Html.Hidden("ModelType", Model.GetType.Name)