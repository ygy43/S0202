@ModelType S0202.ViewModels.Options.RodEndUnitOnlyImageViewModel

<table>
    <tr>
        <td>
            @Model.Message
        </td>
    </tr>

    <tr>
        <td>
            <img src=@Url.Content(Model.Image) alt="Image" />
        </td>
    </tr>
</table>
@Html.HiddenFor(Function(m) m.PatternSymbol)
@Html.HiddenFor(Function(m) m.Image)
@Html.Hidden("ModelType", Model.GetType.Name)