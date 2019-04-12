@ModelType S0202.ViewModels.Options.RodEndUnitNormalViewModel

<table class="col-md-12">
    <tr>
        <td class="col-md-5">
            <img src=@Url.Content(Model.Image) alt="Image" />
        </td>
        <td class="col-md-7">
            <table class="table table-bordered table-autowidth table-hover">
                <thead>
                    <tr>
                        <th class="col-md-2"></th>
                        <th class="col-md-3">@Model.TitleStandard</th>
                        <th class="col-md-5">@Model.TitleCustom</th>
                    </tr>
                </thead>
                <tbody>
                    @For index = 0 To Model.Rows.Count - 1
                        @<tr>
                            <td>@Model.Rows(index).DisplayExternalForm</td>
                            <td>@Model.Rows(index).StandardValue</td>
                            @If Model.Rows(index).ExternalForm = "KK" Then
                                @*DropDownList*@
                                @<td>
                                    @*C外径を計算*@
                                    @If Model.Rows(index).IsCalculateC Then
                                        @Html.DropDownListFor(Function(model) model.Rows(index).CustomValue, Model.Rows(index).CustomValueOptionSelectList, New With {.class = "form-control", .onchange = "calculateC.call(this,event);"})
                                        @Html.HiddenFor(Function(model) model.Rows(index).ActStandardValue)
                                    Else
                                        @Html.DropDownListFor(Function(model) model.Rows(index).CustomValue, Model.Rows(index).CustomValueOptionSelectList, New With {.class = "form-control"})
                                    End If
                                </td>
                            Else
                                @*TextBox*@
                                @<td>
                                    @If Model.Rows(index).IsEnable Then
                                        If Model.Rows(index).IsCalculateC
                                            @*C外径を計算*@
                                            @Html.TextBoxFor(Function(model) model.Rows(index).CustomValue, New With {.class = "form-control", .type = "number", .onblur = "calculateC.call(this,event);"})
                                        Else

                                            @Html.TextBoxFor(Function(model) model.Rows(index).CustomValue, New With {.class = "form-control", .type = "number"})
                                        End If
                                    Else
                                        @Html.HiddenFor(Function(model) model.Rows(index).CustomValue, New With {.Value = Model.Rows(index).StandardValue})
                                        @Html.TextBox("disabledInput" & index, Model.Rows(index).StandardValue, New With {.class = "form-control", .type = "number", .disabled = "disabled"})
                                    End If
                                </td>
                            End If
                            @Html.HiddenFor(Function(m) m.Rows(index).CustomValueOptions)
                            @Html.HiddenFor(Function(m) m.Rows(index).ExternalForm)
                            @Html.HiddenFor(Function(m) m.Rows(index).DisplayExternalForm)
                            @Html.HiddenFor(Function(m) m.Rows(index).StandardValue)
                        </tr>
                    Next
                </tbody>
            </table>
        </td>
    </tr>
</table>
@Html.HiddenFor(Function(m) m.PatternSymbol)
@Html.HiddenFor(Function(m) m.Image)
@Html.Hidden("ModelType", Model.GetType.Name)