@ModelType S0202.ViewModels.Options.RodEndViewModel
@Imports S0202.ViewModels.Options
@Imports KatabanCommon.Constants

<h1>@Model.SeriesName</h1>

@Using Html.BeginForm("Index", "RodEnd", FormMethod.Post, New With {.class = "form-inline text-left", .role = "form"})

    @<div class="container">
        <div class="row">
            @For index = 0 To Model.RodEndUnitInfos.Count - 1
                @<div Class="col-md-6">

                    @*ラジオボタン*@
                    @If Model.RodEndUnitInfos(index).IsEnable
                        @<input type="radio" name="SelectedPatternSymbol" id="SelectedRodEndType" value="@Model.RodEndUnitInfos(index).PatternSymbol" />
                    Else
                        @<input type="radio" name="SelectedPatternSymbol" id="SelectedRodEndType" value="@Model.RodEndUnitInfos(index).PatternSymbol" disabled="disabled" />
                    End If
                    <label>@Model.RodEndUnitInfos(index).PatternSymbol</label>

                    @*各ロッド先端ユニット*@
                    @Code
                        Html.ViewData.TemplateInfo.HtmlFieldPrefix = "RodEndUnitInfos[" & index & "]"
                    End Code

                    @If Model.RodEndUnitInfos(index).PatternType = RodEndUnitDiv.Normal Then

                        @Html.Partial("_RodEndUnitNormal", Model.RodEndUnitInfos(index))

                    ElseIf Model.RodEndUnitInfos(index).PatternType = RodEndUnitDiv.Other Then

                        @Html.Partial("_RodEndUnitOther", Model.RodEndUnitInfos(index))

                    ElseIf Model.RodEndUnitInfos(index).PatternType = RodEndUnitDiv.ImageOnly Then

                        @Html.Partial("_RodEndUnitImageOnly", Model.RodEndUnitInfos(index))

                    End If
                </div>
            Next
        </div>
    </div>


    @<div class="row col-md-12">
        <button type="submit" class="btn btn-primary">OK</button>
        <button type="submit" class="btn btn-primary">Cancel</button>
    </div>
    @Html.HiddenFor(Function(m) m.SeriesName)
    @Html.HiddenFor(Function(m) m.SelectedPatternSymbol)
End Using

<script>

    /**
     *  C外形計算
     */
    function calculateC(event) {
        //KK特注の選択値
        var kkOptionSelect = $(this).closest('table').find(':selected').val();

        //A特注の入力値
        var aValueInput = $(this).closest('table').find('td').filter(
            function () {
                return $(this).text() == 'A';
            }
        ).closest('tr').find('input').val();

        //C特注計算
        if (aValueInput == '') {
            //A特注が入力していない場合は、C特注は、Cの標準値に設定
            var cNormalValue = $(this).closest('table').find('td').filter(
                function () {
                    return $(this).text() == 'C';
                }
            ).closest('tr').find('td:eq(1)').text();

            //C特注設定
            $(this).closest('table').find('td').filter(
                function () {
                    return $(this).text() == 'C';
                }
            ).closest('tr').find('input').val(cNormalValue);
        } else {
            //C特注計算
            var cValue;
            var kkValue = kkOptionSelect.split('|')[1];

            if (kkValue == '') {
                //KKは空白を選択する場合、KK標準差分を使う
                cValue = parseFloat(aValueInput) - parseFloat(kkValueNormal);
            } else {
                cValue = parseFloat(aValueInput) - parseFloat(kkValue);
            }

            //C特注設定
            $(this).closest('table').find('td').filter(
                function () {
                    return $(this).text() == 'C';
                }
            ).closest('tr').find('input').val(cValue);
        }
    }

</script>