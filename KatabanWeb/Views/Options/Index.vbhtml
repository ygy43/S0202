@ModelType S0202.ViewModels.Options.OptionsIndexViewModel
@Imports KatabanCommon.Constants
@Code
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h2>@Model.SelectedSeriesInfo.disp_name</h2>
@Html.ValidationSummary(True, "", New With {.class = "text-danger"})

@Using Html.BeginForm("Index", "Options", FormMethod.Post, New With {.class = "form-inline text-left", .role = "form"})

    @Html.AntiForgeryToken()
    @<div class="row col-md-12">

        @Html.Label(Model.SelectedSeriesInfo.series_kataban)
        @If Model.SelectedSeriesInfo.hyphen_div = HyphenDiv.Necessary Then
            @Html.Label(MyControlChars.Hyphen, New With {.style = "margin:5px;"})
        End If

        @For index = 0 To Model.KatabanStructureInfos.Count - 1

            @*構成（入力用）*@
            If index = 0
                @*0番のデータが使用されていない*@
                @Html.TextBoxFor(Function(model) model.SelectedStructureInfos(index), New With {.style = "display: none;"})
            Else

                @Html.TextBoxFor(Function(model) model.SelectedStructureInfos(index),
                                 New With {.id = "structure" & index,
                                 .class = "form-control structureText",
                                 .style = "width:" & Model.KatabanStructureInfos(index).Width & "px;margin:3px;",
                                 .onchange = "validateInput('" & index & "');"})
            End If

            @*ハイフン*@
            @If Model.KatabanStructureInfos(index).hyphen_div = HyphenDiv.Necessary Then
                @Html.Label(MyControlChars.Hyphen, New With {.style = "margin:5px;"})
            End If

            @*構成区分*@
            @Html.Hidden("structureDiv" & index, Model.KatabanStructureInfos(index).structure_div)

            @*構成名称*@
            @If String.IsNullOrEmpty(Model.KatabanStructureInfos(index).ktbn_strc_nm) Then
                @Html.Hidden("structureName" & index, Model.KatabanStructureInfos(index).default_nm)
            Else
                @Html.Hidden("structureName" & index, Model.KatabanStructureInfos(index).ktbn_strc_nm)
            End If

            @*複数選択可能な場合はグループ情報を出力*@
            @Html.Hidden("pluralGroupData" & index, Model.KatabanStructureInfos(index).PluralGroupData)
        Next
    </div>

    @*警告メッセージ*@
    @<div class="row col-md-12">
        <div class="col-md-8">
            @For Each message In Model.Messages
                @<h5>@message</h5>
            Next
        </div>
        <div class="col-md-4" style="text-align: right;">
            <div class="row">
                @If Model.IsShowRodEnd Then
                    @<input type="submit" formaction=@Url.Action("ShowRodEnd", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.RodEnd Class="btn btn-primary" />
                End If
                @If Model.IsShowOtherOption Then
                    @<button type="submit" formaction=@Url.Action("ShowOtherOption", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.OtherOption Class="btn btn-primary" />
                End If
                @If Model.IsShowStopper Then
                    @<button type="submit" formaction=@Url.Action("ShowStopper", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.Stopper Class="btn btn-primary" />
                End If
                @If Model.IsShowMotor1 Then
                    @<button type="submit" formaction=@Url.Action("ShowMotor1", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.Motor1 Class="btn btn-primary" />
                End If
                @If Model.IsShowMotor2 Then
                    @<button type="submit" formaction=@Url.Action("ShowMotor2", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.Motor2 Class="btn btn-primary" />
                End If
                @If Model.IsShowPortPosition Then
                    @<button type="submit" formaction=@Url.Action("ShowPortPosition", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.PortPosition Class="btn btn-primary" />
                End If
                @If Model.IsShowStock Then
                    @<button type="submit" formaction=@Url.Action("ShowStock", "Options") formmethod="post" value=@S0202.My.Resources.ROptions.Stock Class="btn btn-primary" />
                End If
            </div>
        </div>
    </div>

    @*OKボタン*@
    @<div>
        <button type="submit" class="btn btn-primary" id="ok">OK</button>
    </div>

    @*構成オプション*@
    @<div id="currentOptions"></div>

    @*選択した機種情報*@
    @*機種*@
    @Html.HiddenFor(Function(model) model.SelectedSeriesInfo.series_kataban)

    @*キー形番*@
    @Html.HiddenFor(Function(model) model.SelectedSeriesInfo.key_kataban)

    @*フォカスする構成番号*@
    @Html.HiddenFor(Function(model) model.FocusSeqNo, New With {.id = "focusIndex"})
    @*@Html.Hidden("focusIndex")*@

    @*構成数量*@
    @Html.Hidden("structureNumber", Model.KatabanStructureInfos.Count - 1)
End Using

<script>

    /**
     *  画面ロード
     *  フォカスされる時、オプションリストを更新
     *  次に遷移する時、入力検証
     */
    $(document).ready(function () {
        $(".structureText").on("focus", function() {
            //e.preventDefault;
            var index = parseInt(this.id.replace("structure", ""));
            var structureDiv = $('#structureDiv' + index).val();

            if (structureDiv < '4') {
                $(this).select();
            }
            else {
                $(this).val("");
            }

            //フォカスされる時、オプションリストを更新
            updateResults(index);
        });

        $("#structure" + @Model.FocusSeqNo).focus();
        $("#focusIndex").val(@Model.FocusSeqNo);
    })

    function updateresults() {
                    //e.preventDefault;
            var index = parseInt(this.id.replace("structure", ""));
            var structureDiv = $('#structureDiv' + index).val();

            if (structureDiv < '4') {
                $(this).select();
            }
            else {
                $(this).val("");
            }

            //フォカスされる時、オプションリストを更新
            updateResults(index);
        }

    /**
     * 構成候補の更新
     * index 構成番号
     */
    function updateResults(index) {

        //選択した構成情報
        var selectedStructures = getSelectedStructures(index);
        //構成名称
        var name = $('#structureName' + index).val();
        //構成数量
        var strNumber = parseInt($('#structureNumber').val());
        //構成区分
        var selectedStructureDiv = $('#structureDiv' + index).val();

        //更新データ
        var dataPost = {
            structureName: name,
            focusSeqNo: index,
            structureNumber: strNumber,
            structures: selectedStructures,
            structureDiv: selectedStructureDiv
        };

        //更新
        var nextUrl = '@Url.Action("UpdateOptions", "Options")';

        $.ajax({
            url: nextUrl,
            data: dataPost,
            //cache: false,
            error: function (xhr, status, error) {
                var err = eval("(" + xhr.responseText + ")");
                alert(err.Message);
            }
        }).done(function (partialViewResult) {
            if (partialViewResult.length < 20) {

                if (partialViewResult != '') {
                    //オプションが一つしかいない場合
                    $('#structure' + index).val(partialViewResult);
                }
                //次の構成にフォカス
                var nextIndex = index + 1;

                if (nextIndex > strNumber) {

                    $('#ok').focus();
                    $('#searchResultsTable').empty();
                } else {
                    $('#structure' + nextIndex).focus();
                }

            } else {
                $('#currentOptions').html(partialViewResult);
                //フォカスする構成番号を記録
                $("#focusIndex").val(index);
            }
        });
    }

    /**
     * 入力検証
     */
    function validateInput(index) {
        if (index > 0) {
            var selectedStructures = getSelectedStructures(index);
            //検証データ
            var dataPost = {
                structures: selectedStructures,
                seqNo: index
            };

            //検証
            var validateUrl = '@Url.Action("ValidateInputBySeqNo", "Options")';

            $.ajax({
                url: validateUrl,
                data: dataPost,
                error: function (xhr, status, error) {
                    var err = eval("(" + xhr.responseText + ")");
                    alert(err.Message);
                }
            }).done(function (validateResult) {
                if (validateResult.length != 0) {
                    var results = validateResult.split("|");
                    alert(results[1]);
                    $("#structure" + results[0]).focus();
                    return;
                }
            });
        }
    }

    /**
     * 選択した情報の取得
     */
    function getSelectedStructures(index) {
        //選択した構成情報
        var selectedStructures = '';
        for (var i = 0; i <= index; i++) {
            if (i == index) {
                if ($('#structure' + i).val() != null) {
                    selectedStructures += $('#structure' + i).val();
                }
            } else {
                if ($('#structure' + i).val() == null) {
                    selectedStructures += '|';
                } else {
                    selectedStructures += $('#structure' + i).val() + '|';
                }
            };
        };

        return selectedStructures;
    }

    /**
     * オプションダブルクリック
     * */
    function focusNext(selectedSymbol, rowIndex) {

        //構成数量
        var structureNumber = parseInt($('#structureNumber').val());

        //フォカスする構成番号
        var focusIndex = $('#focusIndex').val();

        //次のインデックス
        var nextIndex = parseInt(focusIndex) + 1;

        //フォカスする構成の区分
        var structureDiv = $('#structureDiv' + focusIndex).val();

        if (structureDiv < '4') {
            //複数選択可能な構成以外の場合

            //選択値(本物)を設定
            $('#structure' + focusIndex).val(selectedSymbol);

            //フォカスの設定
            if (nextIndex > structureNumber) {
                //最後の場合はOKボタンにフォカス
                $('#ok').focus();
                $('#searchResultsTable').empty();
            } else {
                //次の構成にフォカス
                $('#structure' + nextIndex).focus();
            }
        } else {
            //複数選択可能な構成なら

            if (selectedSymbol == '') {
                //選択終了を選択した場合は次の構成に遷移
                $('#searchResultsTable').empty();
            } else {

                //選択した行以前の行を削除
                for (var i = 1; i < rowIndex; i++) {
                    //オプション
                    var symbolValue = $('#row' + i).find('td:first-child').map(function () {
                        return $(this).text();
                    }).get();

                    if (symbolValue != '') {
                        //選択終了以外を削除
                        $('#row' + i).remove();
                    }
                };

                //選択した構成と同じグループのオプションを削除
                var pluralGroupData = $('#pluralGroupData' + focusIndex).val();
                var groupOptions = pluralGroupData.split('|');
                var deleteSymbols = new Array(selectedSymbol);

                for (var i = 0; i < groupOptions.length; i++) {
                    var symbolInSameGroup = groupOptions[i].split(',');

                    if (symbolInSameGroup.indexOf(selectedSymbol) >= 0) {
                        deleteSymbols = deleteSymbols.concat(symbolInSameGroup);
                    }
                }

                for (var i = 0; i < deleteSymbols.length; i++) {
                    $("#searchResultsTable tr td:contains('" + deleteSymbols[i] + "')").filter(function () {
                        return $(this).text().trim() == deleteSymbols[i];
                    }).parent().remove();
                }
            }

            //選択値(表示)を設定
            var structure = $('#structure' + focusIndex).val();
            $('#structure' + focusIndex).val(structure + selectedSymbol);

            if ($('#searchResultsTable tr').length == 1) {
                //選択可能なオプションがなければ、次の構成に移動
                //フォカスの設定
                if (nextIndex > structureNumber) {
                    //最後の場合はOKボタンにフォカス
                    $('#ok').focus();
                    $('#searchResultsTable').empty();
                } else {
                    //次の構成にフォカス
                    $('#structure' + nextIndex).focus();
                }
            } else if ($('#searchResultsTable tr').length == 2) {
                //選択終了しか残されない場合は、次の構成に移動
                //オプション
                var symbolValue = $('#row0').find('td:first-child').map(function () {
                    return $(this).text();
                }).get();

                if (symbolValue == '') {
                    //フォカスの設定
                    if (nextIndex > structureNumber) {
                        //最後の場合はOKボタンにフォカス
                        $('#ok').focus();
                        $('#searchResultsTable').empty();
                    } else {
                        //次の構成にフォカス
                        $('#structure' + nextIndex).focus();
                    }
                }
            } else {
                //タイトル以外、まだ選択可能なオプションが存在するなら、フォカス移動しない
            }
        }
    }

</script>


