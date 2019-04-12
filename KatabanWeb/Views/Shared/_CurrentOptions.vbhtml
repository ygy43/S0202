@ModelType S0202.ViewModels.Options.OptionsUpdateOptionsViewModel

<div style="height: 400px; overflow:auto;">
    <table id="searchResultsTable" class="table table-bordered table-autowidth table-hover">
        @*タイトル*@
        <thead>
            <tr class="info">
                <th colspan="2" style="text-align:center;">@Model.StructureName</th>
            </tr>
        </thead>

        @*検索結果*@
        <tbody>
            @For Each result In Model.CurrentOptions
                @<tr class="clickable" id="row@(Model.CurrentOptions.IndexOf(result))" ondblclick="focusNext('@result.option_symbol','@Model.CurrentOptions.IndexOf(result)');">
                    <td>@result.option_symbol</td>
                    @If String.IsNullOrEmpty(result.option_nm) Then
                        @<td>@result.default_option_nm  </td>
                    Else
                        @<td>@result.option_nm</td>
                    End If
                </tr>
            Next
        </tbody>
    </table>
</div>