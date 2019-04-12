@Imports PagedList
@Imports PagedList.Mvc
@ModelType IPagedList(Of KatabanBusinessLogic.KatabanWcfService.SeriesInfo)

<table id="seriesList" class="table table-bordered table-autowidth table-hover">
    @*タイトル*@
    <thead>
        <tr class="info">
            <th>機種</th>
            <th>説明</th>
        </tr>
    </thead>

    @*検索結果*@
    <tbody>
        @For Each result In Model
            @<tr class="clickable" ondblclick="redirectToNextPage('@result.series_kataban','@result.key_kataban','@result.currency_cd','@result.division');">
                <td>@result.disp_kataban</td>
                <td>@result.disp_name</td>
            </tr>
        Next
    </tbody>
</table>

