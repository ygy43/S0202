@ModelType S0202.ViewModels.Series.SeriesSearchViewModel
@Imports KatabanCommon.Constants.Divisions
@Imports PagedList
@Imports PagedList.Mvc

@Code
    ViewData("Title") = "機種検索"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h2>@ViewData("Title")</h2>

@Using Html.BeginForm("Search", "Series", FormMethod.Post, New With {.class = "form-inline text-left", .role = "form"})

    @Html.AntiForgeryToken()
    @<div class="row">
        <div class="form-group col-md-3">
            <div class="input-group">
                <span class="input-group-addon">形番</span>
                @Html.TextBoxFor(Function(m) m.Series, New With {.class = "form-control", .style = "text-transform:uppercase"})
            </div>
        </div>
        <div class="col-md-6">
            <div class="input-group">
                <span class="input-group-addon">検索区分</span>
                <div class="form-control">
                    @Html.RadioButtonFor(Function(m) m.SearchType, DataTypeDiv.Series, New With {.checked = "checked"})
                    <label>機種</label>
                    @Html.RadioButtonFor(Function(m) m.SearchType, DataTypeDiv.FullKataban)
                    <label>フル形番</label>
                    @Html.RadioButtonFor(Function(m) m.SearchType, DataTypeDiv.Shiire)
                    <label>仕入れ品</label>
                    @Html.RadioButtonFor(Function(m) m.SearchType, DataTypeDiv.All)
                    <label>全て</label>
                </div>
            </div>
        </div>
        <div class="col-md-3">
            <button type="submit" class="btn btn-primary">検索</button>
        </div>
    </div>
    @<hr />
    If Model.SearchResults IsNot Nothing
        @Html.Partial("_SeriesList", Model.SearchResults)
        @Html.PagedListPager(Model.SearchResults, Function(page) Url.Action("Search", "Series", New With {.series = Model.Series, .searchType = Model.SearchType, .page = page}))
    End If
End Using

<script type="text/javascript">
    /**
     *  画面ロード
     */
    $(document).ready(function () {
        if ($("#seriesList tr").length > 1) {
            //検索結果が存在する場合は一行目にフォカス
            $("#seriesList tr:eq(1)").trigger("click");
        } else {
            //デフォルトで機種項目にフォカス
            $("#Series").focus();
        }
    });
    //次の画面へ遷移
    function redirectToNextPage(selectedSeries,
        selectedKeyKataban,
        selectedCurrency,
        selectedSearchType) {

        var nextUrl = '@Url.Action("RedirectToNextPage", "Series")';
        window.location.href = nextUrl + "?series=" + selectedSeries + "&keyKataban=" + selectedKeyKataban + "&currency=" + selectedCurrency + "&searchType=" + selectedSearchType;
    };

</script>