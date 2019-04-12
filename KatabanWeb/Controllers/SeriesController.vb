Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Managers
Imports KatabanCommon.Constants
Imports PagedList
Imports S0202.Filters
Imports S0202.My.Resources
Imports S0202.MyHelpers
Imports S0202.ViewModels.Series

Namespace Controllers
    <AuthorizeFilter>
    Public Class SeriesController
        Inherits Controller

        ' POST: /Series/Search
        Function Search(series As String,
                        searchType As String,
                        sortOrder As String,
                        page As Integer?) As ActionResult

            'ViewBag.CurrentSort = sortOrder
            'ViewBag.CurrentFilter = series
            page = IIf(page Is Nothing, 1, page)

            Dim model As New SeriesSearchViewModel

            If Not String.IsNullOrEmpty(series) Then

                model.Series = series
                model.SearchType = searchType

                '機種を検索
                Dim seriesCount = SeriesManager.GetSeriesInfoCount(series, searchType, "ja", UserHelper.User.country_cd)

                'If seriesCount > 1000 Then
                '    '警告メッセージ
                'Else
                    Dim seriesInfos = SeriesManager.GetSeriesInfo(series,
                                                                  searchType,
                                                                  "ja",
                                                                  UserHelper.User.country_cd,
                                                                  page,
                                                                  Utility.PageSize)

                    model.SearchResults = New StaticPagedList(Of SeriesInfo)(seriesInfos, page, Utility.PageSize, seriesCount)
                End If
            'End If

            Return View(model)
        End Function

        ' POST: /Series/RedirectToNextPage
        Function RedirectToNextPage(series As String,
                                    keyKataban As String,
                                    currency As String,
                                    searchType As String) As ActionResult

            '選択した機種情報を保存
            Dim selectedSeriesInfo As SeriesInfo = SeriesManager.GetSeriesInfoByKey(series,
                                                                                    keyKataban,
                                                                                    currency,
                                                                                    searchType,
                                                                                    "ja",
                                                                                    UserHelper.User.country_cd)
            '選択した機種情報をセッションに保存
            Session(SessionKeys.SelectedSeriesData) = selectedSeriesInfo

            If selectedSeriesInfo.division = DataTypeDiv.Series Then

                '機種の場合はオプション選択画面へ遷移
                Return RedirectToAction("Index", "Options")
            Else
                'フル形番の場合は価格情報画面へ遷移
                Return RedirectToAction("Index", "Prices")
            End If

        End Function
    End Class
End Namespace