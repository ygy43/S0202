Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanCommon.Constants

Namespace Managers
    ''' <summary>
    '''     機種選択画面ビジネスロジック
    ''' </summary>
    Public Class SeriesManager
        ''' <summary>
        '''     機種情報の取得
        ''' </summary>
        ''' <param name="series">入力</param>
        ''' <param name="searchType">検索方法</param>
        ''' <param name="language">言語</param>
        ''' <param name="country">販売国</param>
        ''' <returns></returns>
        Public Shared Function GetSeriesInfo(series As String,
                                             searchType As String,
                                             language As String,
                                             country As String,
                                             page As Integer,
                                             pageSize As Integer) As List(Of SeriesInfo)

            Dim seriesInfos As New List(Of SeriesInfo)

            Using client As New DbAccessServiceClient

                Select Case searchType
                    Case Divisions.DataTypeDiv.Series
                        '機種検索
                        seriesInfos = client.SelectSeriesInfoBySeries(series, country, language, page, pageSize)
                    Case DataTypeDiv.FullKataban
                        'フル形番検索
                        Dim seriesInfoFullKataban = client.SelectSeriesInfoByFullKataban(series, country, language, page, pageSize)
                        seriesInfos = Convert2SeriesInfos(seriesInfoFullKataban)
                    Case DataTypeDiv.Shiire
                        '仕入れ品検索
                        seriesInfos = client.SelectSeriesInfoByShiire()
                End Select

            End Using

            If seriesInfos.Count = 0 Then
                '情報存在しない
                Return New List(Of SeriesInfo)
            Else
                '情報存在する
                Return seriesInfos
            End If
        End Function

        ''' <summary>
        '''     機種情報の取得
        ''' </summary>
        ''' <param name="series">入力</param>
        ''' <param name="searchType">検索方法</param>
        ''' <param name="language">言語</param>
        ''' <param name="country">販売国</param>
        ''' <returns></returns>
        Public Shared Function GetSeriesInfoCount(series As String,
                                             searchType As String,
                                             language As String,
                                             country As String) As Integer

            Dim result As Integer = 0

            Using client As New DbAccessServiceClient

                Select Case searchType
                    Case Divisions.DataTypeDiv.Series
                        '機種検索
                        result = client.SelectSeriesInfoCountBySeries(series, country, language)
                    Case DataTypeDiv.FullKataban
                        'フル形番検索
                        result = client.SelectSeriesInfoCountByFullKataban(series, country, language)
                    Case DataTypeDiv.Shiire
                        '仕入れ品検索
                        'result = client.SelectSeriesInfoByShiire()
                End Select

            End Using

            Return result
        End Function

        ''' <summary>
        '''     キーにより選択された機種情報を取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="currency">通貨</param>
        ''' <param name="searchType">検索タイプ</param>
        ''' <param name="language">言語</param>
        ''' <param name="country">販売国</param>
        ''' <returns></returns>
        Public Shared Function GetSeriesInfoByKey(series As String,
                                                  keyKataban As String,
                                                  currency As String,
                                                  searchType As String,
                                                  language As String,
                                                  country As String) As SeriesInfo
            Dim seriesData As New SeriesInfo

            Using client As New DbAccessServiceClient

                Select Case searchType
                    Case DataTypeDiv.Series
                        '機種検索
                        seriesData = client.SelectSeriesInfoWithKeyBySeries(series, keyKataban, country, language)
                    Case DataTypeDiv.FullKataban
                        'フル形番検索
                        Dim seriesDataFullKataban = client.SelectSeriesInfoWithKeyByFullKataban(series, currency,
                                                                                                country, language)
                        seriesData = Convert2SeriesInfos(
                            New List(Of SeriesInfoFullKataban) _
                                                            From {seriesDataFullKataban}).FirstOrDefault

                    Case DataTypeDiv.Shiire
                        '仕入れ品検索
                        seriesData = client.SelectSeriesInfoWithKeyByShiire()
                End Select

            End Using

            Return seriesData
        End Function

        ''' <summary>
        '''     フル形番情報を機種情報に変換
        ''' </summary>
        ''' <param name="infos"></param>
        ''' <returns></returns>
        Private Shared Function Convert2SeriesInfos(infos As List(Of SeriesInfoFullKataban)) As List(Of SeriesInfo)

            Dim result As New List(Of SeriesInfo)

            For Each info As SeriesInfoFullKataban In infos

                If String.IsNullOrEmpty(info.model_nm) AndAlso String.IsNullOrEmpty(info.parts_nm) Then
                    info.disp_name = "(システム)"
                ElseIf String.IsNullOrEmpty(info.model_nm) AndAlso Not String.IsNullOrEmpty(info.parts_nm) Then
                    info.disp_name = info.parts_nm
                ElseIf Not String.IsNullOrEmpty(info.model_nm) AndAlso String.IsNullOrEmpty(info.parts_nm) Then
                    info.disp_name = info.model_nm
                Else
                    info.disp_name = info.model_nm & MyControlChars.LeftBracket & info.parts_nm &
                                     MyControlChars.RightBracket
                End If

                If CInt(info.kataban_check_div) >= 4 Then
                    info.disp_name = "部品" & MyControlChars.LeftBracket & info.disp_name & MyControlChars.RightBracket
                End If

                Dim series As New SeriesInfo

                series.sort_key = info.sort_key
                series.series_kataban = info.series_kataban
                series.key_kataban = info.key_kataban
                series.hyphen_div = info.hyphen_div
                series.disp_kataban = info.disp_kataban
                series.division = info.division
                series.disp_name = info.disp_name
                series.price_no = info.price_no
                series.spec_no = info.spec_no
                series.currency_cd = info.currency_cd

                result.Add(series)
            Next

            Return result
        End Function
    End Class
End Namespace