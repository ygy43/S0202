' NOTE: You can use the "Rename" command on the context menu to change the class name "DbService" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select DbService.svc or DbService.svc.vb at the Solution Explorer and start debugging.
Imports System.Data.SqlClient
Imports Dapper
Imports KatabanCommon.Constants
Imports S0202.Models

Public Class DbService
    Implements IDbAccessService

    Public Sub New()
    End Sub

#Region "ログイン画面関連"

    ''' <summary>
    '''     ユーザー情報の取得
    ''' </summary>
    ''' <param name="userId"></param>
    ''' <param name="password"></param>
    ''' <returns></returns>
    Public Function SelectUserMstByUserIdAndPassword(userId As String, password As String) As IEnumerable(Of UserInfo) _
        Implements IDbAccessService.SelectUserMstByUserIdAndPassword
        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  a.user_id, ")
            sql.Append("         b.base_cd, ")
            sql.Append("         a.country_cd, ")
            sql.Append("         a.office_cd, ")
            sql.Append("         a.person_cd, ")
            sql.Append("         a.mail_address, ")
            sql.Append("         b.language_cd, ")
            sql.Append("         b.currency_cd, ")
            sql.Append("         b.edit_div, ")
            sql.Append("         a.user_class, ")
            sql.Append("         a.price_disp_lvl, ")
            sql.Append("         a.add_information_lvl, ")
            sql.Append("         a.use_function_lvl, ")
            sql.Append("         a.current_datetime ")
            sql.Append(" FROM    kh_user_mst  a ")
            sql.Append(" INNER JOIN kh_country_mst  b ")
            sql.Append(" ON      a.country_cd = b.country_cd ")
            sql.Append(" WHERE   a.user_id = @userId ")
            sql.Append(" AND     a.password = @password ")
            sql.Append(" AND     a.in_effective_date          <= @standardDate ")
            sql.Append(" AND     a.out_effective_date          > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return connection.Query(Of UserInfo)(sql.ToString,
                                                  New With {
                                                     userId,
                                                     password,
                                                     standardDate
                                                     }
                                                  )

        End Using
    End Function

    ''' <summary>
    '''     パスワード更新
    ''' </summary>
    ''' <param name="userId">ユーザーID</param>
    ''' <param name="newPassword">パスワード</param>
    ''' <returns>更新行数</returns>
    Public Function UpdateUserMstPassword(userId As String, newPassword As String) As Integer _
        Implements IDbAccessService.UpdateUserMstPassword
        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" UPDATE  kh_user_mst ")
            sql.Append(" SET     password = @newPassword, ")
            sql.Append("         current_datetime = @currentTime ")
            sql.Append(" WHERE   user_id = @userId; ")

            connection.Open()

            Dim currentTime = Now

            Return connection.Execute(sql.ToString,
                                      New With {
                                         userId,
                                         newPassword,
                                         currentTime
                                         }
                                      )
        End Using
    End Function

#End Region

#Region "メニュー画面関連"

    ''' <summary>
    '''     更新履歴を取得
    ''' </summary>
    ''' <param name="language">言語コード</param>
    ''' <returns></returns>
    Public Function SelectInformationByLanguage(language As String) As IEnumerable(Of UpdateHistory) _
        Implements IDbAccessService.SelectInformationByLanguage
        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT   language_cd, ")
            sql.Append(" 		  message, ")
            sql.Append("          seq_no ")
            sql.Append(" FROM     kh_Information ")
            sql.Append(" WHERE    language_cd         = @language ")
            sql.Append(" AND      in_effective_date  <= @standardDate ")
            sql.Append(" AND      out_effective_date  > @standardDate ")
            sql.Append(" ORDER BY seq_no ")

            connection.Open()

            Dim standardDate = Now
            Return connection.Query(Of UpdateHistory)(sql.ToString,
                                                       New With {
                                                          language,
                                                          standardDate
                                                          }
                                                       )
        End Using
    End Function

#End Region

#Region "機種選択画面関連"

    ''' <summary>
    '''     機種により機種情報を取得
    ''' </summary>
    ''' <param name="language">言語コード</param>
    ''' <param name="input">入力した機種</param>
    ''' <param name="country">販売国コード</param>
    ''' <returns></returns>
    Private Function SelectSeriesInfoBySeries(input As String,
                                              country As String,
                                              language As String,
                                              page As Integer,
                                              pageSize As Integer) As IEnumerable(Of SeriesInfo) _
        Implements IDbAccessService.SelectSeriesInfoBySeries

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT DISTINCT KT.series_kataban + '_' + KT.key_kataban AS sort_key, ")
            sql.Append("                                           KT.series_kataban, ")
            sql.Append("                                           KT.key_kataban, ")
            sql.Append("                                           KT.hyphen_div, ")
            sql.Append(
                "                                           ISNULL(NM.disp_kataban,DF.disp_kataban) AS disp_kataban, ")
            sql.Append("                                           '1' AS division, ")
            sql.Append("                                           ISNULL(NM.series_nm,DF.series_nm) AS disp_name, ")
            sql.Append("                                           KT.price_no, ")
            sql.Append("                                           KT.spec_no, ")
            sql.Append("                                           KT.order_no, ")
            sql.Append("                                           KT.currency_cd ")
            sql.Append(" FROM ")
            sql.Append("            kh_series_kataban KT ")
            sql.Append(" INNER JOIN ")
            sql.Append("            kh_series_nm_mst DF ")
            sql.Append("            ON KT.series_kataban      = DF.series_kataban ")
            sql.Append("            AND KT.key_kataban        = DF.key_kataban ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_series_nm_mst NM ")
            sql.Append("            ON KT.series_kataban      = NM.series_kataban ")
            sql.Append("            AND KT.key_kataban        = NM.key_kataban ")
            sql.Append("            AND NM.language_cd        = @language ")
            sql.Append("            AND NM.in_effective_date <= @standardDate ")
            sql.Append("            AND NM.out_effective_date > @standardDate ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_country_group_mst DI ")
            sql.Append("            ON KT.country_group_cd    = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("      KT.in_effective_date <= @standardDate ")
            sql.Append("  AND KT.out_effective_date > @standardDate ")
            sql.Append("  AND DF.language_cd        = @defaultLanguage ")
            sql.Append("  AND DF.in_effective_date <= @standardDate ")
            sql.Append("  AND DF.out_effective_date > @standardDate ")
            sql.Append("  AND KT.series_kataban  LIKE @seriesKataban ")
            sql.Append("  AND (KT.country_group_cd = 'ALL' OR DI.country_cd = @country) ")
            sql.Append(" ORDER BY KT.series_kataban, ")
            sql.Append("          KT.order_no, ")
            sql.Append("          KT.key_kataban ")

            'ページング
            sql.Append(" OFFSET       @skipNumber ROWS ")
            sql.Append(" FETCH NEXT   @pageSize ROWS ONLY ")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Dim seriesKataban = input & "%"
            Dim skipNumber = (page - 1) * pageSize

            Return connection.Query(Of SeriesInfo)(sql.ToString,
                                                    New With {
                                                       language,
                                                       defaultLanguage,
                                                       standardDate,
                                                       seriesKataban,
                                                       country,
                                                       skipNumber,
                                                       pageSize
                                                       }
                                                    )
        End Using
    End Function

    ''' <summary>
    '''     機種により機種情報の件数を取得
    ''' </summary>
    ''' <param name="input"></param>
    ''' <param name="country"></param>
    ''' <param name="language"></param>
    ''' <returns></returns>
    Public Function SelectSeriesInfoCountBySeries(input As String, country As String, language As String) As Integer _
        Implements IDbAccessService.SelectSeriesInfoCountBySeries

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT COUNT(*) FROM ( ")
            sql.Append(" SELECT DISTINCT KT.series_kataban + '_' + KT.key_kataban AS sort_key, ")
            sql.Append("                                           KT.series_kataban, ")
            sql.Append("                                           KT.key_kataban, ")
            sql.Append("                                           KT.hyphen_div, ")
            sql.Append(
                "                                           ISNULL(NM.disp_kataban,DF.disp_kataban) AS disp_kataban, ")
            sql.Append("                                           '1' AS division, ")
            sql.Append("                                           ISNULL(NM.series_nm,DF.series_nm) AS disp_name, ")
            sql.Append("                                           KT.price_no, ")
            sql.Append("                                           KT.spec_no, ")
            sql.Append("                                           KT.order_no, ")
            sql.Append("                                           KT.currency_cd ")
            sql.Append(" FROM ")
            sql.Append("            kh_series_kataban KT ")
            sql.Append(" INNER JOIN ")
            sql.Append("            kh_series_nm_mst DF ")
            sql.Append("            ON KT.series_kataban      = DF.series_kataban ")
            sql.Append("            AND KT.key_kataban        = DF.key_kataban ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_series_nm_mst NM ")
            sql.Append("            ON KT.series_kataban      = NM.series_kataban ")
            sql.Append("            AND KT.key_kataban        = NM.key_kataban ")
            sql.Append("            AND NM.language_cd        = @language ")
            sql.Append("            AND NM.in_effective_date <= @standardDate ")
            sql.Append("            AND NM.out_effective_date > @standardDate ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_country_group_mst DI ")
            sql.Append("            ON KT.country_group_cd    = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("      KT.in_effective_date <= @standardDate ")
            sql.Append("  AND KT.out_effective_date > @standardDate ")
            sql.Append("  AND DF.language_cd        = @defaultLanguage ")
            sql.Append("  AND DF.in_effective_date <= @standardDate ")
            sql.Append("  AND DF.out_effective_date > @standardDate ")
            sql.Append("  AND KT.series_kataban  LIKE @seriesKataban ")
            sql.Append("  AND (KT.country_group_cd = 'ALL' OR DI.country_cd = @country) ")
            sql.Append(" ) A")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Dim seriesKataban = input & "%"

            Return connection.ExecuteScalar(Of Integer)(sql.ToString,
                                                         New With {
                                                            language,
                                                            defaultLanguage,
                                                            standardDate,
                                                            seriesKataban,
                                                            country
                                                            }
                                                         )

        End Using
    End Function

    ''' <summary>
    '''     キーにより機種情報の取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <param name="country">販売国</param>
    ''' <param name="language">言語</param>
    ''' <returns></returns>
    Public Function SelectSeriesInfoWithKeyBySeries(series As String,
                                                    keyKataban As String,
                                                    country As String,
                                                    language As String) As SeriesInfo _
        Implements IDbAccessService.SelectSeriesInfoWithKeyBySeries

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT DISTINCT KT.series_kataban + '_' + KT.key_kataban AS sort_key, ")
            sql.Append("                                           KT.series_kataban, ")
            sql.Append("                                           KT.key_kataban, ")
            sql.Append("                                           KT.hyphen_div, ")
            sql.Append(
                "                                           ISNULL(NM.disp_kataban,DF.disp_kataban) AS disp_kataban, ")
            sql.Append("                                           '1' AS division, ")
            sql.Append("                                           ISNULL(NM.series_nm,DF.series_nm) AS disp_name, ")
            sql.Append("                                           KT.price_no, ")
            sql.Append("                                           KT.spec_no, ")
            sql.Append("                                           KT.order_no, ")
            sql.Append("                                           KT.currency_cd ")
            sql.Append(" FROM ")
            sql.Append("            kh_series_kataban KT ")
            sql.Append(" INNER JOIN ")
            sql.Append("            kh_series_nm_mst DF ")
            sql.Append("            ON KT.series_kataban      = DF.series_kataban ")
            sql.Append("            AND KT.key_kataban        = DF.key_kataban ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_series_nm_mst NM ")
            sql.Append("            ON KT.series_kataban      = NM.series_kataban ")
            sql.Append("            AND KT.key_kataban        = NM.key_kataban ")
            sql.Append("            AND NM.language_cd        = @language ")
            sql.Append("            AND NM.in_effective_date <= @standardDate ")
            sql.Append("            AND NM.out_effective_date > @standardDate ")
            sql.Append(" LEFT OUTER JOIN ")
            sql.Append("            kh_country_group_mst DI ")
            sql.Append("            ON KT.country_group_cd    = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("      KT.in_effective_date <= @standardDate ")
            sql.Append("  AND KT.out_effective_date > @standardDate ")
            sql.Append("  AND DF.language_cd        = @defaultLanguage ")
            sql.Append("  AND DF.in_effective_date <= @standardDate ")
            sql.Append("  AND DF.out_effective_date > @standardDate ")
            sql.Append("  AND KT.series_kataban     = @seriesKataban ")
            sql.Append("  AND KT.key_kataban        = @keyKataban ")
            sql.Append("  AND (KT.country_group_cd = 'ALL' OR DI.country_cd = @country) ")
            sql.Append(" ORDER BY KT.series_kataban, ")
            sql.Append("          KT.order_no, ")
            sql.Append("          KT.key_kataban ")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Dim seriesKataban = series
            Return connection.Query(Of SeriesInfo)(sql.ToString,
                                                    New With {
                                                       language,
                                                       defaultLanguage,
                                                       standardDate,
                                                       seriesKataban,
                                                       keyKataban,
                                                       country
                                                       }
                                                    ).FirstOrDefault()
        End Using
    End Function

    ''' <summary>
    '''     フル形番により機種情報を取得
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectSeriesInfoByFullKataban(input As String,
                                                  country As String,
                                                  language As String,
                                                  page As Integer,
                                                  pageSize As Integer) As IEnumerable(Of SeriesInfoFullKataban) _
        Implements IDbAccessService.SelectSeriesInfoByFullKataban

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT ")
            sql.Append("        PRC.kataban + CONVERT(VARCHAR, PRC.in_effective_date, 120) AS sort_key, ")
            'sql.Append("        ' ' AS series_kataban, ")
            sql.Append("        PRC.kataban AS series_kataban, ")
            sql.Append("        ' ' AS key_kataban, ")
            sql.Append("        ' ' AS hyphen_div, ")
            sql.Append("        PRC.kataban AS disp_kataban, ")
            sql.Append("        '2' AS division, ")
            sql.Append("        PRC.kataban_check_div, ")
            sql.Append("        PNM.parts_nm, ")
            sql.Append("        PNM.model_nm, ")
            sql.Append("        '' AS disp_name, ")
            sql.Append("        '' AS price_no, ")
            sql.Append("        '' AS spec_no, ")
            sql.Append("        PRC.currency_cd ")
            sql.Append(" FROM  ")
            sql.Append("        kh_price PRC ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_parts_nm_mst PNM  ")
            sql.Append("        ON PRC.kataban = PNM.kataban ")
            sql.Append("        AND PNM.language_cd = @language ")
            sql.Append("        AND PNM.in_effective_date <= @standardDate ")
            sql.Append("        AND PNM.out_effective_date > @standardDate ")
            sql.Append("        AND PRC.currency_cd = PNM.currency_cd ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_country_group_mst DI  ")
            sql.Append("        ON PRC.country_group_cd = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("       PRC.kataban LIKE @seriesKataban ")
            sql.Append("       AND PRC.in_effective_date <= @standardDate ")
            sql.Append("       AND PRC.out_effective_date > @standardDate ")
            sql.Append("       AND (PRC.country_group_cd = 'ALL' OR DI.country_cd = @country) ")
            sql.Append(" ORDER BY PRC.kataban ")

            'ページング
            sql.Append(" OFFSET       @skipNumber ROWS ")
            sql.Append(" FETCH NEXT   @pageSize ROWS ONLY ")

            connection.Open()

            Dim standardDate = Now
            Dim seriesKataban = input & "%"
            Dim skipNumber = (page - 1) * pageSize

            Return connection.Query(Of SeriesInfoFullKataban)(sql.ToString,
                                                               New With {
                                                                  language,
                                                                  standardDate,
                                                                  seriesKataban,
                                                                  country,
                                                                  skipNumber,
                                                                  pageSize
                                                                  }
                                                               )
        End Using
    End Function

    ''' <summary>
    '''     フル形番により機種情報の件数を取得
    ''' </summary>
    ''' <param name="input"></param>
    ''' <param name="country"></param>
    ''' <param name="language"></param>
    ''' <returns></returns>
    Public Function SelectSeriesInfoCountByFullKataban(input As String, country As String, language As String) _
        As Integer Implements IDbAccessService.SelectSeriesInfoCountByFullKataban
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT COUNT(*) ")
            sql.Append(" FROM  ")
            sql.Append("        kh_price PRC ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_parts_nm_mst PNM  ")
            sql.Append("        ON PRC.kataban = PNM.kataban ")
            sql.Append("        AND PNM.language_cd = @language ")
            sql.Append("        AND PNM.in_effective_date <= @standardDate ")
            sql.Append("        AND PNM.out_effective_date > @standardDate ")
            sql.Append("        AND PRC.currency_cd = PNM.currency_cd ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_country_group_mst DI  ")
            sql.Append("        ON PRC.country_group_cd = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("       PRC.kataban LIKE @seriesKataban ")
            sql.Append("       AND PRC.in_effective_date <= @standardDate ")
            sql.Append("       AND PRC.out_effective_date > @standardDate ")
            sql.Append("       AND (PRC.country_group_cd = 'ALL' OR DI.country_cd = @country) ")

            connection.Open()

            Dim standardDate = Now
            Dim seriesKataban = input & "%"

            Return connection.ExecuteScalar(Of Integer)(sql.ToString,
                                                         New With {
                                                            language,
                                                            standardDate,
                                                            seriesKataban,
                                                            country
                                                            }
                                                         )
        End Using
    End Function

    ''' <summary>
    '''     キーによりフル形番機種情報の取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="currency">通貨</param>
    ''' <param name="country">販売国</param>
    ''' <param name="language">言語</param>
    ''' <returns></returns>
    Public Function SelectSeriesInfoWithKeyByFullKataban(series As String,
                                                         currency As String,
                                                         country As String,
                                                         language As String) As SeriesInfoFullKataban _
        Implements IDbAccessService.SelectSeriesInfoWithKeyByFullKataban

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT TOP 15 ")
            sql.Append("        PRC.kataban + CONVERT(VARCHAR, PRC.in_effective_date, 120) AS sort_key, ")
            'sql.Append("        ' ' AS series_kataban, ")
            sql.Append("        PRC.kataban AS series_kataban, ")
            sql.Append("        ' ' AS key_kataban, ")
            sql.Append("        ' ' AS hyphen_div, ")
            sql.Append("        PRC.kataban AS disp_kataban, ")
            sql.Append("        '2' AS division, ")
            sql.Append("        PRC.kataban_check_div, ")
            sql.Append("        PNM.parts_nm, ")
            sql.Append("        PNM.model_nm, ")
            sql.Append("        '' AS disp_name, ")
            sql.Append("        '' AS price_no, ")
            sql.Append("        '' AS spec_no, ")
            sql.Append("        PRC.currency_cd ")
            sql.Append(" FROM  ")
            sql.Append("        kh_price PRC ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_parts_nm_mst PNM  ")
            sql.Append("        ON PRC.kataban = PNM.kataban ")
            sql.Append("        AND PNM.language_cd = @language ")
            sql.Append("        AND PNM.in_effective_date <= @standardDate ")
            sql.Append("        AND PNM.out_effective_date > @standardDate ")
            sql.Append("        AND PRC.currency_cd = PNM.currency_cd ")
            sql.Append(" LEFT OUTER JOIN  ")
            sql.Append("        kh_country_group_mst DI  ")
            sql.Append("        ON PRC.country_group_cd = DI.country_group_cd ")
            sql.Append(" WHERE ")
            sql.Append("       PRC.kataban = @seriesKataban ")
            sql.Append("       AND PRC.currency_cd = @currency ")
            sql.Append("       AND PRC.in_effective_date <= @standardDate ")
            sql.Append("       AND PRC.out_effective_date > @standardDate ")
            sql.Append("       AND (PRC.country_group_cd = 'ALL' OR DI.country_cd = @country) ")

            connection.Open()

            Dim standardDate = Now
            Dim seriesKataban = series
            Return connection.Query(Of SeriesInfoFullKataban)(sql.ToString,
                                                               New With {
                                                                  language,
                                                                  standardDate,
                                                                  seriesKataban,
                                                                  currency,
                                                                  country
                                                                  }
                                                               ).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     仕入れ品の機種情報を取得
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectSeriesInfoByShiire() As IEnumerable(Of SeriesInfo) _
        Implements IDbAccessService.SelectSeriesInfoByShiire
        Throw New NotImplementedException()
    End Function

    ''' <summary>
    '''     仕入れ品
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectSeriesInfoWithKeyByShiire() As SeriesInfo _
        Implements IDbAccessService.SelectSeriesInfoWithKeyByShiire
        Throw New NotImplementedException()
    End Function

#End Region

#Region "オプション選択画面関連"

    ''' <summary>
    '''     形番構成情報の取得
    ''' </summary>
    ''' <param name="language">言語</param>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <returns></returns>
    Public Function SelectKatabanStructure(series As String, keyKataban As String, language As String) _
        As IEnumerable(Of KatabanStructureInfo) Implements IDbAccessService.SelectKatabanStructure

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  a.ktbn_strc_seq_no, ")
            sql.Append("         a.element_div, ")
            sql.Append("         a.structure_div, ")
            sql.Append("         a.addition_div, ")
            sql.Append("         a.hyphen_div, ")
            sql.Append("         b.ktbn_strc_nm as default_nm, ")
            sql.Append("         c.ktbn_strc_nm ")
            sql.Append(" FROM    kh_kataban_strc a ")
            sql.Append(" INNER JOIN  kh_ktbn_strc_nm_mst b ")
            sql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sql.Append(" AND     a.ktbn_strc_seq_no    = b.ktbn_strc_seq_no ")
            sql.Append(" AND     b.language_cd         = @defaultLanguage ")
            sql.Append(" AND     b.in_effective_date  <= @standardDate ")
            sql.Append(" AND     b.out_effective_date  > @standardDate ")
            sql.Append(" LEFT  JOIN  kh_ktbn_strc_nm_mst c ")
            sql.Append(" ON      a.series_kataban      = c.series_kataban ")
            sql.Append(" AND     a.key_kataban         = c.key_kataban ")
            sql.Append(" AND     a.ktbn_strc_seq_no    = c.ktbn_strc_seq_no ")
            sql.Append(" AND     c.language_cd         = @language ")
            sql.Append(" AND     c.in_effective_date  <= @standardDate ")
            sql.Append(" AND     c.out_effective_date  > @standardDate ")
            sql.Append(" WHERE   a.series_kataban      = @series ")
            sql.Append(" AND     a.key_kataban         = @keyKataban ")
            sql.Append(" AND     a.in_effective_date  <= @standardDate ")
            sql.Append(" AND     a.out_effective_date  > @standardDate ")
            sql.Append(" ORDER BY  a.ktbn_strc_seq_no ")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Return connection.Query(Of KatabanStructureInfo)(sql.ToString,
                                                              New With {
                                                                 series,
                                                                 keyKataban,
                                                                 standardDate,
                                                                 defaultLanguage,
                                                                 language
                                                                 }
                                                              )
        End Using
    End Function

    ''' <summary>
    '''     形番構成オプションの取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <returns></returns>
    Public Function SelectKatabanStructureOptions(series As String, keyKataban As String) _
        As List(Of KatabanStructureOptionInfoAllSeqNo) _
        Implements IDbAccessService.SelectKatabanStructureOptions

        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  ktbn_strc_seq_no, ")
            sql.Append("         option_symbol, ")
            sql.Append("         place_lvl ")
            sql.Append(" FROM    kh_kataban_strc_ele ")
            sql.Append(" WHERE   series_kataban      = @series ")
            sql.Append(" AND     key_kataban         = @keyKataban ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")
            sql.Append(" ORDER BY  ktbn_strc_seq_no ")

            connection.Open()

            Dim standardDate = Now
            Return connection.Query(Of KatabanStructureOptionInfoAllSeqNo)(sql.ToString,
                                                                            New With {
                                                                               series,
                                                                               keyKataban,
                                                                               standardDate
                                                                               }
                                                                            )
        End Using
    End Function

    ''' <summary>
    '''     形番構成オプション情報の取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <returns></returns>
    Public Function SelectKatabanStructureOptionsBySeqNo(series As String, keyKataban As String, seqNo As Integer,
                                                         language As String) _
        As IEnumerable(Of KatabanStructureOptionInfo) Implements IDbAccessService.SelectKatabanStructureOptionsBySeqNo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  A.option_symbol, ")
            sql.Append("         B.option_nm AS option_nm, ")
            sql.Append("         C.option_nm AS default_option_nm, ")
            sql.Append("         A.disp_seq_no, ")
            sql.Append("         A.price_acc_div, ")
            sql.Append("         A.place_lvl ")
            sql.Append(" FROM    kh_kataban_strc_ele A ")
            sql.Append(" INNER JOIN    kh_option_nm_mst B ")
            sql.Append(" ON      A.series_kataban         = B.series_kataban ")
            sql.Append(" AND     A.key_kataban            = B.key_kataban ")
            sql.Append(" AND     A.ktbn_strc_seq_no       = B.ktbn_strc_seq_no ")
            sql.Append(" AND     A.option_symbol          = B.option_symbol ")
            sql.Append(" AND     B.series_kataban         = @series ")
            sql.Append(" AND     B.key_kataban            = @keyKataban ")
            sql.Append(" AND     B.ktbn_strc_seq_no       = @seqNo ")
            sql.Append(" AND     B.language_cd            = @language ")
            sql.Append(" AND     B.in_effective_date     <= @standardDate ")
            sql.Append(" AND     B.out_effective_date     > @standardDate ")
            sql.Append(" LEFT  JOIN  kh_option_nm_mst C ")
            sql.Append(" ON      A.series_kataban         = C.series_kataban ")
            sql.Append(" AND     A.key_kataban            = C.key_kataban ")
            sql.Append(" AND     A.ktbn_strc_seq_no       = C.ktbn_strc_seq_no ")
            sql.Append(" AND     A.option_symbol          = C.option_symbol ")
            sql.Append(" AND     C.series_kataban         = @series ")
            sql.Append(" AND     C.key_kataban            = @keyKataban ")
            sql.Append(" AND     C.ktbn_strc_seq_no       = @seqNo ")
            sql.Append(" AND     C.language_cd            = @DefaultLanguage ")
            sql.Append(" AND     C.in_effective_date     <= @standardDate ")
            sql.Append(" AND     C.out_effective_date     > @standardDate ")

            sql.Append(" WHERE   A.series_kataban      = @series ")
            sql.Append(" AND     A.key_kataban         = @keyKataban ")
            sql.Append(" AND     A.ktbn_strc_seq_no    = @seqNo ")
            sql.Append(" AND     A.in_effective_date  <= @standardDate ")
            sql.Append(" AND     A.out_effective_date  > @standardDate ")
            sql.Append(" ORDER BY  A.disp_seq_no ")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Return _
                connection.Query(Of KatabanStructureOptionInfo)(sql.ToString,
                                                                 New With {
                                                                    series,
                                                                    keyKataban,
                                                                    seqNo,
                                                                    standardDate,
                                                                    defaultLanguage,
                                                                    language
                                                                    }
                                                                 )

        End Using
    End Function

    ''' <summary>
    '''     形番構成オプション検証情報の取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <param name="seqNo">構成番号</param>
    ''' <returns></returns>
    Public Function SelectElePatternInfoAll(series As String, keyKataban As String, seqNo As Integer) _
        As IEnumerable(Of ElePatternInfo) Implements IDbAccessService.SelectElePatternInfoAll
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  '1' as search_seq_no, ")
            sql.Append("         option_symbol, ")
            sql.Append("         condition_cd, ")
            sql.Append("         condition_seq_no, ")
            sql.Append("         condition_seq_no_br, ")
            sql.Append("         cond_option_symbol ")
            sql.Append(" FROM    kh_ele_pattern ")
            sql.Append(" WHERE   series_kataban      = @series ")
            sql.Append(" AND     key_kataban         = @keyKataban ")
            sql.Append(" AND     ktbn_strc_seq_no    = @seqNo ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")
            sql.Append(" AND     option_symbol       = '" & ElePatternDiv.Plural & "' ")
            sql.Append(" UNION ")
            sql.Append(" SELECT  '2' as search_seq_no, ")
            sql.Append("         option_symbol, ")
            sql.Append("         condition_cd, ")
            sql.Append("         condition_seq_no, ")
            sql.Append("         condition_seq_no_br, ")
            sql.Append("         cond_option_symbol ")
            sql.Append(" FROM    kh_ele_pattern ")
            sql.Append(" WHERE   series_kataban      = @series ")
            sql.Append(" AND     key_kataban         = @keyKataban ")
            sql.Append(" AND     ktbn_strc_seq_no    = @seqNo ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")
            sql.Append(" AND     option_symbol       = '" & ElePatternDiv.All & "' ")
            sql.Append(" UNION ")
            sql.Append(" SELECT  '3' as search_seq_no, ")
            sql.Append("         option_symbol, ")
            sql.Append("         condition_cd, ")
            sql.Append("         condition_seq_no, ")
            sql.Append("         condition_seq_no_br, ")
            sql.Append("         cond_option_symbol ")
            sql.Append(" FROM    kh_ele_pattern ")
            sql.Append(" WHERE   series_kataban      = @series ")
            sql.Append(" AND     key_kataban         = @keyKataban ")
            sql.Append(" AND     ktbn_strc_seq_no    = @seqNo ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")
            sql.Append(" AND     option_symbol  Not In ('" & ElePatternDiv.All & "','" & ElePatternDiv.Plural & "') ")
            sql.Append(" ORDER BY  search_seq_no, option_symbol, condition_seq_no, condition_seq_no_br ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of ElePatternInfo)(sql.ToString,
                                                     New With {
                                                        series,
                                                        keyKataban,
                                                        seqNo,
                                                        standardDate
                                                        }
                                                     )

        End Using
    End Function

    ''' <summary>
    '''     複数選択可能なオプションの検証情報を取得
    ''' </summary>
    ''' <param name="series"></param>
    ''' <param name="keyKataban"></param>
    ''' <param name="seqNo"></param>
    ''' <returns></returns>
    Public Function SelectElePatternInfoPlural(series As String, keyKataban As String, seqNo As Integer) _
        As IEnumerable(Of ElePatternInfo) Implements IDbAccessService.SelectElePatternInfoPlural
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  '1' as search_seq_no, ")
            sql.Append("         option_symbol, ")
            sql.Append("         condition_cd, ")
            sql.Append("         condition_seq_no, ")
            sql.Append("         condition_seq_no_br, ")
            sql.Append("         cond_option_symbol ")
            sql.Append(" FROM    kh_ele_pattern ")
            sql.Append(" WHERE   series_kataban      = @series ")
            sql.Append(" AND     key_kataban         = @keyKataban ")
            sql.Append(" AND     ktbn_strc_seq_no    = @seqNo ")
            sql.Append(" AND     option_symbol       = @OptionSymbol ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")
            sql.Append(" ORDER BY  condition_seq_no_br ")

            connection.Open()

            Dim standardDate = Now
            Dim optionSymbol = ElePatternDiv.Plural

            Return _
                connection.Query(Of ElePatternInfo)(sql.ToString,
                                                     New With {
                                                        series,
                                                        keyKataban,
                                                        seqNo,
                                                        optionSymbol,
                                                        standardDate
                                                        }
                                                     )

        End Using
    End Function

    ''' <summary>
    '''     ロッド先端マスタの情報を取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <returns></returns>
    Public Function SelectRodEndInfo(series As String, keyKataban As String) As IEnumerable(Of RodEndInfo) _
        Implements IDbAccessService.SelectRodEndInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  pattern_seq_no, ")
            sql.Append("         rod_pattern_symbol, ")
            sql.Append("         ISNULL(url , '') AS url ")
            sql.Append(" FROM    kh_rod_end_mst ")
            sql.Append(" WHERE   series_kataban         = @series ")
            sql.Append(" AND     key_kataban            = @keyKataban ")
            sql.Append(" ORDER BY  pattern_seq_no ")

            connection.Open()

            Return _
                connection.Query(Of RodEndInfo)(sql.ToString,
                                                 New With {
                                                    series,
                                                    keyKataban
                                                    }
                                                 )

        End Using
    End Function

    ''' <summary>
    '''     ロッド先端特注外径種類/ロッド先端特注標準寸法検索取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <param name="boreSize">口径</param>
    ''' <returns></returns>
    Public Function SelectRodEndExternalFormInfo(series As String, keyKataban As String, boreSize As Integer) _
        As IEnumerable(Of RodEndExternalFormInfo) Implements IDbAccessService.SelectRodEndExternalFormInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)
            try

            Dim sql As New StringBuilder

            sql.Append(" SELECT  b.pattern_seq_no, ")
            sql.Append("         a.rod_pattern_symbol, ")
            sql.Append("         a.external_form, ")
            sql.Append("         a.disp_external_form, ")
            sql.Append("         a.input_div, ")
            sql.Append("         a.js_name, ")
            sql.Append("         c.normal_value, ")
            sql.Append("         c.act_normal_value , ")
            sql.Append("         d.selectable_value,  ")
            sql.Append("         d.act_selectable_value ")
            'sql.Append("         e.wf_max_value ")
            sql.Append(" FROM    kh_rod_end_ext_frm a ")
            sql.Append(" INNER JOIN  kh_rod_end_mst b ")
            sql.Append(" ON      a.series_kataban         = b.series_kataban ")
            sql.Append(" AND     a.key_kataban            = b.key_kataban ")
            sql.Append(" AND     a.rod_pattern_symbol     = b.rod_pattern_symbol ")
            sql.Append(" INNER JOIN  kh_rod_end_std_size c ")
            sql.Append(" ON      a.series_kataban         = c.series_kataban ")
            sql.Append(" AND     a.key_kataban            = c.key_kataban ")
            sql.Append(" AND     a.rod_pattern_symbol     = c.rod_pattern_symbol ")
            sql.Append(" AND     a.external_form          = c.external_form ")
            sql.Append(" AND     c.bore_size              = @boreSize ")
            sql.Append(" LEFT  JOIN  kh_rod_end_selectable_size d")
            sql.Append(" ON      a.series_kataban         = d.series_kataban ")
            sql.Append(" AND     a.key_kataban            = d.key_kataban ")
            sql.Append(" AND     a.rod_pattern_symbol     = d.rod_pattern_symbol ")
            sql.Append(" AND     a.external_form          = d.external_form ")
            sql.Append(" AND     d.bore_size              = @boreSize ")
            'sql.Append(" LEFT  JOIN  kh_rod_end_wf_max_size e")
            'sql.Append(" ON      a.series_kataban         = e.series_kataban ")
            'sql.Append(" AND     a.key_kataban            = e.key_kataban ")
            'sql.Append(" AND     e.bore_size              = @boreSize ")
            sql.Append(" WHERE   a.series_kataban         = @series ")
            sql.Append(" AND     a.key_kataban            = @keyKataban ")
            sql.Append(" ORDER BY  b.pattern_seq_no, a.external_form_seq_no, d.sel_value_seq_no")

            connection.Open()

            Return _
                connection.Query(Of RodEndExternalFormInfo)(sql.ToString,
                                                             New With {
                                                                series,
                                                                keyKataban,
                                                                boreSize
                                                                }
                                                             )
                
            Catch ex As Exception
                dim test = ex.Message
            End Try
        End Using
    End Function

    ''' <summary>
    '''     WF最大値を取得
    ''' </summary>
    ''' <param name="series"></param>
    ''' <param name="keyKataban"></param>
    ''' <param name="boreSize">口径</param>
    ''' <returns></returns>
    Public Function SelectWfMaxValue(series As String, keyKataban As String, boreSize As Integer) As String _
        Implements IDbAccessService.SelectWfMaxValue
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  wf_max_value ")
            sql.Append(" FROM    kh_rod_end_wf_max_size ")
            sql.Append(" WHERE   series_kataban         = @series ")
            sql.Append(" AND     key_kataban            = @keyKataban ")
            sql.Append(" AND     bore_size              = @boreSize ")

            connection.Open()

            Return _
                connection.ExecuteScalar(Of String)(sql.ToString,
                                                     New With {
                                                        series,
                                                        keyKataban,
                                                        boreSize
                                                        }
                                                     )
        End Using
    End Function

#End Region

#Region "製品情報画面関連"

    ''' <summary>
    '''     フル形番価格情報を取得
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectFullKatabanPriceInfo(kataban As String, currency As String) As PriceInfo _
        Implements IDbAccessService.SelectFullKatabanPriceInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  kataban, ")
            sql.Append("         kataban_check_div, ")
            sql.Append("         place_cd, ")
            sql.Append("         ls_price, ")
            sql.Append("         rg_price, ")
            sql.Append("         ss_price, ")
            sql.Append("         bs_price, ")
            sql.Append("         gs_price, ")
            sql.Append("         ps_price, ")
            sql.Append("         currency_cd, ")
            sql.Append("         country_group_cd, ")
            sql.Append("         country_cd ")
            sql.Append(" FROM    kh_price ")
            sql.Append(" WHERE   kataban             = @kataban ")
            sql.Append(" AND     currency_cd         = @currency ")
            sql.Append(" AND     in_effective_date  <= @StandardDate ")
            sql.Append(" AND     out_effective_date  > @StandardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of PriceInfo)(sql.ToString,
                                                New With {
                                                   kataban,
                                                   currency,
                                                   standardDate
                                                   }
                                                ).FirstOrDefault

        End Using
    End Function

    ''' <summary>
    '''     積上げ価格情報を取得
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectAccumulatePriceInfo(kataban As String, currency As String) As PriceInfo _
        Implements IDbAccessService.SelectAccumulatePriceInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  kataban, ")
            sql.Append("         kataban_check_div, ")
            sql.Append("         place_cd, ")
            sql.Append("         ls_price, ")
            sql.Append("         rg_price, ")
            sql.Append("         ss_price, ")
            sql.Append("         bs_price, ")
            sql.Append("         gs_price, ")
            sql.Append("         ps_price  ")
            sql.Append(" FROM    kh_accumulate_price ")
            sql.Append(" WHERE   kataban             = @kataban ")
            sql.Append(" AND     currency_cd         = @currency ")
            sql.Append(" AND     in_effective_date  <= @standardDate ")
            sql.Append(" AND     out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of PriceInfo)(sql.ToString,
                                                New With {
                                                   kataban,
                                                   currency,
                                                   standardDate
                                                   }
                                                ).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     ねじ加算価格情報を取得
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectScrewPriceInfo(kataban As String) As ScrewPriceInfo _
        Implements IDbAccessService.SelectScrewPriceInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)
            Dim sql As New StringBuilder

            sql.Append(" SELECT  kataban, ")
            sql.Append("         ls_price, ")
            sql.Append("         rg_price, ")
            sql.Append("         ss_price, ")
            sql.Append("         bs_price, ")
            sql.Append("         gs_price, ")
            sql.Append("         ps_price ")
            sql.Append(" FROM    kh_price ")
            sql.Append(" WHERE   kataban = @kataban ")

            connection.Open()

            Return _
                connection.Query(Of ScrewPriceInfo)(sql.ToString,
                                                     New With {
                                                        kataban
                                                        }
                                                     ).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     プラス
    ''' </summary>
    ''' <param name="price1"></param>
    ''' <param name="price2"></param>
    ''' <returns></returns>
    Public Function AddPriceInfo(price1 As PriceInfo, price2 As PriceInfo) As PriceInfo _
        Implements IDbAccessService.AddPriceInfo

        Dim checkdiv1 As Integer = IIf(String.IsNullOrEmpty(price1.kataban_check_div), 0,
                                       CType(price1.kataban_check_div, Integer))
        Dim checkdiv2 As Integer = IIf(String.IsNullOrEmpty(price2.kataban_check_div), 0,
                                       CType(price2.kataban_check_div, Integer))

        Return New PriceInfo With {
            .kataban = IIf(String.IsNullOrEmpty(price1.kataban), price2.kataban, price1.kataban),
            .currency_cd = IIf(String.IsNullOrEmpty(price1.currency_cd), price2.currency_cd, price1.currency_cd),
            .kataban_check_div = Math.Max(checkdiv1, checkdiv2),
            .place_cd = price1.place_cd,
            .ls_price = price1.ls_price + price2.ls_price,
            .rg_price = price1.rg_price + price2.rg_price,
            .ss_price = price1.ss_price + price2.ss_price,
            .bs_price = price1.bs_price + price2.bs_price,
            .gs_price = price1.gs_price + price2.gs_price,
            .ps_price = price1.ps_price + price2.ps_price
            }
    End Function

    ''' <summary>
    '''     マイナス
    ''' </summary>
    ''' <param name="price1"></param>
    ''' <param name="price2"></param>
    ''' <returns></returns>
    Public Function MinusPriceInfo(price1 As PriceInfo, price2 As PriceInfo) As PriceInfo _
        Implements IDbAccessService.MinusPriceInfo

        Dim checkdiv1 As Integer = IIf(String.IsNullOrEmpty(price1.kataban_check_div), 0,
                                       CType(price1.kataban_check_div, Integer))
        Dim checkdiv2 As Integer = IIf(String.IsNullOrEmpty(price2.kataban_check_div), 0,
                                       CType(price2.kataban_check_div, Integer))

        Return New PriceInfo With {
            .kataban = IIf(String.IsNullOrEmpty(price1.kataban), price2.kataban, price1.kataban),
            .currency_cd = IIf(String.IsNullOrEmpty(price1.currency_cd), price2.currency_cd, price1.currency_cd),
            .kataban_check_div = Math.Max(checkdiv1, checkdiv2),
            .place_cd = price1.place_cd,
            .ls_price = price1.ls_price - price2.ls_price,
            .rg_price = price1.rg_price - price2.rg_price,
            .ss_price = price1.ss_price - price2.ss_price,
            .bs_price = price1.bs_price - price2.bs_price,
            .gs_price = price1.gs_price - price2.gs_price,
            .ps_price = price1.ps_price - price2.ps_price
            }
    End Function

    ''' <summary>
    '''     掛け算
    ''' </summary>
    ''' <param name="price"></param>
    ''' <param name="ratio"></param>
    ''' <returns></returns>
    Public Function MultiplePriceInfo(price As PriceInfo, ratio As List(Of Decimal)) As PriceInfo _
        Implements IDbAccessService.MultiplePriceInfo

        Return New PriceInfo With {
            .kataban = price.kataban,
            .currency_cd = price.currency_cd,
            .kataban_check_div = price.kataban_check_div,
            .place_cd = price.place_cd,
            .ls_price = price.ls_price * ratio(0),
            .rg_price = price.rg_price * ratio(1),
            .ss_price = price.ss_price * ratio(2),
            .bs_price = price.bs_price * ratio(3),
            .gs_price = price.gs_price * ratio(4),
            .ps_price = price.ps_price * ratio(5)
            }
    End Function

    ''' <summary>
    '''     現地定価為替レートの取得
    ''' </summary>
    ''' <param name="katabanCurrency">形番通貨</param>
    ''' <param name="userCurrency">ユーザー通貨</param>
    ''' <returns></returns>
    Public Function SelectExchangeRate(katabanCurrency As String, userCurrency As String) As Decimal _
        Implements IDbAccessService.SelectExchangeRate
        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT ")
            sql.Append("     exchange_rate ")
            sql.Append(" FROM ")
            sql.Append("     kh_currency_exc_rate_mst ")
            sql.Append(" WHERE    base_currency_cd    = @katabanCurrency ")
            sql.Append(" AND      change_currency_cd  = @userCurrency ")
            sql.Append(" AND      in_effective_date  <= @standardDate ")
            sql.Append(" AND      out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.ExecuteScalar(Of Decimal)(sql.ToString,
                                                      New With {
                                                         katabanCurrency,
                                                         userCurrency,
                                                         standardDate
                                                         }
                                                      )

        End Using
    End Function

    ''' <summary>
    '''     現地定価端数処理方法の取得
    ''' </summary>
    ''' <param name="userCountry">ユーザー国コード</param>
    ''' <param name="katabanFirstHyphen">形番第一ハイフン</param>
    ''' <returns></returns>
    Public Function SelectMathTypeLocalPrice(userCountry As String, katabanFirstHyphen As String) _
        As MathTypeInfoLocalPrice _
        Implements IDbAccessService.SelectMathTypeLocalPrice

        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  ISNULL(list_price_rate1,0) AS list_price_rate1, ")
            sql.Append("         ISNULL(list_price_rate2,0) AS list_price_rate2, ")
            sql.Append("         ISNULL(math_TypeA,-1) AS mathType, ")
            sql.Append("         ISNULL(math_PosA,1) AS mathPosition ")
            sql.Append(" FROM    kh_country_rate_localprice_mst as tblRate ")
            sql.Append(" INNER JOIN kh_country_mst as tblCun ")
            sql.Append(" ON tblRate.country_cd = tblCun.country_cd ")
            sql.Append(" WHERE   tblRate.country_cd          = @userCountry ")
            sql.Append(" AND     tblRate.rate_search_key     = @katabanFirstHyphen ")
            sql.Append(" AND     tblRate.in_effective_date  <= @standardDate ")
            sql.Append(" AND     tblRate.out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of MathTypeInfoLocalPrice)(sql.ToString,
                                                             New With {
                                                                userCountry,
                                                                katabanFirstHyphen,
                                                                standardDate
                                                                }
                                                             ).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     購入価格計算方法を取得
    ''' </summary>
    ''' <param name="userCountry">ユーザー国コード</param>
    ''' <param name="katabanFirstHyphen">形番第一ハイフン</param>
    ''' <param name="shipPlace">出荷場所</param>
    ''' <returns></returns>
    Public Function SelectMathTypeFobPrice(userCountry As String,
                                           katabanFirstHyphen As String,
                                           shipPlace As String) _
        As MathTypeInfoFobPrice Implements IDbAccessService.SelectMathTypeFobPrice

        Using connection As New SqlConnection(My.Settings.khBaseDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  ISNULL(fob_rate,0) AS fob_rate, ")
            sql.Append("         ISNULL(math_Type,-1) AS mathType, ")
            sql.Append("         ISNULL(math_Pos,1) AS mathPosition, ")
            sql.Append("         currency_cd, ")
            sql.Append("         authorization_no ")
            sql.Append(" FROM    kh_country_rate_netprice_mst as tblRate ")
            sql.Append(" INNER JOIN kh_currency_trade_mst as tblCun ")
            sql.Append(" ON tblRate.exp_country_cd = tblCun.exp_country_cd ")
            sql.Append(" AND tblRate.imp_country_cd = tblCun.imp_country_cd ")
            sql.Append(" WHERE   tblRate.exp_country_cd      = @shipPlace ")
            sql.Append(" AND     tblRate.imp_country_cd      = @userCountry ")
            sql.Append(" AND     tblRate.rate_search_key     = @katabanFirstHyphen ")
            sql.Append(" AND     tblRate.in_effective_date  <= @standardDate ")
            sql.Append(" AND     tblRate.out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of MathTypeInfoFobPrice)(sql.ToString,
                                                           New With {
                                                              userCountry,
                                                              shipPlace,
                                                              katabanFirstHyphen,
                                                              standardDate}).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     販売数量単位の取得
    ''' </summary>
    ''' <param name="kataban">形番</param>
    ''' <param name="language">言語</param>
    ''' <returns></returns>
    Public Function SelectQuantityUnit(kataban As String, language As String) As QuantityUnitInfo _
        Implements IDbAccessService.SelectQuantityUnit
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  c.qty_unit_nm, ")
            sql.Append("         b.qty_unit_nm as default_unit_nm,")
            sql.Append("         d.sales_unit,")
            sql.Append("         d.sap_base_unit,")
            sql.Append("         d.quantity_per_sales_unit,")
            sql.Append("         d.order_lot")
            sql.Append(" FROM    kh_qty_unit a")
            sql.Append(" INNER JOIN  kh_qty_unit_nm_mst b")
            sql.Append(" ON      a.qty_unit_cd           = b.qty_unit_cd ")
            sql.Append(" AND     b.language_cd           = @defaultLanguage ")
            sql.Append(" AND     b.in_effective_date    <= @standardDate ")
            sql.Append(" AND     b.out_effective_date    > @standardDate ")
            sql.Append(" LEFT  JOIN  kh_qty_unit_nm_mst c")
            sql.Append(" ON      a.qty_unit_cd           = c.qty_unit_cd ")
            sql.Append(" AND     c.language_cd           = @language ")
            sql.Append(" AND     c.in_effective_date    <= @standardDate ")
            sql.Append(" AND     c.out_effective_date    > @standardDate ")
            sql.Append(" LEFT JOIN kh_qty_unit_mst AS d on a.qty_unit_cd = d.qty_unit_cd")
            sql.Append(" WHERE   a.kataban               = @kataban ")
            sql.Append(" AND     a.in_effective_date     <= @standardDate ")
            sql.Append(" AND     a.out_effective_date    > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Dim defaultLanguage = LanguageDiv.DefaultLang
            Return _
                connection.Query(Of QuantityUnitInfo)(sql.ToString,
                                                       New With {
                                                          kataban,
                                                          language,
                                                          defaultLanguage,
                                                          standardDate}).FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     EL該当チェック
    ''' </summary>
    ''' <param name="kataban">形番</param>
    ''' <param name="elFlag">EL区分</param>
    ''' <returns></returns>
    Public Function CheckEl(kataban As String, elFlag As String) As Boolean Implements IDbAccessService.CheckEl
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  COUNT(*) ")
            sql.Append(" FROM    kh_el_kataban_mst ")
            sql.Append(" WHERE   @Kataban Like kataban ")
            sql.Append(" AND     el_flg = @ElFlag ")

            connection.Open()

            Return connection.Query(Of Integer)(sql.ToString, New With {kataban, elFlag}).FirstOrDefault() > 0

        End Using
    End Function

    ''' <summary>
    '''     在庫情報の取得
    ''' </summary>
    ''' <param name="kataban">形番</param>
    ''' <param name="language">言語</param>
    ''' <param name="shipPlace">出荷場所</param>
    ''' <returns></returns>
    Public Function SelectStock(kataban As String, language As String, shipPlace As String) As StockInfo _
        Implements IDbAccessService.SelectStock
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  kh_stock.stock_place_cd, ")
            sql.Append("         kh_stock.stock_qty, ")
            sql.Append("         kh_stock.shipment_qty, ")
            sql.Append("         kh_stock_content.stock_content ")
            sql.Append(" FROM    kh_stock ")
            sql.Append(" INNER JOIN  kh_stock_content ")
            sql.Append(" ON      kh_stock.stock_cd                    = kh_stock_content.stock_cd ")
            sql.Append(" WHERE   kh_stock.kataban                     = @kataban ")
            sql.Append(" AND     kh_stock_content.language_cd         = @language ")
            sql.Append(" AND     kh_stock.stock_place_cd              = @shipPlace ")
            sql.Append(" AND     kh_stock.in_effective_date          <= @standardDate ")
            sql.Append(" AND     kh_stock.out_effective_date          > @standardDate ")
            sql.Append(" AND     kh_stock_content.in_effective_date  <= @standardDate ")
            sql.Append(" AND     kh_stock_content.out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of StockInfo)(sql.ToString, New With {kataban, language, shipPlace, standardDate}).
                    FirstOrDefault()

        End Using
    End Function

    ''' <summary>
    '''     標準ストローク情報を取得
    ''' </summary>
    ''' <param name="series">機種</param>
    ''' <param name="keyKataban">キー形番</param>
    ''' <param name="boreSize">口径</param>
    ''' <param name="country">生産国</param>
    ''' <returns></returns>
    Public Function SelectStroke(series As String, keyKataban As String, boreSize As Integer, country As String) _
        As List(Of StrokeInfo) Implements IDbAccessService.SelectStroke
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  a.min_stroke ")
            sql.Append("         a.max_stroke ")
            sql.Append("         a.stroke_unit ")
            sql.Append("         b.std_stroke ")
            sql.Append(" FROM    kh_stroke  a ")
            sql.Append(" INNER JOIN  kh_std_stroke_mst  b ")
            sql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sql.Append(" AND     a.bore_size           = b.bore_size ")
            sql.Append(" WHERE   a.series_kataban      = @series ")
            sql.Append(" AND     a.key_kataban         = @keyKataban ")
            sql.Append(" AND     a.bore_size           = @boreSize ")
            sql.Append(" AND     a.in_effective_date  <= @standardDate ")
            sql.Append(" AND     a.out_effective_date  > @standardDate ")
            sql.Append(" AND     b.in_effective_date  <= @standardDate ")
            sql.Append(" AND     b.out_effective_date  > @standardDate ")
            sql.Append(" AND     a.country_cd  = @country ")
            sql.Append(" ORDER BY  b.std_stroke DESC ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of StrokeInfo)(sql.ToString,
                                                 New With {series, keyKataban, boreSize, country, standardDate})

        End Using
    End Function

#End Region

#Region "価格キー関連"

    ''' <summary>
    '''     電圧情報
    ''' </summary>
    ''' <param name="series"></param>
    ''' <param name="keyKataban"></param>
    ''' <param name="portSize"></param>
    ''' <param name="coil"></param>
    ''' <param name="voltageDiv"></param>
    ''' <returns></returns>
    Public Function SelectVoltageInfo(series As String,
                                      keyKataban As String,
                                      portSize As String,
                                      coil As String,
                                      voltageDiv As String,
                                      voltage As Integer) As List(Of VoltageInfo) _
        Implements IDbAccessService.SelectVoltageInfo
        Using connection As New SqlConnection(My.Settings.khdbDevConnectionString)

            Dim sql As New StringBuilder

            sql.Append(" SELECT  a.max_voltage, ")
            sql.Append("         a.min_voltage, ")
            sql.Append("         b.std_voltage, ")
            sql.Append("         b.std_voltage_flag ")
            sql.Append(" FROM    kh_voltage  a ")
            sql.Append(" INNER JOIN  kh_std_voltage_mst  b ")
            sql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sql.Append(" AND     a.port_size           = b.port_size ")
            sql.Append(" AND     a.coil                = b.coil ")
            sql.Append(" AND     a.voltage_div         = b.voltage_div ")
            sql.Append(" WHERE   a.series_kataban      = @series ")
            sql.Append(" AND     a.key_kataban         = @keyKataban ")
            If Not String.IsNullOrEmpty(portSize) Then
                sql.Append(" AND     a.port_size           = @portSize ")
            End If
            If Not String.IsNullOrEmpty(coil) Then
                sql.Append(" AND     a.coil                = @coil ")
            End If
            sql.Append(" AND     a.voltage_div         = @voltageDiv ")
            sql.Append(" AND     a.in_effective_date  <= @standardDate ")
            sql.Append(" AND     a.out_effective_date  > @standardDate ")
            sql.Append(" AND     b.std_voltage         = @voltage ")
            sql.Append(" AND     b.in_effective_date  <= @standardDate ")
            sql.Append(" AND     b.out_effective_date  > @standardDate ")

            connection.Open()

            Dim standardDate = Now
            Return _
                connection.Query(Of VoltageInfo)(sql.ToString,
                                                  New _
                                                     With {series, keyKataban, portSize, coil, voltageDiv, voltage,
                                                     standardDate})

        End Using
    End Function

#End Region
End Class
