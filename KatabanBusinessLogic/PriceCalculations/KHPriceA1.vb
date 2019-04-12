'************************************************************************************
'*  ProgramID  ：KHPriceA1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ギャップスイッチ単体　ＧＰＳ２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceA1

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                Case "KBZ"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & selectedData.Symbols(1).Trim & "-ST-" & _
                                                               selectedData.Symbols(3).Trim & _
                                                               selectedData.Symbols(4).Trim & _
                                                               selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'マスターユニット
                    If selectedData.Symbols(7) = "1" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'ケーブル長
                    If selectedData.Symbols(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & selectedData.Symbols(8).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '軸オプション
                    If selectedData.Symbols(9).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "GPS3"
                    '基本価格キー
                    'RM1802010_オプション追加の為修正
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'コネクタケーブル価格キー
                    'オプション追加により修正  2016/01/16 修正 RM1701010
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'ブラケットケーブル価格キー
                    'オプション追加により修正  2016/01/16 修正 RM1701010
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "MGPS3"
                    '基本価格キー
                    'RM1802010_オプション追加の為修正
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1).Trim & _
                                                                selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'コネクタケーブル価格キー
                    'オプション追加により修正  2016/01/16 修正 RM1701010
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(1).Trim

                    'ブラケットケーブル価格キー
                    'オプション追加により修正  2016/01/16 修正 RM1701010
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 2

                Case "UGPS3"
                    '基本価格キー
                    'RM1802010_オプション追加の為修正
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1).Trim & _
                                                                selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1



                    'コネクタケーブル価格キー
                    'オプション追加により修正  2016/01/16 修正 RM1701010
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(1).Trim


                Case "GPS2"

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '配線オプション加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アタッチメント加算価格キー
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                'OPを価格キーに付与 2017/02/02 追加 松原
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '圧力計加算価格キー
                    If selectedData.Symbols(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case Else

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '配線オプション加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アタッチメント加算価格キー
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '圧力計加算価格キー
                    If selectedData.Symbols(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
