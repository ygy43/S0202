'************************************************************************************
'*  ProgramID  ：KHPriceB3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：2011/11/11   更新者：Y.Tachi
'*
'*  概要       ：薬液用エアオペレイトバルブ　ＳＵＳボディ
'*             ：ＡＭＤ３＊２／ＡＭＤ４＊２／ＡＭＤ５＊２
'*             ：ＡＭＤ３＊３／ＡＭＤ４＊３
'*             ：ＡＭＤＺ＊３
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceB3

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                'RM1307003 2013/07/04追加
                Case "LMD0"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アクチュエータオプション
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '金属コーディングオプション
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '継手セットの有無
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '取付板
                    If selectedData.Symbols(7).Trim <> "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "LMD3", "LMD4", "LMD5"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アクチュエータオプション
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '金属コーディングオプション
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '継手セットの有無
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "AMDZ"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & _
                                                               selectedData.Symbols(5).Trim & _
                                                               selectedData.Symbols(6).Trim & _
                                                               selectedData.Symbols(7).Trim

                    decOpAmount(UBound(decOpAmount)) = 1
                    'Case "AMD3", "AMD4", "AMD5"
                    '    '基本価格キー
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*3-" & MyControlChars.Hyphen & _
                    '                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                    '                                               selectedData.Symbols(4).Trim & _
                    '                                               selectedData.Symbols(5).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1706067 "GAMD0"追加　2017/07/06
                    'Case "GAMD3", "GAMD4", "GAMD5"
                Case "GAMD0", "GAMD3", "GAMD4", "GAMD5"
                    If selectedData.Symbols(2).Trim = "3R" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "*3R-" & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'アクチュエータオプション
                        If selectedData.Symbols(4).Trim <> "00" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "*3R-" & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(5).Trim
                        End If
                    Else
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "*3-" & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'アクチュエータオプション
                        If selectedData.Symbols(4).Trim <> "00" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "*3-" & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(5).Trim
                        End If
                    End If
                Case "GAMDZ"

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "*3R-" & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1


                Case Else
                    If selectedData.Symbols(2).Trim = "3" Then
                        'AMD3*3,AMD4*3シリーズ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*3" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    ElseIf selectedData.Symbols(2).Trim = "3R" Then
                        'AMD3*3R,AMD4*3Rシリーズ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*3R" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        'AMD3*2,AMD4*2,AMD5*2シリーズ
                        '価格キー設定
                        If selectedData.Symbols(5).Trim = "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*2" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*2" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
