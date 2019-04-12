'************************************************************************************
'*  ProgramID  ：KHPrice98
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＰＶ５形マニホールド　ＰＶ５（Ｇ）－６／８（Ｒ）
'*
'*  更新履歴   ：                       更新日：2007/05/17   更新者：NII A.Takahashi
'*               ・CEマーキング対応により、新ISOバルブ形番生成のときの"-ST"を削除
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice98

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Dim fullKataban = PriceManager.GetFullKataban(selectedData)

            '旧ISOバルブ(小牧分)
            If Left(fullKataban.Trim, 6) = "PV5-6-" Or _
               Left(fullKataban.Trim, 6) = "PV5-8-" Then
                '基本価格キー
                'その他電圧指定時の指定電圧部分を削除する
                If InStr(1, fullKataban.Trim, "-AC") <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, InStr(1, fullKataban.Trim, "-AC") - 1)
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = fullKataban.Trim
                End If
                '切削油対応部分を削除する
                If InStr(1, fullKataban.Trim, "-F1AW") <> 0 Then
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(1, fullKataban.Trim, "-F1AW") - 1)
                End If
                'その他電圧("-9")を"-1"へ変更する
                If InStr(1, fullKataban.Trim, "-9") <> 0 Then
                    strOpRefKataban(UBound(strOpRefKataban)) = Replace(strOpRefKataban(UBound(strOpRefKataban)), "-9", "-1")
                End If
                '不要なハイフンを削除する
                strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut(strOpRefKataban(UBound(strOpRefKataban)))
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1

                '切削油対応価格加算
                If InStr(1, fullKataban.Trim, "-F1AW") <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & MyControlChars.Hyphen & "F1AW"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                'その他電圧価格加算
                If InStr(1, fullKataban.Trim, "-9") <> 0 Then
                    If InStr(1, fullKataban.Trim, "-S") <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & "-S-OTH"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & "-D-OTH"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                '↓RM1303003 2013/03/04 Y.Tachi
                If Left(fullKataban.Trim, 4) = "PV5S" Then
                    If selectedData.Symbols(5).Trim = "" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-" & selectedData.Symbols(1).Trim & "-" & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Else
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-" & selectedData.Symbols(1).Trim & "-" & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    If selectedData.Symbols(4).Trim = "ML" Then
                        If selectedData.Symbols(2).Trim = "FG-S" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-FG-S-" & selectedData.Symbols(4).Trim

                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-FG-D-" & selectedData.Symbols(4).Trim

                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Else
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = fullKataban.Trim
                    '切削油対応部分を削除する
                    If InStr(1, strOpRefKataban(UBound(strOpRefKataban)), "A-") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(1, fullKataban.Trim, "A-") - 1) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(1, fullKataban.Trim, "A-") + 1, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If Right(strOpRefKataban(UBound(strOpRefKataban)), 1) = "A" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(strOpRefKataban(UBound(strOpRefKataban)), 1, Len(strOpRefKataban(UBound(strOpRefKataban))) - 1)
                    End If
                    '電圧部分を削除する
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-1") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-1")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-1") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-2") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-2")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-2") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-3") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-3")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-3") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-4") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-4")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-4") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-5") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-5")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-5") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    If InStr(6, strOpRefKataban(UBound(strOpRefKataban)), "-6") <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-6")) & Mid(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-6") + 2, Len(strOpRefKataban(UBound(strOpRefKataban))))
                    End If
                    ' CEマーキング部分を削除する
                    'If InStr(1, strOpRefKataban(UBound(strOpRefKataban)), "-ST") <> 0 Then
                    '    '2010/11/18 MOD RM1011020(12月VerUP:PV5シリーズ_不具合修正) START--->
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-ST") - 1)
                    '    'strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-ST") - 2)
                    '    '2010/11/18 MOD RM1011020(12月VerUP:PV5シリーズ_不具合修正) <---END
                    'End If

                    'RM1305XXX 2013/05/02
                    If InStr(1, strOpRefKataban(UBound(strOpRefKataban)), "-ST") <> 0 Then
                        If Left(fullKataban.Trim, 4) = "PV5G" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-ST") - 2)
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), InStr(6, fullKataban.Trim, "-ST") - 1)
                        End If
                    End If

                    ' 不要なハイフンを削除する
                    strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut(strOpRefKataban(UBound(strOpRefKataban)))
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1

                    ' 切削油対応価格加算
                    If InStr(1, fullKataban.Trim, "A-") <> 0 Or _
                       Right(fullKataban.Trim, 1) = "A" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & "-A"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    ' その他電圧価格加算
                    '2010/06/24 T.Fuji RM1005050(電圧追加：PV5Gシリーズ) --->
                    'If InStr(6, fullKataban.Trim, "-5") <> 0 Or _
                    '   InStr(6, fullKataban.Trim, "-6") <> 0 Then
                    If InStr(6, fullKataban.Trim, "-2") <> 0 Or _
                       InStr(6, fullKataban.Trim, "-5") <> 0 Or _
                       InStr(6, fullKataban.Trim, "-6") <> 0 Then
                        '2010/06/24 T.Fuji RM1005050(電圧追加：PV5Gシリーズ) <---
                        If InStr(1, fullKataban.Trim, "-S") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & "-S-OTH"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 3) & "-D-OTH"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
