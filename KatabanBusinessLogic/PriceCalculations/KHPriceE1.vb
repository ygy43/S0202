'************************************************************************************
'*  ProgramID  ：KHPriceE1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：パルスジェットバルブ操作用ボックス形多連式電磁弁
'*             ：ＰＤ２／ＰＤＶ２／ＰＤ３／ＰＤＶ３／ＰＪＶＢ／ＰＤＶＥ４／ＮＰ１３／ＮＰ１４
'*
'*  更新履歴   ：                       更新日：2007/05/18   更新者：NII A.Takahashi
'*               ・コイルオプション追加により、2CS/2ES/2HS/3RSのコイルについても加算価格キーを生成
'*  ・受付No：RM1001045  二次電池対応  ＮＰ１３／ＮＰ１４／ＮＶＰ１１
'*                                      更新日：2010/02/24   更新者：Y.Miura
'*  ・受付No：RM1004012  二次電池対応  ＮＰ１３／ＮＰ１４／ＮＶＰ１１／ＮＡＰ１１
'*                                      更新日：2010/04/22   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceE1

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                Case "PD2", "PDV2"
                    '基本価格キー
                    If selectedData.Symbols(2).Trim = "F" Then
                        If selectedData.Symbols(3).Trim = "4A" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'コイルオプション加算価格キー
                    Select Case selectedData.Symbols(3).Trim
                        '2011/03/07 MOD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                        Case "2E", "2G", "2H", "3A", "3K", "3H", "3M", "3N"
                            'Case "2E", "2G", "2H", "3A", "3K", "3H"
                            '2011/03/07 MOD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'その他オプション加算価格キー
                    Select Case selectedData.Symbols(4).Trim
                        Case "S"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '電圧加算価格キー
                    If selectedData.Series.series_kataban.Trim = "PDV2" Then
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(5).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                            Case Divisions.VoltageDiv.Other
                                If Left(selectedData.Symbols(5).Trim, 2) = Divisions.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-AC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-DC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    End If
                Case "PD3", "PDV3"
                    '基本価格キー
                    Select Case selectedData.Symbols(2).Trim
                        Case "N", "F"
                            If selectedData.Symbols(3).Trim = "4A" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'コイルオプション加算価格キー
                    Select Case selectedData.Symbols(3).Trim
                        Case "2E", "2G", "2H", "2CG", "2CH", "3A", "3T", "3R", "2CS", "2ES", "2HS", "3RS"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '電圧加算価格キー
                    If selectedData.Series.series_kataban.Trim = "PDV3" Then
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(4).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                            Case Divisions.VoltageDiv.Other
                                If Left(selectedData.Symbols(4).Trim, 2) = Divisions.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-AC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-DC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    End If
                Case "PJVB"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電線管ねじポート２箇所加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '電圧加算価格キー
                    If selectedData.Symbols(2).Trim <> "5" Then
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(5).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                            Case Divisions.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "PDVE4"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電圧加算価格キー
                    strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                   selectedData.Symbols(3).Trim)
                    Select Case strStdVoltageFlag
                        Case Divisions.VoltageDiv.Standard
                        Case Divisions.VoltageDiv.Options
                        Case Divisions.VoltageDiv.Other
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "NP13", "NP14", "NVP11"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'コイルオプション加算価格キー
                    Select Case selectedData.Symbols(3).Trim
                        Case "2G", "2H", "3T", "3R"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'その他オプション加算価格キー
                    Select Case selectedData.Symbols(4).Trim
                        Case "S"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '電圧加算価格キー
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "NP13", "NP14", "NVP11"
                            strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                           selectedData.Symbols(5).Trim)
                        Case Else
                            strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                           selectedData.Symbols(6).Trim)
                    End Select
                    Select Case strStdVoltageFlag
                        Case Divisions.VoltageDiv.Standard
                        Case Divisions.VoltageDiv.Options
                        Case Divisions.VoltageDiv.Other
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'RM1001045 2010/02/24 Y.Miura 二次電池対応
                    'RM1004012 2010/04/22 Y.Miura 二次電池対応（オプション位置）
                    'オプション２
                    Dim intP4Position As Integer
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "NP13", "NP14", "NVP11"
                            intP4Position = 6
                        Case Else
                            intP4Position = 5
                    End Select
                    If selectedData.Symbols(intP4Position).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                   selectedData.Symbols(intP4Position).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "NAP11"    'RM1004012 2010/04/22 Y.Miura 追加
                    '基本価格キー
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & "-1"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション 二次電池(P4)
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
