'************************************************************************************
'*  ProgramID  ：KHPriceB9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/30   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：クーラントバルブ
'*             ：ＣＶ２（高圧・エアオペレイト形）
'*             ：ＣＶ２Ｅ（低損圧形・電磁弁搭載形）
'*             ：ＣＶＳ２（高圧・電磁弁搭載形）
'*             ：ＣＶＳ２Ｅ（低損圧形・電磁弁搭載形）
'*             ：ＣＶ３１（中・高圧・エアオペレイト形）
'*             ：ＣＶＳ３１（中・高圧・電磁弁搭載形）
'*             ：ＣＶ３Ｅ（低損圧形・エアオペレイト形）
'*             ：ＣＶＳ３Ｅ（低損圧形・電磁弁搭載形）
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceB9

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String = Divisions.VoltageDiv.Standard
        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "CV2", "CVS2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "CV2E", "CVS2E", "CV31", "CVS31", "CV3E", "CVS3E"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "CVS2"
                    Select Case selectedData.Symbols(5).Trim
                        Case "2H", "3T", "3R"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "CVS2E", "CVS31", "CVS3E"
                    Select Case selectedData.Symbols(4).Trim
                        Case "2H", "3T", "3R"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            'その他オプション加算価格
            Select Case selectedData.Series.series_kataban.Trim
                Case "CV2", "CVS2", "CV2E", "CVS2E", "CV31", "CVS31", "CVS3E"
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "CV31"
                            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                        Case "CV2", "CVS31", "CV2E", "CVS2E", "CVS3E"
                            strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                        Case "CVS2"
                            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                    End Select

                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case selectedData.Series.series_kataban.Trim
                            Case ""
                            Case "CV2", "CVS2"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select

            'スイッチ加算価格
            Select Case selectedData.Series.series_kataban.Trim
                Case "CV2"
                    'スイッチ個数判定
                    If selectedData.Symbols(6).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        If selectedData.Symbols(6).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   selectedData.Symbols(7).Trim
                        If selectedData.Symbols(6).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                        End If
                    End If

                    'リード線長さ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(8).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                Case "CVS2"
                    'スイッチ個数判定
                    If selectedData.Symbols(7).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                        If selectedData.Symbols(7).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim & _
                                                                   selectedData.Symbols(8).Trim
                        If selectedData.Symbols(7).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        End If
                    End If

                    'リード線長さ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(9).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                Case "CV31"
                    'スイッチ個数判定
                    If selectedData.Symbols(5).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        If selectedData.Symbols(5).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim & _
                                                                   selectedData.Symbols(6).Trim
                        If selectedData.Symbols(5).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                        End If
                    End If

                    'リード線長さ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                Case "CVS31"
                    'スイッチ個数判定
                    If selectedData.Symbols(6).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        If selectedData.Symbols(6).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   selectedData.Symbols(7).Trim
                        If selectedData.Symbols(6).Trim = "D" Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                        End If
                    End If

                    'リード線長さ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(8).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
            End Select

            'その他電圧加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "CVS2", "CVS2E", "CVS3", "CVS31", "CVS3E"
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "CVS2"
                            strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                           selectedData.Symbols(10).Trim)
                        Case "CVS2E", "CVS3E"
                            strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                           selectedData.Symbols(6).Trim)
                        Case "CVS3", "CVS31"
                            strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                           selectedData.Symbols(9).Trim)
                    End Select

                    Select Case strStdVoltageFlag
                        Case Divisions.VoltageDiv.Standard
                        Case Divisions.VoltageDiv.Options
                            Select Case selectedData.Series.series_kataban.Trim
                                Case "CVS2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & "OPT"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "CVS31", "CVS3E"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OPT"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Case Divisions.VoltageDiv.Other
                            Select Case selectedData.Series.series_kataban.Trim
                                Case "CVS2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "A" & MyControlChars.Hyphen & "OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "CVS2E", "CVS3", "CVS31", "CVS3E"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
