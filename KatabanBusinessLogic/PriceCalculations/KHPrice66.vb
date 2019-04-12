'************************************************************************************
'*  ProgramID  ：KHPrice66
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＲＶ３Ｓ／Ｄ　小形・標準形
'*  　　       ：ＲＶ３Ｓ／Ｄ　小形・バルブ付
'*  　　       ：ＲＶ３Ｓ／Ｄ　大形・標準形
'*  　　       ：ＲＶ３Ｓ／Ｄ　大形・バルブ付
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice66

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.key_kataban.Trim
                Case "1"
                    If selectedData.Symbols(6).Trim = "" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                   "SR"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "2"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "RV3S-SR-" & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "3"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        If selectedData.Symbols(9).IndexOf("C") >= 0 Then
                            If selectedData.Symbols(8).Trim = "D" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           "C" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           "C" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Else
                            If selectedData.Symbols(8).Trim = "D" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                    'リード線長さ
                    If selectedData.Symbols(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                    End If

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "FA", "LS"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "4"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        If selectedData.Symbols(9).IndexOf("C") >= 0 Then
                            If selectedData.Symbols(8).Trim = "D" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           "C" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           "C" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Else
                            If selectedData.Symbols(8).Trim = "D" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                    'リード線長さ
                    If selectedData.Symbols(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                    End If

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "FA", "LS"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
