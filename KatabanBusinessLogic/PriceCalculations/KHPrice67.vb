'************************************************************************************
'*  ProgramID  ：KHPrice67
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：セレックスロータリー
'*             ：ＲＶ３Ｓ／ＤＡ　小形・角度可変形
'*             ：ＲＶ３Ｓ／ＤＨ　大形・低油圧形
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice67

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionC As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                Case "RV3SA", "RV3DA"
                    '基本価格キー
                    Select Case selectedData.Symbols(2).Trim
                        Case "0"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "HOPEANGLE" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select


                    'スイッチ加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "RV3S" & MyControlChars.Hyphen & "FR" & MyControlChars.Hyphen & selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "RV3SH", "RV3DH"
                    '選択オプション分解＆ショックキラー選択判定
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "C"
                                bolOptionC = True
                        End Select
                    Next

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        Select Case bolOptionC
                            Case True
                                Select Case selectedData.Symbols(6).Trim
                                    Case "D"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "C" & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "C" & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(4).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Case Else
                                Select Case selectedData.Symbols(6).Trim
                                    Case "D"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(4).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                        End Select

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & MyControlChars.Hyphen & selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                        End If
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "FA", "LS"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "T"
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
