'************************************************************************************
'*  ProgramID  ：KHPrice42
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＲＧ＊＊／ＰＣ＊Ｓ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice42

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            Select Case True
                Case selectedData.Series.key_kataban.Trim = "W"
                    If Left(selectedData.Symbols(7).Trim, 1) >= "0" And Left(selectedData.Symbols(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "W" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "FC" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "W" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "AL" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case selectedData.Series.key_kataban.Trim = "E"
                    If Left(selectedData.Symbols(7).Trim, 1) >= "0" And Left(selectedData.Symbols(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "E" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "FC" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "E" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "AL" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case selectedData.Series.key_kataban.Trim = "G"
                    If Left(selectedData.Symbols(7).Trim, 1) >= "0" And Left(selectedData.Symbols(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "G" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "FC" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(15).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "G" & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "AL" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim & _
                                                                   selectedData.Symbols(15).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case Else
                    If Left(selectedData.Symbols(7).Trim, 1) >= "0" And Left(selectedData.Symbols(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "FC"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "AL"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            '入出力仕様加算価格キー
            For intLoopCnt = 8 To 10
                If selectedData.Symbols(intLoopCnt).Trim <> "" And _
                   selectedData.Symbols(intLoopCnt).Trim <> "N" And _
                   (intLoopCnt <> 9 Or selectedData.Symbols(8).Trim <> "K" Or selectedData.Symbols(9).Trim <> "K") Then
                    Select Case True
                        Case intLoopCnt = 10
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                       Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "O" & _
                                                                       selectedData.Symbols(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            If selectedData.Symbols(intLoopCnt).Trim = "H" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "I" & _
                                                                           selectedData.Symbols(intLoopCnt).Trim & _
                                                                           selectedData.Symbols(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "I" & _
                                                                           selectedData.Symbols(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                End If
            Next

            'オプション加算価格キー
            Select Case selectedData.Series.key_kataban.Trim
                Case "W", "E"
                    intLoopCnt = 15
                Case "G", "H"
                    intLoopCnt = 16
                Case Else
                    intLoopCnt = 11
            End Select

            If selectedData.Symbols(intLoopCnt).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                           selectedData.Symbols(1).Trim

                Select Case selectedData.Symbols(10).Trim
                    Case "F", "A", "S", "B"
                        Select Case Mid(selectedData.Series.series_kataban.Trim, 4, 1)
                            Case "S", "L"
                                Select Case selectedData.Series.series_kataban.Trim
                                    Case "RGCS", "RGIL", "RGIS", "RGOL", "RGOS", "PCIS", "PCOS"
                                        If selectedData.Symbols(12).Trim <> "" Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "TSF" & _
                                                                                       selectedData.Symbols(intLoopCnt).Trim
                                        Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "TSF" & MyControlChars.Hyphen & "NO"
                                        End If
                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "TSF" & _
                                                                                   selectedData.Symbols(intLoopCnt).Trim
                                End Select
                            Case "T"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "TST" & _
                                                                           selectedData.Symbols(intLoopCnt).Trim
                        End Select
                    Case "X", "C", "Y", "D"
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "TGX" & _
                                                                   selectedData.Symbols(intLoopCnt).Trim
                End Select

                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
