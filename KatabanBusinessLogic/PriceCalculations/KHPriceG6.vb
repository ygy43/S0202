'************************************************************************************
'*  ProgramID  ：KHPriceG6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：薬液用エアオペレイトバルブ
'*             ：ＡＭＤ３＊２
'*             ：ＡＭＤ４＊２
'*             ：ＡＭＤ５＊２
'*             ：ＡＭＤ０＊２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceG6

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "AMD0"
                    If selectedData.Symbols(7).Trim = "Y" Then
                        Select Case selectedData.Symbols(5).Trim
                            Case "0", "6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "0" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "1" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "2", "7"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "2" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "3"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "3" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "8"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "8" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        'RM1310067 2013/10/23
                        If selectedData.Symbols(4).Trim = "" Then
                            Select Case selectedData.Symbols(5).Trim
                                Case "0", "6"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "4-0"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "4-1"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "2", "7"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "4-2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "3"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "4-3"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "8"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "4-8"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            Select Case selectedData.Symbols(5).Trim
                                Case "0", "6"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "0"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "1"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "2", "7"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "3"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "3"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "8"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "8"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                    End If
                Case "AMD3"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case selectedData.Series.key_kataban.Trim
                        Case "1"
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                "8" & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                            If selectedData.Symbols(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(7).Trim
                            End If
                        Case "2"
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                "10" & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                            If selectedData.Symbols(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(7).Trim
                            End If
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(4).Trim & MyControlChars.Hyphen
                            If selectedData.Symbols(5).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(5).Trim
                            End If
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AMD4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case selectedData.Series.key_kataban.Trim
                        Case ""
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                "16" & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                            If selectedData.Symbols(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(7).Trim
                            End If
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(4).Trim & MyControlChars.Hyphen
                            If selectedData.Symbols(5).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(5).Trim
                            End If
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AMD5"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case selectedData.Series.key_kataban.Trim
                        Case ""
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                "20" & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5).Trim

                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '流体加算価格キー
            If selectedData.Symbols(7).Trim = "P" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*" & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'AMD**2シリーズR,X追加 2008/5/2
            '↓RM1310067 2013/10/23
            If (Left(selectedData.Series.series_kataban.Trim, 4) = "AMD0" And selectedData.Series.key_kataban.Trim = "1") Or _
               (Left(selectedData.Series.series_kataban.Trim, 4) = "AMD3" And (selectedData.Series.key_kataban.Trim = "1" Or selectedData.Series.key_kataban.Trim = "2")) Or _
               (Left(selectedData.Series.series_kataban.Trim, 4) = "AMD4" And selectedData.Series.key_kataban.Trim = "") Or _
               (Left(selectedData.Series.series_kataban.Trim, 4) = "AMD5" And selectedData.Series.key_kataban.Trim = "") Then
                If selectedData.Symbols(8).Trim = "R" Then

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(selectedData.Series.series_kataban.Trim, 4) & _
                    "*" & selectedData.Symbols(2).Trim & _
                    "-" & selectedData.Symbols(8).Trim

                    decOpAmount(UBound(decOpAmount)) = 1

                ElseIf selectedData.Symbols(8).Trim = "X" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If Len(selectedData.Symbols(7).Trim) <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(selectedData.Series.series_kataban.Trim, 4) & _
                        "*" & selectedData.Symbols(2).Trim & _
                        "-" & selectedData.Symbols(7).Trim & _
                        selectedData.Symbols(8).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(selectedData.Series.series_kataban.Trim, 4) & _
                        "*" & selectedData.Symbols(2).Trim & _
                        "- " & selectedData.Symbols(8).Trim
                    End If

                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf selectedData.Symbols(8).Trim = "R,X" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(selectedData.Series.series_kataban.Trim, 4) & _
                    "*" & selectedData.Symbols(2).Trim & "-" & "R"

                    decOpAmount(UBound(decOpAmount)) = 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If Len(selectedData.Symbols(7).Trim) <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(selectedData.Series.series_kataban.Trim, 4) & _
                        "*" & selectedData.Symbols(2).Trim & _
                        "-" & selectedData.Symbols(7).Trim & "X"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(selectedData.Series.series_kataban.Trim, 4) & _
                        "*" & selectedData.Symbols(2).Trim & _
                        "- " & "X"
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
