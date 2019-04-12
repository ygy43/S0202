'************************************************************************************
'*  ProgramID  ：KHPriceA0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパードライヤユニット　　　　　　　　ＳＵ
'*             ：スーパードライヤユニット（Ｄシリーズ）　ＳＵ***Ｄ
'*             ：スーパードライヤユニット（Ｅシリーズ）　ＳＵ***Ｅ
'*             ：スーパードライヤ　　　　　　　　　　　　ＳＤ
'*             ：スーパードライヤ（Ｄシリーズ）　　　　　ＳＤ***Ｄ
'*             ：スーパードライヤ（Ｅシリーズ）　　　　　ＳＤ***Ｅ
'*             ：スーパードライヤ・モジュラーシリーズ　　ＳＤＭ
'*
'* RM1003086 　：白色シリーズ対応 2010/03/26 Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceA0

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
                Case "HD"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim 
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Next

                Case Else

                    '基本価格キー
                    Select Case Mid(selectedData.Symbols(1).Trim, 4, 1)
                        Case "D", "E"
                            Select Case Left(selectedData.Series.series_kataban.Trim, 2)
                                Case "SD"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "SU"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    'RM1003086 2010/03/26 Y.Miura 白色シリーズ追加 
                                    If selectedData.Symbols(3).Trim = "W" Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                  selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                  selectedData.Symbols(2).Trim & "-W-" & _
                                                                                  selectedData.Symbols(4).Trim
                                    Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(3).Trim
                                    End If
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Case Else
                            Select Case Left(selectedData.Series.series_kataban.Trim, 2)
                                Case "SD"
                                    If Mid(selectedData.Series.series_kataban.Trim, 3, 1) = "M" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & _
                                                                                   selectedData.Symbols(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case "SU"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim & _
                                                                               selectedData.Symbols(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                    End Select

                    '2011/10/27 ADD RM1110032(11月VerUP:二次電池) START--->
                    Dim isDE As Boolean = False
                    '2011/10/27 ADD RM1110032(11月VerUP:二次電池) <---END

                    'オプション加算価格キー
                    Select Case Left(selectedData.Series.series_kataban.Trim, 2)
                        Case "SD"
                            Select Case True
                                Case Mid(selectedData.Series.series_kataban.Trim, 3, 1) <> "M" And _
                                     Mid(selectedData.Symbols(1).Trim, 4, 1) <> "D" And _
                                     Mid(selectedData.Symbols(1).Trim, 4, 1) <> "E"
                                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                                    For intLoopCnt = 0 To strOpArray.Length - 1
                                        Select Case strOpArray(intLoopCnt).Trim
                                            Case ""
                                            Case Else
                                                Select Case Left(selectedData.Symbols(1).Trim, 1)
                                                    Case "3"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "3000" & MyControlChars.Hyphen & _
                                                                                                   strOpArray(intLoopCnt).Trim
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Case "4"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "4000" & MyControlChars.Hyphen & _
                                                                                                   strOpArray(intLoopCnt).Trim
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                End Select
                                        End Select
                                    Next
                                Case Mid(selectedData.Symbols(1).Trim, 4, 1) = "D" Or _
                                     Mid(selectedData.Symbols(1).Trim, 4, 1) = "E"
                                    'RM1003086 2010/03/26 Y.Miura 白色シリーズ追加 
                                    If selectedData.Symbols(3) = "W" Then
                                        strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                                    Else
                                        strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                                    End If
                                    For intLoopCnt = 0 To strOpArray.Length - 1
                                        Select Case strOpArray(intLoopCnt).Trim
                                            Case ""
                                            Case Else
                                                Select Case Left(selectedData.Symbols(1).Trim, 1)
                                                    Case "3"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        'RM1003086 2010/03/26 Y.Miura 白色シリーズ追加 
                                                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "300" & MyControlChars.Hyphen & _
                                                        '                                           strOpArray(intLoopCnt).Trim
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "300" & "-OP-" & _
                                                                                                   strOpArray(intLoopCnt).Trim
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Case "4"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        'RM1003086 2010/03/26 Y.Miura 白色シリーズ追加 
                                                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "400" & MyControlChars.Hyphen & _
                                                        '                                           strOpArray(intLoopCnt).Trim
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "400" & "-OP-" & _
                                                                                                   strOpArray(intLoopCnt).Trim
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                End Select
                                        End Select
                                    Next
                                    '2011/10/27 ADD RM1110032(11月VerUP:二次電池) START--->
                                    isDE = True
                                    '2011/10/27 ADD RM1110032(11月VerUP:二次電池) <---END
                                Case Else
                                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                                    For intLoopCnt = 0 To strOpArray.Length - 1
                                        Select Case strOpArray(intLoopCnt).Trim
                                            Case ""
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                                           strOpArray(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = 1
                                        End Select
                                    Next
                            End Select
                            '2011/10/27 ADD RM1110032(11月VerUP:二次電池) START--->
                            Select Case selectedData.Series.key_kataban.Trim
                                Case "2", "X", "Z"
                                    Dim strPricePart As String = ""
                                    'SD3000,SD4000シリーズ
                                    If isDE Then
                                        strPricePart = "00-OP"
                                    Else
                                        strPricePart = "000"
                                    End If

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                            Left(selectedData.Symbols(1), 1) & strPricePart & MyControlChars.Hyphen & _
                                                                            selectedData.Symbols(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'オプション判定
                                    If selectedData.Symbols(6).Trim = "X1" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                Left(selectedData.Symbols(1), 1) & strPricePart & MyControlChars.Hyphen & _
                                                                                selectedData.Symbols(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1

                                    End If

                                    '二次電池加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                            Left(selectedData.Symbols(1), 1) & strPricePart & MyControlChars.Hyphen & _
                                                                            selectedData.Symbols(7).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                            End Select
                            '2011/10/27 ADD RM1110032(11月VerUP:二次電池) <---END
                        Case "SU"
                            'RM1003086 2010/03/26 Y.Miura 白色シリーズ追加 
                            Select Case selectedData.Series.key_kataban.Trim
                                '2011/10/27 ADD RM1110032(11月VerUP:二次電池) START--->
                                Case "F", "H", "J", "G", "I", "K"
                                    'Case "W", "X", "Y"
                                    '2011/10/27 ADD RM1110032(11月VerUP:二次電池) <---END
                                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                                Case Else
                                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                            End Select
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "E"
                                        Select Case Mid(selectedData.Symbols(1).Trim, 4, 1)
                                            Case "D", "E"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                           Left(selectedData.Symbols(1).Trim, 1) & "00" & _
                                                                                           Mid(selectedData.Symbols(1).Trim, 4, 1) & "-OP-" & _
                                                                                           strOpArray(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                           Left(selectedData.Symbols(1).Trim, 1) & "000-OP-" & _
                                                                                           strOpArray(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = 1
                                        End Select
                                End Select
                            Next
                            '2011/10/27 ADD RM1110032(11月VerUP:二次電池) START--->
                            Select Case selectedData.Series.key_kataban.Trim
                                Case "G", "I", "K"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'シリーズ判定
                                    Select Case Mid(selectedData.Symbols(1), 4, 1)
                                        Case "D", "E"
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                       Left(selectedData.Symbols(1).Trim, 1) & "00" & _
                                                                                       Mid(selectedData.Symbols(1).Trim, 4, 1) & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(7).Trim
                                        Case Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                       Left(selectedData.Symbols(1).Trim, 1) & "000" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(7).Trim
                                    End Select
                            End Select
                            '2011/10/27 ADD RM1110032(11月VerUP:二次電池) <---END
                    End Select
            End Select
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
