'************************************************************************************
'*  ProgramID  ：KHPrice23
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ペンシルシリンダ　ＳＣＰ＊２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice23

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim intIndex As Integer
        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(3).Trim))

            Select Case selectedData.Series.series_kataban
                Case "SCPG2-D", "SCPG2-DL", "SCPG2-M", "SCPG2-ML", "SCPG2-T"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション(2)加算価格キー
                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or _
                       Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & MyControlChars.Hyphen & "L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    Select Case True
                        Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "D"
                            intIndex = 3
                        Case Else
                            intIndex = 4
                    End Select

                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                        intIndex = intIndex + 1
                        If selectedData.Symbols(intIndex).Trim <> "" Then
                            'スイッチ加算価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(intIndex).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)

                            If selectedData.Symbols(intIndex + 1).Trim <> "" Then
                                'リード線長さ加算価格キー
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(intIndex + 1).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)
                            End If
                        End If
                        intIndex = intIndex + 2
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 1), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '付属品"I","Y"選択チェック
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                        End Select
                    Next

                    '付属品加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim

                                If Mid(selectedData.Series.series_kataban.Trim, 6, 2) = "-D" Then
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "I"
                                            Select Case True
                                                Case bolOptionY = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case "Y"
                                            Select Case True
                                                Case bolOptionI = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Next

                Case "SCPG2-XML", "SCPG2-XM"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & "-M" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション(2)加算価格キー
                    If Mid(selectedData.Series.series_kataban.Trim, 9, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & MyControlChars.Hyphen & "L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    If Left(selectedData.Symbols(1).Trim, 2) = "CB" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    intIndex = 4

                    If Mid(selectedData.Series.series_kataban.Trim, 9, 1) = "L" Then
                        intIndex = intIndex + 1
                        If selectedData.Symbols(intIndex).Trim <> "" Then
                            'スイッチ加算価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(intIndex).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)

                            If selectedData.Symbols(intIndex + 1).Trim <> "" Then
                                'リード線長さ加算価格キー
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(intIndex + 1).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)
                            End If
                        End If
                        intIndex = intIndex + 2
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 1), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '付属品"I","Y"選択チェック
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                        End Select
                    Next

                    '付属品加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim

                                If Mid(selectedData.Series.series_kataban.Trim, 6, 2) = "-D" Then
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "I"
                                            Select Case True
                                                Case bolOptionY = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case "Y"
                                            Select Case True
                                                Case bolOptionI = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Next

                Case Else
                    '基本価格キー
                    Select Case True
                        Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "D" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "K" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "M" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "O" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "T" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "V" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "Z"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Mid(selectedData.Series.series_kataban.Trim, 5, 1) = ""
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'シリーズオプション(1)加算価格キー
                    If Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "T" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & MyControlChars.Hyphen & "T" & MyControlChars.Hyphen & selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'シリーズオプション(2)加算価格キー
                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or _
                       Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & MyControlChars.Hyphen & "L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    If Left(selectedData.Symbols(1).Trim, 2) = "CB" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    Select Case True
                        Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "D" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "Z" Or _
                             Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "V"
                            intIndex = 3
                        Case Else
                            intIndex = 4
                    End Select

                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "V" Then
                        intIndex = intIndex + 1
                        If selectedData.Symbols(intIndex).Trim <> "" Then
                            '電圧特注加算価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & MyControlChars.Hyphen & selectedData.Symbols(intIndex).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If

                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                        intIndex = intIndex + 1
                        If selectedData.Symbols(intIndex).Trim <> "" Then
                            'スイッチ加算価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(intIndex).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)

                            If selectedData.Symbols(intIndex + 1).Trim <> "" Then
                                'リード線長さ加算価格キー
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(intIndex + 1).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)
                            End If
                        End If
                        intIndex = intIndex + 2
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 1), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '付属品"I","Y"選択チェック
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                        End Select
                    Next

                    '付属品加算価格キー
                    strOpArray = Split(selectedData.Symbols(intIndex + 2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim

                                If Mid(selectedData.Series.series_kataban.Trim, 6, 2) = "-D" Then
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "I"
                                            Select Case True
                                                Case bolOptionY = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case "Y"
                                            Select Case True
                                                Case bolOptionI = True
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case Else
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Next
            End Select
        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
