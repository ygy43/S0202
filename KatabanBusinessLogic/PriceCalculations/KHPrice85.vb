'************************************************************************************
'*  ProgramID  ：KHPrice23
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ペンシルシリンダ複動形　ＳＣＰＤ２／ＳＣＰＤ２－Ｌ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice85

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionP4 As Boolean = False      'RM1001045 2010/02/23 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
            If selectedData.Symbols.Count > 9 Then
                strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "P4", "P40"
                            bolOptionP4 = True
                    End Select
                Next
            End If


            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(3).Trim), _
                                                  CInt(selectedData.Symbols(4).Trim))

            Select Case selectedData.Series.series_kataban
                Case "SCPG2", "SCPG2-L", "SCPG2-X", "SCPG2-XL", "SCPG2-Y", "SCPG2-YL"
                    Select Case selectedData.Series.series_kataban
                        Case "SCPG2-X", "SCPG2-XL"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-X" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "SCPG2-Y", "SCPG2-YL"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-Y" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'シリーズオプション加算価格キー(2)
                    If Mid(selectedData.Series.series_kataban, 7, 1) = "L" Or _
                       Mid(selectedData.Series.series_kataban, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        End If

                    End If

                    'オプション・付属品加算価格キー
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                Case Else
                    'バリエーション(微速)加算価格キー
                    Select Case selectedData.Symbols(1).Trim
                        Case "F"
                            Select Case selectedData.Symbols(3).Trim
                                Case "6"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 30
                                            'ストローク10～30
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 31 And _
                                             CInt(selectedData.Symbols(4).Trim) <= 60
                                            'ストローク31～60
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR31" & MyControlChars.Hyphen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 61
                                            'ストローク61～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR61" & MyControlChars.Hyphen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "10"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 45
                                            'ストローク10～45
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "45"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 46 And _
                                             CInt(selectedData.Symbols(4).Trim) <= 100
                                            'ストローク46～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR46" & MyControlChars.Hyphen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 101
                                            'ストローク101～200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR101" & MyControlChars.Hyphen & "200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 45
                                            'ストローク10～45
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "45"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 46 And _
                                             CInt(selectedData.Symbols(4).Trim) <= 100
                                            'ストローク46～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR46" & MyControlChars.Hyphen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 101
                                            'ストローク101～260
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR101" & MyControlChars.Hyphen & "260"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション加算価格キー(2)
                    If Mid(selectedData.Series.series_kataban, 7, 1) = "L" Or _
                       Mid(selectedData.Series.series_kataban, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    If selectedData.Symbols(2).Trim = "CB" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        End If

                        'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                        'P4加算
                        If bolOptionP4 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        End If
                    End If

                    'オプション・付属品加算価格キー
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    'クリーン仕様加算価格キー
                    If selectedData.Symbols(11).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & MyControlChars.Hyphen & _
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
