'************************************************************************************
'*  ProgramID  ：KHPriceE5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*
'*  概要       ：落下防止付クランプシリンダ　ＵＣＡＣ／ＵＣＡＣ－Ｌ２／ＵＣＡＣ２／ＵＣＡＣ２－Ｌ２
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*  ・受付No：RM0811133  UCAC2新発売
'*                                      更新日：2009/07/28   更新者：Y.Miura
'*  ・受付No：RM1001018　スイッチT2YDUをC5扱いとする
'*                                      更新日：2010/01/18   更新者：Y.Miura
'*  ・受付No：RM1003086　タイロッド取付位置追加   2010/03/26 Y.Miura        
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceE5

    'Public Sub subPriceCalculation(selectedData As SelectedInfo, _
    '                               ByRef strOpRefKataban() As String, _
    '                               ByRef decOpAmount() As Decimal)
    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim bolC5Flag As Boolean    'RM0811133 2009/07/28 Y.Miura 追加

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)    'RM0811133 2009/07/28 Y.Miura 追加

            'RM0811133 2009/07/28 Y.Miura　↓↓
            'シリーズ形番の第1ハイフン前を取得
            Dim strHySeriesKataban As String = selectedData.Series.series_kataban.Trim
            If InStr(strHySeriesKataban, "-") > 0 Then
                strHySeriesKataban = strHySeriesKataban.Substring(0, InStr(strHySeriesKataban, "-") - 1)
            End If
            '要素位置の設定
            'UCAC2は3番目の要素『配管ねじ種類』が存在するのでstrOpSymbol(3)以降はプラス1する
            'RM1003086 2010/03/26 Y.Miura 
            'タイロッド取付位置追加に伴い、UCAC2の『付属品』以降はさらにプラス1する
            Dim intOpt As Integer
            Dim intOpt2 As Integer
            Select Case strHySeriesKataban
                Case "UCAC"
                    intOpt = 0
                    intOpt2 = 0
                Case "UCAC2"
                    intOpt = 1
                    intOpt2 = 1
            End Select

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)

            'ストローク取得
            'intStroke = KHKataban.fncGetStrokeSize(selectedData.Series.series_kataban, _
            '                                      selectedData.Series.key_kataban, _
            '                                      CInt(selectedData.Symbols(2).Trim), _
            '                                      CInt(selectedData.Symbols(4).Trim))
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                     CInt(selectedData.Symbols(2).Trim), _
                                                     CInt(selectedData.Symbols(4 + intOpt).Trim))
            'RM0811133 2009/07/28 Y.Miura　↑↑

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            'RM0811133 2009/07/28 Y.Miura
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            'クレビス幅加算価格キー
            Select Case selectedData.Symbols(1).Trim
                Case "AL", "BL", "C", "CL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'RM0811133 2009/07/28 Y.Miura
                    'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                    '                                           selectedData.Symbols(1).Trim
                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0811133 2009/07/28 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            'RM0811133 2009/07/28 Y.Miura
            'If selectedData.Symbols(7).Trim <> "" Then
            If selectedData.Symbols(7 + intOpt).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                'RM0811133 2009/07/28 Y.Miura ↓↓
                'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(7).Trim
                Select Case strHySeriesKataban
                    Case "UCAC"
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7 + intOpt).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                    Case "UCAC2"
                        '↓RM1309001 2013/09/02 追加
                        If selectedData.Symbols(11).Trim = "Z" Or selectedData.Symbols(12).Trim = "Z" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & _
                                                                       selectedData.Symbols(7 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                        Else

                            strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
                'RM0811133 2009/07/28 Y.Miura ↑↑

                'リード線の長さ加算価格キー
                'RM0811133 2009/07/28 Y.Miura
                'If selectedData.Symbols(8).Trim <> "" Then
                If selectedData.Symbols(8 + intOpt).Trim <> "" Then

                    Select Case selectedData.Series.series_kataban.Trim
                        'RM0811133 2009/07/28 Y.Miura ↓↓
                        'Case "UCAC"
                        '    Select Case selectedData.Symbols(7).Trim
                        '        Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-T*H-" & _
                        '                                                       selectedData.Symbols(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        '        Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-T*V-" & _
                        '                                                       selectedData.Symbols(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        '        Case "T2YH", "T3YH", "T2WH", "T3WH"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-T*YH-" & _
                        '                                                       selectedData.Symbols(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        '        Case "T2YV", "T3YV", "T2WV", "T3WV"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-T*YV-" & _
                        '                                                       selectedData.Symbols(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        '        Case Else
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        '                                                       selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                        '                                                       selectedData.Symbols(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        '    End Select
                        'Case "UCAC-L2"
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        '                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                        '                                               selectedData.Symbols(8).Trim
                        '    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                        Case "UCAC-L2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7 + intOpt).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                            '↓RM1309001 2013/09/02 追加
                        Case "UCAC2"
                            Select Case selectedData.Symbols(7 + intOpt).Trim
                                Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H", "T3PH", "T2JH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*H-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V", "T3PV", "T2JV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*V-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T2YH", "T3YH", "T2WH", "T3WH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YH-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T2YV", "T3YV", "T2WV", "T3WV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YV-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(7 + intOpt).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If
                            End Select
                            '↑RM1309001 2013/09/02 追加
                        Case Else
                            Select Case selectedData.Symbols(7 + intOpt).Trim
                                Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*H-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*V-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T2YH", "T3YH", "T2WH", "T3WH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YH-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case "T2YV", "T3YV", "T2WV", "T3WV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YV-" & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(7 + intOpt).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                            'RM0811133 2009/07/28 Y.Miura ↑↑
                    End Select
                End If

                '取付用タイロッド加算価格キー
                'RM0811133 2009/07/28 Y.Miura ↓↓
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-TIEROD"
                Select Case strHySeriesKataban
                    Case "UCAC"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-TIEROD"
                    Case "UCAC2"
                        Select Case selectedData.Series.key_kataban.Trim
                            Case "R", "S"
                                'タイロット加算なし
                            Case Else
                                '↓RM1309001 2013/09/02 追加
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                If selectedData.Symbols(11).Trim = "Z" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & "TIEROD" & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               intStroke.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                End Select
                '↓RM1309001 2013/09/02 追加
                ''RM0811133 2009/07/28 Y.Miura ↑↑
                'decOpAmount(UBound(decOpAmount)) = 1
                ''RM0811133 2009/07/28 Y.Miura
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If


            '取付用タイロッド加算価格キー(スイッチなし時) RM1003086 2010/03/26 Y.Miura 追加
            Select Case strHySeriesKataban
                Case "UCAC2"
                    Select Case selectedData.Series.key_kataban.Trim
                        Case "R", "S"
                            'タイロット加算なし
                        Case Else
                            If selectedData.Symbols(12).Trim <> "" Then
                                '2013/09/02 追加
                                If selectedData.Symbols(12).Trim = "Z" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9 + intOpt).Trim)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & "TIEROD" & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               intStroke.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            End If
                    End Select
            End Select

            '付属品加算価格キー
            'RM1003086 2010/03/26 Y.Miura 変更
            'strOpArray = Split(selectedData.Symbols(11 + intOpt), MyControlChars.Comma)
            strOpArray = Split(selectedData.Symbols(11 + intOpt + intOpt2), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        '                                           strOpArray(intLoopCnt).Trim
                        Select Case strHySeriesKataban
                            Case "UCAC"
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                            Case "UCAC2"
                                strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0811133 2009/07/28 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select
            Next

            'スズキ向け特注
            Select Case selectedData.Series.series_kataban.Trim
                Case "UCAC2", "UCAC2-L2"
                    Select Case selectedData.Series.key_kataban.Trim
                        Case "R", "S"
                            'If selectedData.Symbols(8).Trim <> "" Then
                            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            '    strOpRefKataban(UBound(strOpRefKataban)) = "UCAC2-TS-" & _
                            '                                               selectedData.Symbols(14).Trim
                            '    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(10).Trim)
                            'End If
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "UCAC2" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(14).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
