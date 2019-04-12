'************************************************************************************
'*  ProgramID  ：KHPriceN0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*
'*  概要       ：セルトップシリンダ　ＪＳＫ２
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceN0

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim strOption As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(3).Trim))

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                       selectedData.Symbols(1).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '特殊電圧加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "JSK2-V"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'スイッチ加算
            Select Case selectedData.Series.series_kataban.Trim
                Case "JSK2"
                    'スイッチ加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-" & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(5).Trim <> "" Then
                            Select Case selectedData.Symbols(4).Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                                     "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                     "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(1)-" & _
                                                                               selectedData.Symbols(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                                Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(2)-" & _
                                                                               selectedData.Symbols(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                                Case "T2JH", "T2JV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(3)-" & _
                                                                               selectedData.Symbols(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                            End Select
                        End If
                    End If
                Case "JSK2-V"
                    'スイッチ加算価格キー
                    If selectedData.Symbols(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-" & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(6).Trim <> "" Then
                            Select Case selectedData.Symbols(5).Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                                     "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                     "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(1)-" & _
                                                                               selectedData.Symbols(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                                Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(2)-" & _
                                                                               selectedData.Symbols(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                                Case "T2JH", "T2JV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SWLW(3)-" & _
                                                                               selectedData.Symbols(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                            End Select
                        End If
                    End If
            End Select

            'オプション・付属品加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "JSK2"
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                Case Else
                    strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
            End Select
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                   strOpArray(intLoopCnt).Trim

                        Select Case strOpArray(intLoopCnt).Trim
                            Case "J", "L"
                                strOption = "20"
                            Case "M"
                                Select Case selectedData.Symbols(2).Trim
                                    Case "25", "30"
                                        strOption = "20"
                                    Case Else
                                        strOption = selectedData.Symbols(2).Trim
                                End Select
                            Case Else
                                strOption = selectedData.Symbols(2).Trim
                        End Select

                        Select Case True
                            Case strOpArray(intLoopCnt).Trim = "J" Or _
                                 strOpArray(intLoopCnt).Trim = "L" Or _
                                 Mid(selectedData.Series.series_kataban.Trim, 4, 1) = "3" And strOpArray(intLoopCnt).Trim = "M" Or (Mid(selectedData.Series.series_kataban.Trim, 3, 1) = "K" Or Mid(selectedData.Series.series_kataban.Trim, 3, 1) = "M") And strOpArray(intLoopCnt).Trim = "M" And (selectedData.Symbols(2).Trim <> "20" And selectedData.Symbols(2).Trim <> "25" And selectedData.Symbols(2).Trim <> "30")
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & strOption & MyControlChars.Hyphen & intStroke.ToString
                            Case strOpArray(intLoopCnt).Trim = "N" Or _
                                 strOpArray(intLoopCnt).Trim = "V" Or _
                                 strOpArray(intLoopCnt).Trim = "P" Or _
                                 strOpArray(intLoopCnt).Trim = "M" And _
                                 (selectedData.Symbols(2).Trim = "20" Or selectedData.Symbols(2).Trim = "25" Or selectedData.Symbols(2).Trim = "30")
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & strOption
                        End Select

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '付属品加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "JSK2"
                    strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
                Case Else
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
            End Select
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
