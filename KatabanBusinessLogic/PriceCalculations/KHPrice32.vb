'************************************************************************************
'*  ProgramID  ：KHPrice32
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*
'*  概要       ：セルシリンダ　ＣＫＶ２／ＣＫＶ２－Ｍ
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice32

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

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
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            Select Case selectedData.Symbols(1).Trim
                Case "TA", "TB"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                               selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '結線方法加算価格キー
            If selectedData.Symbols(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If selectedData.Symbols(6).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & "SW" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(6).Trim
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                'リード線長さ加算価格キー
                If selectedData.Symbols(7).Trim <> "" Then
                    Select Case selectedData.Symbols(6).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V", "T2YH", "T2YV", "T3YH", "T3YV", _
                             "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & "SWLW(1)" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & "SWLW(2)" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & "SWLW(3)" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                    End Select
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "N"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "J", "K", "L"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   intStroke.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "M"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   intStroke.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
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
