'************************************************************************************
'*  ProgramID  ：KHPrice40
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＨＶＢ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice40

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
                Case ""
                    '価格キー設定
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                '価格キー設定
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                    If selectedData.Symbols(5).Equals("AC110V") OrElse _
                        selectedData.Symbols(5).Equals("AC220V") Then
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If
                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END

                Case "1"
                    '価格キー設定
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "2CS", "2G", "2HS"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    'その他オプション加算価格キー設定
                    If selectedData.Symbols(6).Trim = "B" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "2"
                    '価格キー設定
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション価格
                    strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next

                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                    If selectedData.Symbols(7).Equals("AC110V") OrElse _
                        selectedData.Symbols(7).Equals("AC220V") Then
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If
                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
