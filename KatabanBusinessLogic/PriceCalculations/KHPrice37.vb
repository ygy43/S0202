'************************************************************************************
'*  ProgramID  ：KHPrice37
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＨＢ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice37

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            Select Case True
                Case Left(selectedData.Symbols(4), 1) = "1" Or Left(selectedData.Symbols(4), 1) = "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & _
                                                               MyControlChars.Hyphen & "1"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Left(selectedData.Symbols(4), 1) = "3" Or Left(selectedData.Symbols(4), 1) = "4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & _
                                                               MyControlChars.Hyphen & "3"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & _
                                                               MyControlChars.Hyphen & "5"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション価格
            strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "B3"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "MB3"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "L3"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "ML3"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B", "B1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "MB"
                        decOpAmount(UBound(decOpAmount)) = 1
                        '2011/03/07 ADD RM1103016(4月VerUP：HVB,PDV2,AB71(他)シリーズ) START--->
                    Case "3M", "3N"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "M" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        '2011/03/07 ADD RM1103016(4月VerUP：HVB,PDV2,AB71(他)シリーズ) <---END
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "MG"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算
            Select Case selectedData.Symbols(6).Trim
                Case "AC100V", "AC200V"
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & selectedData.Symbols(1).Trim & _
                                                                     Left(selectedData.Symbols(6).Trim, 2)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
