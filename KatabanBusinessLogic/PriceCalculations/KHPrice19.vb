'************************************************************************************
'*  ProgramID  ：KHPrice19
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＰＫＡ／ＰＫＳ／ＰＫＷ／ＰＶＳ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice19

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim strStdVoltageFlag As String = Divisions.VoltageDiv.Standard
        Dim strOp As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            If Mid(selectedData.Series.series_kataban.Trim, 2, 1) = "K" Then
                If selectedData.Symbols(3).Trim = "C" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If selectedData.Symbols(1).Trim = "4" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    If selectedData.Symbols(3).Trim = "NO" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

            'オプション価格
            If selectedData.Series.series_kataban.Trim = "PKS" Then
                strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
            Else
                strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            End If
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算
            Select Case selectedData.Series.series_kataban.Trim
                Case "PKA", "PVS"
                    strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                   selectedData.Symbols(6).Trim)
                    If selectedData.Symbols(5).Trim = "M" Then
                        strOp = Left(selectedData.Symbols(6).Trim, 2) & "M"
                    Else
                        strOp = Left(selectedData.Symbols(6).Trim, 2)
                    End If
                Case "PKS"
                    strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                   selectedData.Symbols(4).Trim)
                    strOp = Left(selectedData.Symbols(4).Trim, 2)
                Case "PKW"
                    strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                   selectedData.Symbols(6).Trim)
                    If selectedData.Symbols(5).Trim = "M" Then
                        strOp = Left(selectedData.Symbols(6).Trim, 2) & "M"
                    Else
                        Select Case selectedData.Symbols(6).Trim
                            Case "AC100V", "AC200V"
                                strOp = Left(selectedData.Symbols(6).Trim, 2) & "M"
                            Case Else
                                strOp = Left(selectedData.Symbols(6).Trim, 2)
                        End Select
                    End If
            End Select
            If strStdVoltageFlag = Divisions.VoltageDiv.Standard Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & _
                                                           "STD" & strOp
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & _
                                                           "OTH" & strOp
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
