'************************************************************************************
'*  ProgramID  ：KHPriceH2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：レギュレータ・リバースレギュレータ
'*             ：C3000/C3010/C3020/C3030/C3040/C3050/C3060/C3070
'*             ：C4000/C4010/C4020/C4030/C4040/C4050/C4060/C4070
'*             ：C8000/C8010/C8020/C8030/C8040/C8050/C8060/C8070
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceH2

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Left(selectedData.Series.series_kataban.Trim, 2) = "C4" Then
                If Left(selectedData.Symbols(1).Trim, 2) = "20" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "20"
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "F", "F1", "FF", "FF1"
                        Select Case selectedData.Series.series_kataban.Trim
                            Case "C3030", "C3040", "C3060", "C4030", "C4040", _
                                 "C4060", "C8030", "C8040", "C8060"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-30-40-60"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C3070", "C4070", "C8070"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-70"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "Z", "M"
                        Select Case selectedData.Series.series_kataban.Trim
                            Case "C1020", "C1050", "C3020", "C3050", "C4020", _
                                 "C4050", "C8020", "C8050"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-20-50"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C3070", "C4070", "C8070"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-70"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "Q"
                        Select Case selectedData.Series.series_kataban.Trim
                            Case "C8030", "C8060"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-30-60"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C8070"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim & "-70"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '組付けアタッチメント加算価格キー
            If selectedData.Symbols(4).Trim <> "" Then
                Select Case selectedData.Series.series_kataban.Trim
                    Case "C4020", "C4030"
                        If (selectedData.Symbols(1).Trim = "20" Or _
                        selectedData.Symbols(1).Trim = "20N" Or _
                        selectedData.Symbols(1).Trim = "20G") And _
                           selectedData.Symbols(4).IndexOf("S") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "20" & MyControlChars.Hyphen
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen
                End Select

                strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & strOpArray(intLoopCnt).Trim
                    End Select
                Next
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'アタッチメント加算価格キー
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        Select Case True
                            Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 3)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
