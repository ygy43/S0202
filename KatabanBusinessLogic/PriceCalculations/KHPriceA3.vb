'************************************************************************************
'*  ProgramID  ：KHPriceA3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ギャップスイッチユニット　ＵＧＰＳ２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceA3

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
            Select Case selectedData.Symbols(6).Trim
                Case "C0", "C1", "C3", "C5"
                    'コネクタ形ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "C" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "F"
                    'DIN端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "CTL", "CTR"
                    'コネクタ形集中端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "CT" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "TL", "TR", "T1", "T2", "T3", "T4"
                    'リード線形集中端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "T" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '配線オプション加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(6).Trim
            decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(3).Trim)

            'アタッチメント加算価格キー
            strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(3).Trim)
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            '圧力計加算価格キー
            If selectedData.Symbols(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(9).Trim
                decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(3).Trim)
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
