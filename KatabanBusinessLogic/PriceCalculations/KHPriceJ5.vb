'************************************************************************************
'*  ProgramID  ：KHPriceJ5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＦＲユニット　Ｂ７０１９
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceJ5

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
            If selectedData.Symbols(3).Trim = "P6" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
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
                    Case ""
                    Case "-G"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   Mid(strOpArray(intLoopCnt).Trim, 2, 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'アタッチメント加算価格キー
            Select Case selectedData.Series.series_kataban
                Case "BW7019"
                    strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                           Select strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
