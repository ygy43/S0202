'************************************************************************************
'*  ProgramID  ：KHPriceA8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン　ＲＧＩＢ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceA8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                       Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                       selectedData.Symbols(1).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー
            For intLoopCnt = 8 To 9
                Select Case selectedData.Symbols(intLoopCnt).Trim
                    Case "C", "F", "P"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(intLoopCnt)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "H"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(intLoopCnt) & _
                                                                   selectedData.Symbols(12)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next


        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
