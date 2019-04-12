'************************************************************************************
'*  ProgramID  ：KHPriceL7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン
'*             ：ＲＧＩＤ
'*             ：ＲＧＣＤ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceL7

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

            '入力仕様加算価格キー
            For intLoopCnt = 8 To 10
                If selectedData.Symbols(intLoopCnt).Trim <> "" And _
                   selectedData.Symbols(intLoopCnt).Trim <> "N" And _
                   (intLoopCnt <> 9 Or selectedData.Symbols(8).Trim <> "K" Or selectedData.Symbols(9).Trim <> "K") Then
                    If intLoopCnt = 10 Then
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "O" & _
                                                                   selectedData.Symbols(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                   Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "I" & _
                                                                   selectedData.Symbols(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
