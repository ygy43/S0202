'************************************************************************************
'*  ProgramID  ：KHPriceH7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ショックキラー　ＳＣＫ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceH7

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
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            If selectedData.Symbols(1).Trim <> "00" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
