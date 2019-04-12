'************************************************************************************
'*  ProgramID  ：KHPriceG7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：薬液用マニュアルバルブ
'*             ：ＭＭＤ３０２
'*             ：ＭＭＤ４０２
'*             ：ＭＭＤ５０２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceG7

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '流体加算価格キー
            If selectedData.Symbols(3).Trim = "P" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "*" & _
                                                           Mid(selectedData.Series.series_kataban.Trim, 6, 1) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If


            '取付加算価格キー
            If selectedData.Symbols(4).Trim = "X" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                           MyControlChars.Hyphen & selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
