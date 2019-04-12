'************************************************************************************
'*  ProgramID  ：KHPriceN1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ウォータハンマ低減バルブ　ＡＭＤ＊Ｌ１
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceN1

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
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                       selectedData.Symbols(1).Trim & _
                                                       selectedData.Symbols(2).Trim & _
                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(5).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '補強リング加算価格キー
            If selectedData.Symbols(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & _
                                                           selectedData.Symbols(1).Trim & _
                                                           selectedData.Symbols(2).Trim & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
