'************************************************************************************
'*  ProgramID  ：KHPriceD8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：リニアノームセンサ付シリンダ
'*             ：ＬＮ
'*             ：ＢＨ＊－ＬＮ
'*             ：ＳＳＤ－ＬＮ
'*             ：ＬＣＳ－ＬＮ
'*
'*  更新履歴   ：                       更新日：2007/06/25   更新者：NII A.Takahashi
'*               ・SSD-LN/LCS-LN追加によりロジック修正
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceD8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case Left(selectedData.Series.series_kataban.Trim, 2)
                Case "LN"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & _
                                                               selectedData.Symbols(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "BH"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim & _
                                                               selectedData.Symbols(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "SS", "LC"
                    '基本価格キー
                    If selectedData.Symbols(1).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & _
                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & _
                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & _
                                                                   selectedData.Symbols(12).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
