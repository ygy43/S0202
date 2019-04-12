'************************************************************************************
'*  ProgramID  ：KHPriceJ2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インライン形クリーンフィルタ　ＦＣＳ５００／ＦＣＳ１０００
'*
'*　変更       ：二次電池機種追加に伴いプログラムを改修　RM1001045 2010/02/24 Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceJ2

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.key_kataban.Trim
                Case "1"        'FCS1000,FCS500
                    '基本価格キー： 機種シリーズ-BASE-モータサイズ-ストローク
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'その他オプション加算
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case Else
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '添付加算価格キー
                    If selectedData.Symbols(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
