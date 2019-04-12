'************************************************************************************
'*  ProgramID  ：KHPriceM0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン　ＲＧ＊＊／ＰＣ＊Ｓ
'*             ：（ＨＯ減速機取付用インデックスマン本体形番用）
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceM0

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Left(selectedData.Symbols(7).Trim, 1) >= "0" And _
               Left(selectedData.Symbols(7).Trim, 1) <= "9" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "W" & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "FC"
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & MyControlChars.Hyphen & "W" & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "AL"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'ＨＯ減速機取付用加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                       Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "HO"
            decOpAmount(UBound(decOpAmount)) = 1

            Select Case Left(Trim(selectedData.Series.series_kataban.Trim), 4)
                Case "RGCS", "RGIL", "RGIS", "RGOL", "RGOS", "PCIS", "PCOS"
                    If selectedData.Symbols(12).Trim <> "" Then
                        Select Case selectedData.Symbols(10).Trim
                            Case "F", "A", "S", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "TSF" & _
                                                                           selectedData.Symbols(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "X", "C", "Y", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*" & _
                                                                           Mid(selectedData.Series.series_kataban.Trim, 4, 1) & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "TGX" & _
                                                                           selectedData.Symbols(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
