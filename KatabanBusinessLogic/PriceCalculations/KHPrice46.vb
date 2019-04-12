'************************************************************************************
'*  ProgramID  ：KHPrice43
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーマウントシリンダ　ＳＭＤ２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice46

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(3).Trim))

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            'マグネット内臓(L)加算価格キー
            If Mid(selectedData.Series.series_kataban.Trim, 2, 1) = "M" And _
              (Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '支持形式加算
            If Mid(selectedData.Series.series_kataban.Trim, 2, 1) = "M" And _
               selectedData.Symbols(1).Trim = "DC" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            Select Case selectedData.Series.series_kataban.Trim
                Case "SMD2-ML", "SMD2-XL", "SMD2-YL"
                    If selectedData.Symbols(4).Trim <> "" Then
                        'スイッチ加算価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)

                        'リード線長さ加算価格
                        If selectedData.Symbols(5).Trim <> "" Then
                            Select Case selectedData.Symbols(4).Trim
                                Case "K0H", "K0V", "K2H", "K2V", "K3H", _
                                     "K3V", "K5H", "K5V", "K2YH", "K2YV", _
                                     "K3YH", "K3YV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                                Case "K2YFH", "K2YFV", "K3YFH", "K3YFV", "K2YMH", _
                                     "K2YMV", "K3YMH", "K3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(5).Trim & "Y"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                            End Select
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
