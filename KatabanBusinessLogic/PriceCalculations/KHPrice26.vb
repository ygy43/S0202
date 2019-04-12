'************************************************************************************
'*  ProgramID  ：KHPrice26
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ユニットシリンダ　ＵＣＡ２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice26

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionP4 As Boolean = False      'RM1001045 2010/02/23 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'RM1306001 2013/06/06
            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)

            'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
            Select Case True
                Case Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or _
                     Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L"
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                Case Else
                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            End Select
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                End Select
            Next

            '基本価格キー
            Select Case True
                Case Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "B"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'RM1306001 2013/06/05 追加
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            If Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                'L(スイッチ付き)加算
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                If selectedData.Symbols(4).Trim <> "" Then
                    'スイッチ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    If selectedData.Symbols(5).Trim <> "" Then
                        'リード線の長さ加算価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                    End If

                    'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                    'P4加算
                    If bolOptionP4 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-P4"
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                    End If

                End If
            End If

            'オプション加算価格キー
            'RM1001045 記述を先頭へ
            'Select Case True
            '    Case Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L"
            '        strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
            '    Case Else
            '        strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            'End Select
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "P4", "P40"        'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & "-" & selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1306001 2013/06/05 追加
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1306001 2013/06/05 追加
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select

            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
