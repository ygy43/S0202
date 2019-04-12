'************************************************************************************
'*  ProgramID  ：KHPrice15
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*
'*  概要       ：リニアスライドシリンダ　ＬＣＳ／ＬＣＳ－Ｑ／ＬＣＳ－Ｆ
'*
'*  更新履歴   ：                       更新日：2007/06/25   更新者：NII A.Takahashi
'*               ・選択ボックス追加/LCS-Fをシリーズ形番に追加するため修正
'*  ・受付No：RM0906034  二次電池対応機器　LCS
'*                                      更新日：2009/08/20   更新者：Y.Miura
'*  ・受付No：RM0912XXX  二次電池対応機器のC5価格適用
'*                                      更新日：2009/12/09   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice15

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer
        Dim bolOptionP4 As Boolean = False              'RM0906034 2009/08/20 Y.Miura　二次電池対応
        Dim strOptionP4 As String = String.Empty        'RM0906034 2009/08/28 Y.Miura　二次電池対応
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean
        Dim intSeries As Integer                         'RM1103016 2011/03/04 ADD
        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応

            'オプション加算価格キー
            'RM0906034 2009/08/20 Y.Miura　二次電池対応
            strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                        strOptionP4 = strOpArray(intLoopCnt).Trim
                End Select
            Next

            'C5チェック
            '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
            Select Case selectedData.Series.series_kataban
                Case "LCS2", "LCS2-Q"
                    '"LCS2", "LCS2-Q"の場合は、C5だが、P4加算を行わない
                    bolC5Flag = False
                    'シリーズ名桁数
                    intSeries = 4

                Case Else
                    '上記以外、C5ではない

                    'RM1001043 2010/02/22 Y.Miura  二次電池C5加算廃止
                    'RM0906034 2009/08/20 Y.Miura　二次電池対応
                    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                    'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
                    'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
                    bolC5Flag = False
                    'シリーズ名桁数
                    intSeries = 3
            End Select
            '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END

            'ストローク算出
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  selectedData.Symbols(1).Trim, _
                                                  selectedData.Symbols(3).Trim)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, intSeries) & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
            '                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
            '                                           intStroke.ToString
            '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            'バリエーション加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                Case "LCS-F", "LCS-Q", "LCS2-Q"
                    'Case "LCS-F", "LCS-Q"
                    '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            If selectedData.Symbols(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, intSeries) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(4).Trim
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)

                'リード線長さ加算
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                Dim isADD As Boolean = False
                'リード線長さがブランク以外の場合
                If selectedData.Symbols(5).Trim <> "" Then
                    'シリーズ判定
                    Select Case selectedData.Series.series_kataban
                        Case "LCS2", "LCS2-Q"
                            '"LCS2", "LCS2-Q"の場合、加算
                            isADD = True
                        Case Else
                            '上記以外、スイッチが"T2VR"以外の場合加算
                            If selectedData.Symbols(4).Trim <> "T2VR" Then
                                isADD = True
                            End If
                    End Select

                    If isADD Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                    End If
                End If
                'If selectedData.Symbols(5).Trim <> "" Then
                '    If selectedData.Symbols(4).Trim <> "T2VR" Then
                '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                '                                                   selectedData.Symbols(5).Trim
                '        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                '    End If
                'End If
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END

                'RM0906034 2009/08/20 Y.Miura　二次電池対応
                'P4加算
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)

                End If
            End If

            'オプション加算価格キー
            If selectedData.Symbols(7).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, intSeries) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(7).Trim & _
                                                           selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(7).Trim & _
                '                                           selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(1).Trim
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

                'RM0906034 2009/08/28 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    Select Case selectedData.Symbols(7).Trim
                        Case "A1", "A2", "A3", "A4" 'ショックキラー付
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                      "A" & MyControlChars.Hyphen & _
                                                                      strOptionP4 & MyControlChars.Hyphen & _
                                                                      selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "A5", "A6"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                      "A" & MyControlChars.Hyphen & _
                                                                      strOptionP4 & MyControlChars.Hyphen & _
                                                                      selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
                End If
            End If

            'オプション加算価格キー
            If selectedData.Symbols(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, intSeries) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                '                                           selectedData.Symbols(1).Trim
                '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'クリーン仕様加算価格キー
            'RM0906034 2009/08/20 Y.Miura　二次電池対応
            'If selectedData.Symbols(10).Trim <> "" Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
            '                                               selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
            '                                               selectedData.Symbols(1).Trim
            '    decOpAmount(UBound(decOpAmount)) = 1
            'End If
            strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) START--->
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, intSeries) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                        '                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                        '                                           selectedData.Symbols(1).Trim
                        '2011/03/04 MOD RM1103016(4月VerUP:LCS2シリーズ追加) <---END
                        decOpAmount(UBound(decOpAmount)) = 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P4", "P40"
                            Case Else
                                If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                        End Select
                End Select
            Next


        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
