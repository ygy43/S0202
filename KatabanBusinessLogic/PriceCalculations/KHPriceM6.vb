'************************************************************************************
'*  ProgramID  ：KHPriceM8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/07   作成者：NII K.Sudoh
'*
'*  概要       ：スーパーコンパクトシリンダ　ＳＳＤ
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*                                      更新日：2007/10/23   更新者：NII A.Takahashi
'*               ・ロッド先端特注画面の追加により、ロッド先端特注価格ロジック変更
'*  ・受付No：RM0906034  二次電池対応機器　SSD
'*                                      更新日：2009/08/05   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceM6
    Public Sub subPriceCalculation(selectedData As SelectedInfo,
                                   ByRef strOpRefKataban() As String,
                                   ByRef decOpAmount() As Decimal,
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStrokeS1 = 0
        Dim intStrokeS2 = 0
        Dim bolC5Flag As Boolean

        Dim bolOptionN = False
        Dim bolOptionP5 = False
        Dim bolOptionP51 = False
        Dim bolOptionA2 = False
        Dim bolOptionP4 = False              'RM0906034 2009/08/05 Y.Miura　二次電池対応 追加

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(19), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "N"
                        bolOptionN = True
                    Case "P4", "P40" 'RM0906034 2009/08/05 Y.Miura　二次電池対応 追加
                        bolOptionP4 = True
                    Case "P5"
                        bolOptionP5 = True
                    Case "P51"
                        bolOptionP51 = True
                    Case "A2"
                        bolOptionA2 = True
                End Select
            Next

            'C5チェック
            bolC5Flag = fncCylinderC5Check(selectedData, False)

            'ストローク設定(S1)
            If selectedData.Symbols(1).IndexOf("B") >= 0 Or
               selectedData.Symbols(1).IndexOf("W") >= 0 Then
                intStrokeS1 = KatabanUtility.GetStrokeSize(selectedData,
                                                           CInt(selectedData.Symbols(4).Trim),
                                                           CInt(selectedData.Symbols(7).Trim))
            End If
            'ストローク設定(S2)
            intStrokeS2 = KatabanUtility.GetStrokeSize(selectedData,
                                                       CInt(selectedData.Symbols(4).Trim),
                                                       CInt(selectedData.Symbols(14).Trim))

            '基本価格キー
            If selectedData.Symbols(1).IndexOf("B") >= 0 Or
               selectedData.Symbols(1).IndexOf("W") >= 0 Then
                'S1
                If selectedData.Symbols(6).IndexOf("K") >= 0 Or
                   selectedData.Symbols(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-K-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If

                'S2
                If selectedData.Symbols(13).IndexOf("K") >= 0 Or
                   selectedData.Symbols(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-K-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            Else
                If selectedData.Symbols(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-K-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-BASE-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

            'バリエーション加算価格キー
            '(*B*)背合せ形
            If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-B-" &
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*G*)強力スクレーパ形
            If selectedData.Symbols(1).IndexOf("G") >= 0 And
               selectedData.Symbols(1).IndexOf("G1") < 0 And
               selectedData.Symbols(1).IndexOf("G2") < 0 And
               selectedData.Symbols(1).IndexOf("G3") < 0 And
               selectedData.Symbols(1).IndexOf("G4") < 0 And
               selectedData.Symbols(1).IndexOf("G5") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*G1*)コイルスクレーパ形
            If selectedData.Symbols(1).IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G1-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*G2*)耐切削油スクレーパ形(一般用)
            If selectedData.Symbols(1).IndexOf("G2") >= 0 Then
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    'S1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G2-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G2-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G2-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

            '(*G3*)耐切削油スクレーパ形(塩素系用)
            If selectedData.Symbols(1).IndexOf("G3") >= 0 Then
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    'S1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G3-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G3-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G3-" &
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

            '(*G4*)スパッタ付着防止形
            If selectedData.Symbols(1).IndexOf("G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G4-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*G5*)スパッタ付着防止形
            If selectedData.Symbols(1).IndexOf("G5") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-G5-" &
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*M*)回り止め形
            If selectedData.Symbols(1).IndexOf("M") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-M-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*O*)低速形
            If selectedData.Symbols(1).IndexOf("O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-O-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*Q*)落下防止形
            If selectedData.Symbols(1).IndexOf("Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-Q-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*T*)耐熱形120℃
            If selectedData.Symbols(1).IndexOf("T") >= 0 And
               selectedData.Symbols(1).IndexOf("T1") < 0 And
               selectedData.Symbols(1).IndexOf("T1L") < 0 And
               selectedData.Symbols(1).IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-T-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*T1*)耐熱形150℃
            If selectedData.Symbols(1).IndexOf("T1") >= 0 And
               selectedData.Symbols(1).IndexOf("T1L") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-T1-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*T1L*)耐熱形スイッチ付
            If selectedData.Symbols(1).IndexOf("T1L") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-T1L-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*T2*)パッキン材質フッ素ゴム
            If selectedData.Symbols(1).IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-T2-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*W*)二段形
            If selectedData.Symbols(1).IndexOf("W") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-W-" &
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*X*)押出し形
            If selectedData.Symbols(1).IndexOf("X") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-X-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '(*Y*)引込み形
            If selectedData.Symbols(1).IndexOf("Y") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-Y-" &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション(M)回り止め加算価格キー
            'S1
            Select Case selectedData.Symbols(6).Trim
                Case "M", "KM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-M-" &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select
            'S2
            Select Case selectedData.Symbols(13).Trim
                Case "M", "KM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-M-" &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select

            '微速加算価格キー
            Select Case selectedData.Symbols(3).Trim
                Case "F"
                    If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                       selectedData.Symbols(1).IndexOf("W") >= 0 Then
                        'S1
                        If selectedData.Symbols(1).IndexOf("K") >= 0 Or
                           selectedData.Symbols(6).IndexOf("K") >= 0 Then
                            '内径判定
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 16 And
                                             CInt(selectedData.Symbols(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 16 And
                                             CInt(selectedData.Symbols(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 51 And
                                             CInt(selectedData.Symbols(7).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 51 And
                                             CInt(selectedData.Symbols(7).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 101 And
                                             CInt(selectedData.Symbols(7).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 151 And
                                             CInt(selectedData.Symbols(7).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 51 And
                                             CInt(selectedData.Symbols(7).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 101 And
                                             CInt(selectedData.Symbols(7).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 51 And
                                             CInt(selectedData.Symbols(7).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 101 And
                                             CInt(selectedData.Symbols(7).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If

                        'S2
                        If selectedData.Symbols(1).IndexOf("K") >= 0 Or
                           selectedData.Symbols(13).IndexOf("K") >= 0 Then
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16 And
                                             CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16 And
                                             CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 151 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If
                    Else
                        'S2
                        If selectedData.Symbols(1).IndexOf("K") >= 0 Or
                           selectedData.Symbols(13).IndexOf("K") >= 0 Then
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16 And
                                             CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16 And
                                             CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 151 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-KF-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case selectedData.Symbols(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 51 And
                                             CInt(selectedData.Symbols(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 101 And
                                             CInt(selectedData.Symbols(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) =
                                                selectedData.Series.series_kataban.Trim & "-F-" &
                                                selectedData.Symbols(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If
                    End If
            End Select

            'NPTねじ、Gねじ加算価格キー
            Select Case selectedData.Symbols(5).Trim
                Case "GD", "ND"
                    'D加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SCREW-" &
                                                               Right(selectedData.Symbols(5).Trim, 1) &
                                                               MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Case "D"
                    'D加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SCREW-" &
                                                               selectedData.Symbols(5).Trim & MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select

            'スイッチ付加算価格キー
            If selectedData.Symbols(2).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-" &
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen &
                                                           selectedData.Symbols(4).Trim
                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'スイッチ形番＆リード線長さ加算価格キー
            'S1
            If selectedData.Symbols(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-" &
                                                           selectedData.Symbols(9).Trim
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)

                If selectedData.Symbols(10).Trim <> "" Then
                    Select Case selectedData.Symbols(9).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H",
                            "T3V", "T5H", "T5V", "T2YH", "T2YV",
                            "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V",
                            "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(1)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(2)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(3)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH",
                            "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(4)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(5)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(6)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(7)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(8)-" &
                                                                       selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(11).Trim)
                    End Select
                End If
            End If

            'S2
            If selectedData.Symbols(16).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-" &
                                                           selectedData.Symbols(16).Trim
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)

                If selectedData.Symbols(17).Trim <> "" Then
                    Select Case selectedData.Symbols(16).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H",
                            "T3V", "T5H", "T5V", "T2YH", "T2YV",
                            "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V",
                            "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(1)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(2)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(3)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH",
                            "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(4)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(5)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(6)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(7)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-SWLW(8)-" &
                                                                       selectedData.Symbols(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                    End Select
                End If
                'RM0906034 2009/08/05 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(18).Trim)
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(19), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "M"
                        Select Case selectedData.Symbols(4).Trim
                            Case "12", "16", "20", "25"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                           "-OP-" &
                                                                           strOpArray(intLoopCnt).Trim &
                                                                           MyControlChars.Hyphen &
                                                                           selectedData.Symbols(4).Trim
                                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                            Case "32", "40", "50", "63", "80",
                                "100", "125", "140", "160"
                                If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                                   selectedData.Symbols(1).IndexOf("W") >= 0 Then
                                    'S1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                               "-OP-" &
                                                                               strOpArray(intLoopCnt).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               selectedData.Symbols(4).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               intStrokeS1.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If

                                    'S2
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                               "-OP-" &
                                                                               strOpArray(intLoopCnt).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               selectedData.Symbols(4).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               intStrokeS2.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                               "-OP-" &
                                                                               strOpArray(intLoopCnt).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               selectedData.Symbols(4).Trim &
                                                                               MyControlChars.Hyphen &
                                                                               intStrokeS2.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If
                                End If
                        End Select

                        '背合せ形＆二段形の時(+α加算する)
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-OP-(B/W)" &
                                                                       strOpArray(intLoopCnt).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        End If
                    Case "M1"
                        '背合せ形＆二段形の時
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            'S1
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                       strOpArray(intLoopCnt).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       selectedData.Symbols(4).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       intStrokeS1.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If

                            'S2
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                       strOpArray(intLoopCnt).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       selectedData.Symbols(4).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       intStrokeS2.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                       strOpArray(intLoopCnt).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       selectedData.Symbols(4).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       intStrokeS2.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        End If

                        '背合せ形＆二段形の時(+α加算する)
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                                       "-OP(B/W)" &
                                                                       strOpArray(intLoopCnt).Trim &
                                                                       MyControlChars.Hyphen &
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        End If
                    Case "N"
                        '￥0
                    Case "S"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                        'RM0906034 2009/08/05 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "P5", "P51"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & "*" &
                                                                   MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                        '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                    Case "P6", "R1", "R2"
                        'Case "P6"
                        '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case "P7", "P71"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & "*" &
                                                                   MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Or
                           selectedData.Symbols(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" &
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) &
                                                                   MyControlChars.Hyphen &
                                                                   selectedData.Symbols(4).Trim
                        If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                            Select Case True
                                Case selectedData.Symbols(12).Trim = "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case selectedData.Symbols(12).Trim <> "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Symbols(12).Trim = "N" And bolOptionN = False
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select
            Next

            '支持金具加算価格キー
            If selectedData.Symbols(20).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim &
                                                           MyControlChars.Hyphen &
                                                           selectedData.Symbols(20).Trim & MyControlChars.Hyphen &
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '付属品加算価格キー
            Select Case selectedData.Symbols(21).Trim
                Case "I", "I2", "Y", "Y2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-ACC-" &
                                                               selectedData.Symbols(21).Trim & MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    If selectedData.Symbols(1).IndexOf("B") >= 0 Then
                        Select Case True
                            Case selectedData.Symbols(12).Trim = "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case selectedData.Symbols(12).Trim <> "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case selectedData.Symbols(12).Trim = "N" And bolOptionN = False
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "IY"
                    'I加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-ACC-" &
                                                               Left(selectedData.Symbols(21).Trim, 1) &
                                                               MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-ACC-" &
                                                               Right(selectedData.Symbols(21).Trim, 1) &
                                                               MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "I2Y2"
                    'I2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-ACC-" &
                                                               Left(selectedData.Symbols(21).Trim, 2) &
                                                               MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-ACC-" &
                                                               Right(selectedData.Symbols(21).Trim, 2) &
                                                               MyControlChars.Hyphen &
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'ロッド先端オーダーメイド加算価格キー
            If selectedData.RodEnd.RodEndOption.Trim <> "" Then
                If InStr(1, selectedData.RodEnd.RodEndOption.Trim, "WF") = 0 Then
                    decWFLength = 1
                Else
                    For intLoopCnt = InStr(1, selectedData.RodEnd.RodEndOption.Trim, "WF") + 2 To _
                        Len(selectedData.RodEnd.RodEndOption.Trim)
                        If Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "0" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "1" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "2" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "3" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "4" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "5" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "6" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "7" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "8" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "9" Or
                           Mid(selectedData.RodEnd.RodEndOption.Trim, intLoopCnt, 1) = "." Then
                            If intLoopCnt = Len(selectedData.RodEnd.RodEndOption.Trim) Then
                                decLength = intLoopCnt - (InStr(1, selectedData.RodEnd.RodEndOption.Trim, "WF") + 2) + 1
                            End If
                        Else
                            decLength = intLoopCnt - (InStr(1, selectedData.RodEnd.RodEndOption.Trim, "WF") + 2) + 1
                            Exit For
                        End If
                    Next

                    decWFLength =
                        CDec(Mid(selectedData.RodEnd.RodEndOption.Trim, InStr(1, selectedData.RodEnd.RodEndOption.Trim, "WF") + 2,
                                 decLength)) - selectedData.RodEnd.RodEndWFStdVal
                End If

                Select Case True
                    Case 0 <= decWFLength And decWFLength <= 100
                        strStdWFLength = "100"
                    Case 101 <= decWFLength And decWFLength <= 200
                        strStdWFLength = "200"
                    Case 201 <= decWFLength And decWFLength <= 300
                        strStdWFLength = "300"
                    Case 301 <= decWFLength And decWFLength <= 400
                        strStdWFLength = "400"
                    Case 401 <= decWFLength And decWFLength <= 500
                        strStdWFLength = "500"
                    Case 501 <= decWFLength And decWFLength <= 600
                        strStdWFLength = "600"
                    Case 601 <= decWFLength
                        strStdWFLength = "700"
                End Select
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) &
                                                           "-TIP-OF-ROD-" &
                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen &
                                                           strStdWFLength
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

            End If

        Catch ex As Exception

            Throw ex

        Finally


        End Try
    End Sub
End Module
