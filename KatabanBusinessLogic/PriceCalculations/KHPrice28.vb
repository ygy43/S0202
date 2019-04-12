'************************************************************************************
'*  ProgramID  ：KHPrice28
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/28   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーマイクロシリンダ　ＳＣＭ
'*
'*  ・受付No：RM0907070  二次電池対応機器　SCM
'*                                      更新日：2009/08/21   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice28

#Region " Definition "

    Private bolC5Flag As Boolean
    Private strSelStrokeS1() As String = Nothing
    Private strSelStrokeS2() As String = Nothing
    Dim bolOptionP4 As Boolean                'RM0907070 2009/08/21 Y.Miura　二次電池対応

#End Region

    Public Sub subPriceCalculation(selectedData As SelectedInfo,
                                   ByRef strOpRefKataban() As String,
                                   ByRef decOpAmount() As Decimal,
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)
            ReDim strSelStrokeS1(1)
            ReDim strSelStrokeS2(1)

            bolOptionP4 = False     'RM0907070 2009/08/21 Y.Miura　二次電池対応

            'C5チェック
            bolC5Flag = fncCylinderC5Check(selectedData, False)

            '基本タイプ毎に設定
            Select Case selectedData.Series.key_kataban
                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                'Case ""
                Case "", "4", "F"
                    '基本ベース
                    Call subStandardBase(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "B", "G"
                    '背合せ・二段形ベース
                    Call subDoubleRodBase(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "D", "H"
                    '両ロッドベース
                    Call subHighLoadBase(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*  ProgramID  ：subStandardBase
    '*  Program名  ：基本ベース
    '************************************************************************************
    Private Sub subStandardBase(selectedData As SelectedInfo,
                                ByRef strOpRefKataban() As String,
                                ByRef decOpAmount() As Decimal,
                                Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'RM0907070 2009/08/21 Y.Miura　二次電池対応
            Dim strOpArray() As String
            Dim intLoopCnt As Integer
            strOpArray = Split(selectedData.Symbols(13), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                End Select
            Next

            'ストローク設定
            strSelStrokeS1(0) = selectedData.Symbols(6).Trim
            strSelStrokeS1(1) = CStr(KatabanUtility.GetStrokeSize(selectedData,
                                                       CInt(selectedData.Symbols(3).Trim),
                                                       CInt(selectedData.Symbols(6).Trim)))

            '基本価格キー
            If selectedData.Symbols(1).IndexOf("P") < 0 Then
                If selectedData.Symbols(5).Trim = "D" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                               selectedData.Symbols(3).Trim & "D-" &
                                                               strSelStrokeS1(1)
                    If selectedData.Symbols(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                               selectedData.Symbols(3).Trim & "B-" &
                                                               strSelStrokeS1(1)
                    If selectedData.Symbols(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            Else
                If selectedData.Symbols(5).Trim = "D" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-D-" &
                                                               selectedData.Symbols(3).Trim & "D-" &
                                                               strSelStrokeS1(1)
                    If selectedData.Symbols(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-D-" &
                                                               selectedData.Symbols(3).Trim & "B-" &
                                                               strSelStrokeS1(1)
                    If selectedData.Symbols(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(selectedData,
                                 strOpRefKataban,
                                 decOpAmount,
                                 strPriceDiv,
                                 selectedData.Symbols(1).Trim,
                                 selectedData.Symbols(3).Trim,
                                 strSelStrokeS1(1))

            '支持形式加算価格キー
            Call subSCMSupport(selectedData,
                               strOpRefKataban,
                               decOpAmount,
                               strPriceDiv,
                               selectedData.Symbols(2).Trim,
                               selectedData.Symbols(3).Trim)

            'スイッチ加算価格キー
            Call subSCMSwitch(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(9).Trim,
                              selectedData.Symbols(11).Trim)

            'リード線長さ加算価格キー
            Call subSCMSwitchLead(selectedData,
                                  strOpRefKataban,
                                  decOpAmount,
                                  strPriceDiv,
                                  selectedData.Symbols(9).Trim,
                                  selectedData.Symbols(10).Trim,
                                  selectedData.Symbols(11).Trim)

            'スイッチ取付け方式加算価格キー
            Call subSCMSwitchJoint(selectedData,
                                   strOpRefKataban,
                                   decOpAmount,
                                   strPriceDiv,
                                   selectedData.Symbols(9).Trim,
                                   selectedData.Symbols(12).Trim,
                                   selectedData.Symbols(11).Trim,
                                   selectedData.Symbols(3).Trim,
                                   selectedData.Symbols(6).Trim)

            'オプション加算価格キー
            Call subSCMOption(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(13).Trim,
                              selectedData.Symbols(3).Trim,
                              strSelStrokeS1,
                              selectedData.Symbols(9).Trim)


            Select Case selectedData.Series.key_kataban
                Case "F"
                    'オプション加算価格キー
                    Call subSCMOption(selectedData,
                            strOpRefKataban,
                            decOpAmount,
                            strPriceDiv,
                            selectedData.Symbols(14).Trim,
                            selectedData.Symbols(3).Trim,
                            strSelStrokeS1,
                            selectedData.Symbols(9).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                             strOpRefKataban,
                             decOpAmount,
                             strPriceDiv,
                             selectedData.Symbols(15).Trim,
                             selectedData.Symbols(3).Trim,
                             String.Empty)
                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                             strOpRefKataban,
                             decOpAmount,
                             strPriceDiv,
                             selectedData.Symbols(14).Trim,
                             selectedData.Symbols(3).Trim,
                             selectedData.Symbols(15).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(15).Trim,
                                        selectedData.Symbols(3).Trim)

            End Select



        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*  ProgramID  ：subDoubleRodBase
    '*  Program名  ：背合せ・二段形ベース
    '************************************************************************************
    Private Sub subDoubleRodBase(selectedData As SelectedInfo,
                                 ByRef strOpRefKataban() As String,
                                 ByRef decOpAmount() As Decimal,
                                 Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'ストローク設定(S1)
            strSelStrokeS1(0) = selectedData.Symbols(6).Trim
            strSelStrokeS1(1) = CStr(KatabanUtility.GetStrokeSize(selectedData,
                                                               CInt(selectedData.Symbols(3).Trim),
                                                               CInt(selectedData.Symbols(6).Trim)))
            'ストローク設定(S2)
            strSelStrokeS2(0) = selectedData.Symbols(12).Trim
            strSelStrokeS2(1) = KatabanUtility.GetStrokeSize(selectedData,
                                                    CInt(selectedData.Symbols(3).Trim),
                                                    CInt(selectedData.Symbols(12).Trim))

            '基本価格キー
            'S1
            If selectedData.Symbols(5).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                           selectedData.Symbols(3).Trim & "D-" &
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                           selectedData.Symbols(3).Trim & "B-" &
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If
            'S2
            If selectedData.Symbols(11).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                           selectedData.Symbols(3).Trim & "D-" &
                                                           strSelStrokeS2(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-" &
                                                           selectedData.Symbols(3).Trim & "B-" &
                                                           strSelStrokeS2(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(selectedData,
                                 strOpRefKataban,
                                 decOpAmount,
                                 strPriceDiv,
                                 selectedData.Symbols(1).Trim,
                                 selectedData.Symbols(3).Trim,
                                 strSelStrokeS1(1),
                                 strSelStrokeS2(1))

            '支持形式加算価格キー
            Call subSCMSupport(selectedData,
                               strOpRefKataban,
                               decOpAmount,
                               strPriceDiv,
                               selectedData.Symbols(2).Trim,
                               selectedData.Symbols(3).Trim)

            'スイッチ加算価格キー
            'S1
            Call subSCMSwitch(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(7).Trim,
                              selectedData.Symbols(9).Trim)
            'S2
            Call subSCMSwitch(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(13).Trim,
                              selectedData.Symbols(15).Trim)

            'リード線長さ加算価格キー
            'S1
            Call subSCMSwitchLead(selectedData,
                                  strOpRefKataban,
                                  decOpAmount,
                                  strPriceDiv,
                                  selectedData.Symbols(7).Trim,
                                  selectedData.Symbols(8).Trim,
                                  selectedData.Symbols(9).Trim)
            'S2
            Call subSCMSwitchLead(selectedData,
                                  strOpRefKataban,
                                  decOpAmount,
                                  strPriceDiv,
                                  selectedData.Symbols(13).Trim,
                                  selectedData.Symbols(14).Trim,
                                  selectedData.Symbols(15).Trim)

            'スイッチ取付け方式加算価格キー
            If selectedData.Symbols(17).IndexOf("Q") < 0 Then
                'S1
                Call subSCMSwitchJoint(selectedData,
                                       strOpRefKataban,
                                       decOpAmount,
                                       strPriceDiv,
                                       selectedData.Symbols(7).Trim,
                                       selectedData.Symbols(16).Trim,
                                       selectedData.Symbols(9).Trim,
                                       selectedData.Symbols(3).Trim,
                                       selectedData.Symbols(6).Trim)
                'S2
                Call subSCMSwitchJoint(selectedData,
                                       strOpRefKataban,
                                       decOpAmount,
                                       strPriceDiv,
                                       selectedData.Symbols(13).Trim,
                                       selectedData.Symbols(16).Trim,
                                       selectedData.Symbols(15).Trim,
                                       selectedData.Symbols(3).Trim,
                                       selectedData.Symbols(12).Trim)
            End If

            'オプション加算価格キー
            Call subSCMOption(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(17).Trim,
                              selectedData.Symbols(3).Trim,
                              strSelStrokeS1,
                              selectedData.Symbols(7).Trim,
                              strSelStrokeS2,
                              selectedData.Symbols(13).Trim)

            Select Case selectedData.Series.key_kataban
                '食品製造工程向け商品
                Case "G"
                    'オプション加算価格キー
                    Call subSCMOption(selectedData,
                                      strOpRefKataban,
                                      decOpAmount,
                                      strPriceDiv,
                                      selectedData.Symbols(18).Trim,
                                      selectedData.Symbols(3).Trim,
                                      strSelStrokeS1,
                                      selectedData.Symbols(7).Trim,
                                      strSelStrokeS2,
                                      selectedData.Symbols(13).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(19).Trim,
                                        selectedData.Symbols(3).Trim,
                                        String.Empty)

                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(18).Trim,
                                        selectedData.Symbols(3).Trim,
                                        selectedData.Symbols(19).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(19).Trim,
                                        selectedData.Symbols(3).Trim)
            End Select




        Catch ex As Exception

            Throw ex

        End Try


    End Sub

    '************************************************************************************
    '*  ProgramID  ：subHighLoadBase
    '*  Program名  ：両ロッドベース
    '************************************************************************************
    Private Sub subHighLoadBase(selectedData As SelectedInfo,
                                ByRef strOpRefKataban() As String,
                                ByRef decOpAmount() As Decimal,
                                Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'ストローク設定
            strSelStrokeS1(0) = selectedData.Symbols(6).Trim
            strSelStrokeS1(1) = KatabanUtility.GetStrokeSize(selectedData,
                                                    CInt(selectedData.Symbols(3).Trim),
                                                    CInt(selectedData.Symbols(6).Trim))

            '基本価格キー
            If selectedData.Symbols(5).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-D-" &
                                                           selectedData.Symbols(3).Trim & "D-" &
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-BASE-D-" &
                                                           selectedData.Symbols(3).Trim & "B-" &
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(selectedData,
                                strOpRefKataban,
                                decOpAmount,
                                strPriceDiv,
                                selectedData.Symbols(1).Trim,
                                selectedData.Symbols(3).Trim,
                                strSelStrokeS1(1))

            '支持形式加算価格キー
            Call subSCMSupport(selectedData,
                               strOpRefKataban,
                               decOpAmount,
                               strPriceDiv,
                               selectedData.Symbols(2).Trim,
                               selectedData.Symbols(3).Trim)

            'スイッチ加算価格キー
            Call subSCMSwitch(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(8).Trim,
                              selectedData.Symbols(10).Trim)
            Call subSCMSwitchLead(selectedData,
                                  strOpRefKataban,
                                  decOpAmount,
                                  strPriceDiv,
                                  selectedData.Symbols(8).Trim,
                                  selectedData.Symbols(9).Trim,
                                  selectedData.Symbols(10).Trim)

            'スイッチ取付け方式加算価格キー
            Call subSCMSwitchJoint(selectedData,
                                   strOpRefKataban,
                                   decOpAmount,
                                   strPriceDiv,
                                   selectedData.Symbols(8).Trim,
                                   selectedData.Symbols(11).Trim,
                                   selectedData.Symbols(10).Trim,
                                   selectedData.Symbols(3).Trim,
                                   selectedData.Symbols(6).Trim)

            'オプション加算価格キー
            Call subSCMOption(selectedData,
                              strOpRefKataban,
                              decOpAmount,
                              strPriceDiv,
                              selectedData.Symbols(12).Trim,
                              selectedData.Symbols(3).Trim,
                              strSelStrokeS1,
                              selectedData.Symbols(8).Trim)


            Select Case selectedData.Series.key_kataban
                '食品製造工程向け商品
                Case "H"
                    'オプション加算価格キー
                    Call subSCMOption(selectedData,
                                      strOpRefKataban,
                                      decOpAmount,
                                      strPriceDiv,
                                      selectedData.Symbols(13).Trim,
                                      selectedData.Symbols(3).Trim,
                                      strSelStrokeS1,
                                      selectedData.Symbols(8).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(14).Trim,
                                        selectedData.Symbols(3).Trim,
                                        String.Empty)

                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(13).Trim,
                                        selectedData.Symbols(3).Trim,
                                        selectedData.Symbols(14).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(selectedData,
                                        strOpRefKataban,
                                        decOpAmount,
                                        strPriceDiv,
                                        selectedData.Symbols(14).Trim,
                                        selectedData.Symbols(3).Trim)
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　バリエーションによる加算を算出する
    '************************************************************************************
    Private Sub subSCMVariation(selectedData As SelectedInfo,
                                ByRef strOpRefKataban() As String,
                                ByRef decOpAmount() As Decimal,
                                ByRef strPriceDiv() As String,
                                ByVal strVariation As String,
                                ByVal strBoreSize As String,
                                ByVal strStrokeS1 As String,
                                Optional ByVal strStrokeS2 As String = "")

        Try

            'バリエーション「X」
            If strVariation.IndexOf("X") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-X-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「Y」
            If strVariation.IndexOf("Y") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-Y-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「W4」
            If strVariation.IndexOf("W4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-W4-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「P」
            If strVariation.IndexOf("P") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-P-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「R」
            If strVariation.IndexOf("R") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-R-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「Q」
            If strVariation.IndexOf("Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-Q-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「M」
            If strVariation.IndexOf("M") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-M-" &
                                                           strBoreSize & MyControlChars.Hyphen & strStrokeS1
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

                If strStrokeS2.Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-M-" &
                                                               strBoreSize & MyControlChars.Hyphen & strStrokeS2
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

            'バリエーション「H」
            If strVariation.IndexOf("H") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-H-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「T」
            If strVariation.IndexOf("T") >= 0 And
               strVariation.IndexOf("T1") < 0 And
               strVariation.IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-T-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「T1」
            If strVariation.IndexOf("T1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-T1-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「T2」
            If strVariation.IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-T2-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「O」
            If strVariation.IndexOf("O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-O-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「U」
            If strVariation.IndexOf("U") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-U-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「G」
            If strVariation.IndexOf("G") >= 0 And
               strVariation.IndexOf("G1") < 0 And
               strVariation.IndexOf("G2") < 0 And
               strVariation.IndexOf("G3") < 0 And
               strVariation.IndexOf("G4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「G1」
            If strVariation.IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G1-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「G2」
            If strVariation.IndexOf("G2") >= 0 Then
                'S1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G2-" &
                                                           strBoreSize & MyControlChars.Hyphen & strStrokeS1
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

                'S2
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G2-" &
                                                           strBoreSize & MyControlChars.Hyphen & strStrokeS2
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「G3」
            If strVariation.IndexOf("G3") >= 0 Then
                'S1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G3-" &
                                                           strBoreSize & MyControlChars.Hyphen & strStrokeS1
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

                'S2
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G3-" &
                                                           strBoreSize & MyControlChars.Hyphen & strStrokeS2
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「G4」
            If strVariation.IndexOf("G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-G4-" &
                                                           strBoreSize
                Select Case selectedData.Series.key_kataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「F」
            If strVariation.IndexOf("F") >= 0 Then
                'S1
                Select Case True
                    Case CInt(strStrokeS1) <= 50
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                   strBoreSize & "-10-50"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case CInt(strStrokeS1) >= 51 And CInt(strStrokeS1) <= 300
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                   strBoreSize & "-51-300"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case CInt(strStrokeS1) >= 301
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                   strBoreSize & "-301-500"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select

                'S2
                If strStrokeS2.Trim <> "" Then
                    Select Case True
                        Case CInt(strStrokeS2) <= 50
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                       strBoreSize & "-10-50"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case CInt(strStrokeS2) >= 51 And CInt(strStrokeS2) <= 300
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                       strBoreSize & "-51-300"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case CInt(strStrokeS2) >= 301
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-F-" &
                                                                       strBoreSize & "-301-500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select
                End If
            End If

            'バリエーション「B」
            If strVariation.IndexOf("B") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-B-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーション「W」
            If strVariation.IndexOf("W") >= 0 And
               strVariation.IndexOf("W4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-VAR-W-" &
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　支持形式による加算を算出する
    '************************************************************************************
    Private Sub subSCMSupport(selectedData As SelectedInfo,
                              ByRef strOpRefKataban() As String,
                              ByRef decOpAmount() As Decimal,
                              ByRef strPriceDiv() As String,
                              ByVal strSupport As String,
                              ByVal strBoreSize As String)

        Try

            If strSupport.Trim <> "00" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SUPPORT-" &
                                                           strSupport.Trim & MyControlChars.Hyphen & strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If strSupport.Trim = "LD" Then
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　スイッチによる加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitch(selectedData As SelectedInfo,
                             ByRef strOpRefKataban() As String,
                             ByRef decOpAmount() As Decimal,
                             ByRef strPriceDiv() As String,
                             ByVal strSwitch As String,
                             ByVal strSwitchNum As String)

        Try

            If strSwitch.Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-" &
                                                           strSwitch.Trim
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　リード線の長さによる加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitchLead(selectedData As SelectedInfo,
                                 ByRef strOpRefKataban() As String,
                                 ByRef decOpAmount() As Decimal,
                                 ByRef strPriceDiv() As String,
                                 ByVal strSwitch As String,
                                 ByVal strSwitchLead As String,
                                 ByVal strSwitchNum As String)

        Try

            If strSwitch.Trim <> "" Then
                If strSwitchLead.Trim <> "" Then
                    Select Case strSwitch.Trim
                        Case "T2H", "T2V", "T2YH", "T2YV", "T3H",
                             "T3V", "T3YH", "T3YV", "T0H", "T0V",
                             "T5H", "T5V", "T1H", "T1V", "T8H", "T8V",
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(1)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH",
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(2)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(3)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(4)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(5)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SWLW(6)-" &
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                    End Select
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　スイッチ取付け方式による加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitchJoint(selectedData As SelectedInfo,
                                  ByRef strOpRefKataban() As String,
                                  ByRef decOpAmount() As Decimal,
                                  ByRef strPriceDiv() As String,
                                  ByVal strSwitch As String,
                                  ByVal strSwitchJoint As String,
                                  ByVal strSwitchNum As String,
                                  ByVal strBoreSize As String,
                                  ByVal strStroke As String)

        Try

            If strSwitch.Trim <> "" Then
                If strSwitchJoint.Trim = "" Then
                    Select Case True
                        Case CInt(strStroke) <= 300
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-JOINT-" &
                                                                       strBoreSize.Trim & "-5-300"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case CInt(strStroke) >= 301 And CInt(strStroke) <= 500
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-JOINT-" &
                                                                       strBoreSize.Trim & "-301-500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case CInt(strStroke) >= 501 And CInt(strStroke) <= 1000
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-JOINT-" &
                                                                       strBoreSize.Trim & "-501-1000"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case CInt(strStroke) >= 1001
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-JOINT-" &
                                                                       strBoreSize.Trim & "-1001-1500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-JOINT-" &
                                                               strSwitchJoint.Trim & MyControlChars.Hyphen & strBoreSize.Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
                End If
                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                'P4加算
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchNum)
                End If

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　オプションによる加算を算出する
    '************************************************************************************
    Private Sub subSCMOption(selectedData As SelectedInfo,
                             ByRef strOpRefKataban() As String,
                             ByRef decOpAmount() As Decimal,
                             ByRef strPriceDiv() As String,
                             ByVal strOptionVar As String,
                             ByVal strBoreSize As String,
                             ByVal strStrokeS1() As String,
                             ByVal strSwitchS1 As String,
                             Optional ByVal strStrokeS2() As String = Nothing,
                             Optional ByVal strSwitchS2 As String = "")

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            'オプション分解
            strOpArray = Split(strOptionVar, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "Q"
                        Select Case selectedData.Series.key_kataban.Trim
                            Case "B"
                                Select Case True
                                    Case strSwitchS1 <> "" And strSwitchS2 = ""
                                        'S1
                                        Select Case True
                                            Case CInt(strStrokeS1(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                        End Select
                                    Case strSwitchS1 = "" And strSwitchS2 <> ""
                                        'S2
                                        Select Case True
                                            Case CInt(strStrokeS2(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                        End Select
                                    Case Else
                                        'S1
                                        Select Case True
                                            Case CInt(strStrokeS1(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                        End Select

                                        'S2
                                        Select Case True
                                            Case CInt(strStrokeS2(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                           strBoreSize.Trim & "-1001-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If
                                        End Select
                                End Select
                            Case Else
                                'スイッチ選択無しの時のみ加算
                                If strSwitchS1 = "" Then
                                    Select Case True
                                        Case CInt(strStrokeS1(0)) <= 300
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                       strBoreSize.Trim & "-10-300"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS1(0)) <= 500
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                       strBoreSize.Trim & "-301-500"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS1(0)) <= 1000
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                       strBoreSize.Trim & "-501-1000"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 1001
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-Q-" &
                                                                                       strBoreSize.Trim & "-1001-1500"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select

                                    Select Case selectedData.Series.key_kataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                End If
                        End Select
                    Case "J", "K", "L"
                        'S1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS1(1)
                        Select Case selectedData.Series.key_kataban.Trim
                            Case "D", "H"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If

                        'S2
                        If strStrokeS2 IsNot Nothing Then
                            If strStrokeS2(1) <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS2(1)
                                Select Case selectedData.Series.key_kataban.Trim
                                    Case "D", "H"
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Case Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                            End If
                        End If
                    Case "M"
                        'S1
                        If selectedData.Symbols(1).IndexOf("M") >= 0 And
                           strBoreSize.Trim = "32" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP*-" &
                                                                       strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS1(1)
                            Select Case selectedData.Series.key_kataban.Trim
                                Case "D", "H"
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                       strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS1(1)
                            Select Case selectedData.Series.key_kataban.Trim
                                Case "D", "H"
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        End If

                        'S2
                        If strStrokeS2 IsNot Nothing Then
                            If strStrokeS2(1) <> "" Then
                                If selectedData.Symbols(1).IndexOf("M") >= 0 And
                                   strBoreSize.Trim = "32" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP*-" &
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS2(1)
                                    Select Case selectedData.Series.key_kataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim & MyControlChars.Hyphen & strStrokeS2(1)
                                    Select Case selectedData.Series.key_kataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                    End If
                                End If
                            End If
                        End If
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim
                        Select Case selectedData.Series.key_kataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case Else
                                Dim fullKataban = PriceManager.GetFullKataban(selectedData)
                                If fullKataban.IndexOf("N13-N11") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "FP1"
                        '食品製造工程向け商品
                        Select Case selectedData.Symbols(1).Trim
                            Case "W4", "B", "W"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim &
                                                                            MyControlChars.Hyphen & selectedData.Symbols(1).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize.Trim
                        Select Case selectedData.Series.key_kataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　付属品による加算を算出する
    '************************************************************************************
    Private Sub subSCMAccesary(selectedData As SelectedInfo,
                               ByRef strOpRefKataban() As String,
                               ByRef decOpAmount() As Decimal,
                               ByRef strPriceDiv() As String,
                               ByVal strAccesary As String,
                               ByVal strBoreSize As String,
                               ByVal strTipOfRod As String)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            strOpArray = Split(strAccesary, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-ACC-" &
                                                                   Left(strOpArray(intLoopCnt).Trim, 1) & MyControlChars.Hyphen & strBoreSize
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-ACC-" &
                                                                   Right(strOpArray(intLoopCnt).Trim, 1) & MyControlChars.Hyphen & strBoreSize
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "I", "Y"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-ACC-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                        Select Case selectedData.Series.key_kataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B", "G"
                                If selectedData.Symbols(1).IndexOf("B") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "H"
                                If strTipOfRod.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-ACC-" &
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                        Select Case selectedData.Series.key_kataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B", "G"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "D", "H"
                                If strTipOfRod.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　ロッド先端オ－ダ－メイドによる加算を算出する
    '************************************************************************************
    Private Sub subSCMTipOfRod(selectedData As SelectedInfo,
                               ByRef strOpRefKataban() As String,
                               ByRef decOpAmount() As Decimal,
                               ByRef strPriceDiv() As String,
                               ByVal strTipOfRod As String,
                               ByVal strBoreSize As String)

        Try

            If strTipOfRod.Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-TIP-OF-ROD-" &
                                                           strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
