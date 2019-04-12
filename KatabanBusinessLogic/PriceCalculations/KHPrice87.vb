'************************************************************************************
'*  ProgramID  ：KHPrice87
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーマウントシリンダ　ＳＭＤ２／ＳＭＤ２－Ｌ
'*
'*  ・受付No：RM0908030  二次電池対応機器　
'*                                      更新日：2009/09/04   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'*  ・受付No：RM1112XXX  SMGシリーズ追加　
'*                                      更新日：2011/12/22   更新者：Y.Tachi
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice87

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer
        Dim bolOptionP4 As Boolean = False      'RM0908030 2009/09/04 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応

            'RM0908030 2009/09/04 Y.Miura　二次電池対応
            Select Case selectedData.Symbols(8).Trim
                Case "P4", "P40"
                    bolOptionP4 = True
            End Select

            'C5チェック
            'RM1001043 2010/02/22 Y.Miura 廃止
            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
            bolC5Flag = False

            Select Case selectedData.Series.series_kataban.Trim
                Case "SMG"
                    Select Case selectedData.Series.key_kataban.Trim
                        Case "2"
                            bolC5Flag = True

                            '基本価格キー
                            If selectedData.Symbols(1).Trim <> "" Then
                                If selectedData.Symbols(1).Trim = "M" Then
                                    If selectedData.Symbols(7).Trim = "35" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "40"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    ElseIf selectedData.Symbols(7).Trim = "45" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(7).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(7).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                If selectedData.Symbols(7).Trim = "35" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "40"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    If selectedData.Symbols(7).Trim = "45" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        If selectedData.Symbols(7).Trim = "55" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            If selectedData.Symbols(7).Trim = "65" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "70"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                If selectedData.Symbols(7).Trim = "75" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                               selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "80"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    If selectedData.Symbols(7).Trim = "85" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                   selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "90"
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        If selectedData.Symbols(7).Trim = "95" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "100"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(7).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                End If

                                    '支持形式加算
                                    If selectedData.Symbols(1).Trim <> "M" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(5).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    '微速F加算
                                    If selectedData.Symbols(3).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        Select Case selectedData.Symbols(5).Trim
                                            Case "6", "10", "16"
                                                If selectedData.Symbols(7).Trim <= 15 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                        selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           "VAR" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                           "5"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                ElseIf selectedData.Symbols(7).Trim <= 30 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           "VAR" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                           "16"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                            Case "20", "25", "32"
                                                If selectedData.Symbols(7).Trim <= 25 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           "VAR" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                           "5"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                ElseIf selectedData.Symbols(7).Trim <= 50 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           "VAR" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                           "26"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                        End Select
                                    End If


                                    'スイッチ加算価格キー
                                    If selectedData.Symbols(2).Trim <> "" Then
                                        If selectedData.Symbols(1).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(1).Trim & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(5).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(5).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If

                                        'リード線長さ加算価格キー
                                        If selectedData.Symbols(8).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(10).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If

                                        If selectedData.Symbols(9).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(9).Trim
                                            If selectedData.Symbols(10).Trim = "D" Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                If selectedData.Symbols(10).Trim = "T" Then
                                                    decOpAmount(UBound(decOpAmount)) = 3
                                                Else
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                            End If
                                        End If
                                    End If

                        Case Else

                            Select Case selectedData.Symbols(10).Trim
                                Case "P4", "P40"
                                    bolOptionP4 = True
                            End Select

                            'スイッチがK3P＊の場合はＣ５
                            If selectedData.Symbols(7).Trim = "K3PH" Or _
                               selectedData.Symbols(7).Trim = "K3PV" Then
                                bolC5Flag = True
                            End If

                            'ねじがNN,GNの場合はＣ５
                            If selectedData.Symbols(5).Trim = "NN" Or _
                               selectedData.Symbols(5).Trim = "GN" Then
                                bolC5Flag = True
                            End If

                            ''微速Fの場合はＣ５
                            'If selectedData.Symbols(3).Trim = "F" Then
                            '    bolC5Flag = True
                            'End If

                            ''クリーン仕様P5,P51,P7,P71の場合はＣ５
                            'If selectedData.Symbols(10).Trim = "P5" Or _
                            '    selectedData.Symbols(10).Trim = "P51" Or _
                            '    selectedData.Symbols(10).Trim = "P7" Or _
                            '    selectedData.Symbols(10).Trim = "P71" Then
                            '    bolC5Flag = True
                            'End If

                            '基本価格キー
                            If selectedData.Symbols(1).Trim <> "" Then

                                '2016/12/06 問い合わせ対応（バグ修正）
                                '価格キーに使用するために、マスタよりストロークを取得
                                intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                                      CInt(selectedData.Symbols(4).Trim), _
                                                                      CInt(selectedData.Symbols(6).Trim))
                                '2016/12/06 修正End

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                                '2016/12/06 問い合わせ対応（バグ修正）
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                           intStroke.ToString

                                'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                '                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                '                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                '                                           selectedData.Symbols(6).Trim
                                '2016/12/06 修正End
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                If selectedData.Symbols(6).Trim = "35" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "40"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    If selectedData.Symbols(6).Trim = "45" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        If selectedData.Symbols(6).Trim = "55" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            If selectedData.Symbols(6).Trim = "65" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "70"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                If selectedData.Symbols(6).Trim = "75" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "80"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    If selectedData.Symbols(6).Trim = "85" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "90"
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        If selectedData.Symbols(6).Trim = "95" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "100"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                                       selectedData.Symbols(6).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            '微速F加算
                            If selectedData.Symbols(3).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                Select Case selectedData.Symbols(4).Trim
                                    Case "6", "10", "16"
                                        If selectedData.Symbols(6).Trim <= 15 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                   "VAR" & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                   "5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        ElseIf selectedData.Symbols(6).Trim <= 30 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                   "VAR" & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                   "16"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case "20", "25", "32"
                                        If selectedData.Symbols(6).Trim <= 25 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                   "VAR" & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                   "5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        ElseIf selectedData.Symbols(6).Trim <= 50 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                   "VAR" & MyControlChars.Hyphen & _
                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                   "26"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                End Select
                            End If


                            'スイッチ加算価格キー
                            If selectedData.Symbols(2).Trim <> "" Then
                                If selectedData.Symbols(1).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                                'リード線長さ加算価格キー
                                If selectedData.Symbols(7).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(9).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                                If selectedData.Symbols(8).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(8).Trim
                                    If selectedData.Symbols(9).Trim = "D" Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        If selectedData.Symbols(9).Trim = "T" Then
                                            decOpAmount(UBound(decOpAmount)) = 3
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                                End If

                                If bolOptionP4 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-P4"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(9).Trim)
                                End If
                            End If

                            'クリーン仕様加算
                            If selectedData.Symbols(10).Trim <> "" Then
                                Select Case selectedData.Symbols(10).Trim
                                    Case "P5", "P51"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                   "P5"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P7", "P71"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                   "P7"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P4"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                   "P4"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P40"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                   "P40"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                    End Select
                Case Else

                    'ストローク取得
                    intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                          CInt(selectedData.Symbols(3).Trim), _
                                                          CInt(selectedData.Symbols(4).Trim))

                    'バリエーション(微速)加算価格キー
                    Select Case selectedData.Symbols(1).Trim
                        Case "F"
                            Select Case selectedData.Symbols(3).Trim
                                Case "6", "10", "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 15
                                            'ストローク5～15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR5" & MyControlChars.Hyphen & "15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(4).Trim) >= 16 And _
                                             CInt(selectedData.Symbols(4).Trim) <= 30
                                            'ストローク16～30
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR16" & MyControlChars.Hyphen & "30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(4).Trim) >= 31
                                            'ストローク31～60
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR31" & MyControlChars.Hyphen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                                Case "20", "25", "32"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 25
                                            'ストローク5～25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR5" & MyControlChars.Hyphen & "25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(4).Trim) >= 26 And _
                                             CInt(selectedData.Symbols(4).Trim) <= 50
                                            'ストローク26～50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR26" & MyControlChars.Hyphen & "50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                        Case CInt(selectedData.Symbols(4).Trim) >= 51
                                            'ストローク51～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                                       MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR51" & MyControlChars.Hyphen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                            End If
                                    End Select
                            End Select
                    End Select

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    'マグネット内臓(L)加算価格キー
                    If Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & _
                                                                   MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    End If

                    '支持形式加算価格キー
                    If selectedData.Symbols(2).Trim = "DC" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    End If

                    'スイッチ加算価格キー
                    If selectedData.Symbols(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(6).Trim <> "" Then
                            Select Case selectedData.Symbols(5).Trim
                                Case "K0H", "K0V", "K2H", "K2V", "K3H", _
                                     "K3V", "K5H", "K5V", "K2YH", "K2YV", _
                                     "K3YH", "K3YV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                                Case "K2YFH", "K2YFV", "K3YFH", "K3YFV", "K2YMH", _
                                     "K2YMV", "K3YMH", "K3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(6).Trim & "Y"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                            End Select
                        End If

                        'RM0908030 2009/09/04 Y.Miura　二次電池対応
                        'Ｐ４加算　SW数
                        If bolOptionP4 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        End If

                    End If

                    'クリーン仕様加算価格キー
                    If selectedData.Symbols(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
