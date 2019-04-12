'************************************************************************************
'*  ProgramID  ：KHPriceL3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：リニアガイドハンド　ＬＨＡ
'*             ：ゴムカバー付リニアガイドハンド　ＬＨＡＧ
'*             ：薄形広角ハンド　ＨＭＤ
'*             ：薄形チャック　ＣＫＳ
'*  ・受付No：RM0908030  二次電池対応機器　HMD
'*                                      更新日：2009/10/19   更新者：Y.Miura
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceL3

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionO As Boolean = False
        Dim bolOptionC As Boolean = False
        Dim bolOptionF As Boolean = False
        Dim bolOptionT As Boolean = False

        Dim bolOptionW As Boolean = False
        Dim bolOptionY1 As Boolean = False
        Dim bolOptionY2 As Boolean = False
        Dim strNumber As String = Nothing
        Dim strOptionP4 As String = String.Empty            'RM0908030 2009/10/19 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean                            'RM0908030 2009/10/19 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)            'RM0908030 2009/10/28 Y.Miura


            'C5チェック                      'RM0908030 2009/10/19 Y.Miura
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData)

            Select Case selectedData.Series.series_kataban.Trim
                '201501月次更新
                Case "LML"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スクレーパ付加算価格
                    If selectedData.Symbols(1).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    Select Case selectedData.Series.key_kataban
                        Case "1"
                            'ポート加算価格
                            If selectedData.Symbols(5).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '手動解除加算価格
                            If selectedData.Symbols(6).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "2"
                            'ポート加算価格
                            If selectedData.Symbols(4).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '手動解除加算価格
                            If selectedData.Symbols(5).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select

                Case "LHA"
                    'RM1310067 2013/10/23 追加
                    If selectedData.Series.series_kataban.Trim = "L" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'スイッチ加算価格キー
                        If selectedData.Symbols(3).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)

                            'リード線長さ加算価格キー
                            If selectedData.Symbols(4).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                            End If
                        End If
                    Else
                        '基本価格キー
                        If selectedData.Symbols(2).Trim = "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/28 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        'スイッチ加算価格キー
                        If selectedData.Symbols(3).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)

                            'リード線長さ加算価格キー
                            If selectedData.Symbols(4).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                            End If
                        End If
                    End If
                Case "LHAG"
                    'オプション判定
                    strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "O"
                                bolOptionO = True
                            Case "C"
                                bolOptionC = True
                            Case "F"
                                bolOptionF = True
                            Case "T"
                                bolOptionT = True
                        End Select
                    Next

                    'RM0908030 2009/10/19 Y.Miura
                    'P4判定
                    If selectedData.Symbols.Count >= 7 Then
                        strOptionP4 = selectedData.Symbols(6)
                    End If

                    '基本価格キー
                    Select Case True
                        Case bolOptionO = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "O"
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/28 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case bolOptionC = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "C"
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/28 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/28 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select

                    'スイッチ加算価格キー
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                            selectedData.Symbols(4).Trim()
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                        End If

                        'RM0908030 2009/10/19 Y.Miura
                        'SW二次電池加算
                        If strOptionP4 <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                        End If

                    End If

                    'オプション加算価格キー
                    Select Case True
                        Case bolOptionF = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "F"
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/30 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case bolOptionT = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "T"
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0908030 2009/10/30 Y.Miura
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select

                    'RM0908030 2009/10/19 Y.Miura
                    '二次電池対応
                    If strOptionP4 <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOptionP4 & MyControlChars.Hyphen & selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "HMD"

                    'RM0908030 2009/10/19 Y.Miura
                    'P4判定
                    If selectedData.Symbols.Count >= 6 Then
                        strOptionP4 = selectedData.Symbols(5)
                    End If

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/28 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    'スイッチ加算価格キー
                    If selectedData.Symbols(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(4).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(3).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(4).Trim)
                        End If

                        'RM0908030 2009/10/19 Y.Miura
                        'SW二次電池加算
                        If strOptionP4 <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(4).Trim)
                        End If

                    End If

                    'RM0908030 2009/10/19 Y.Miura
                    '二次電池対応
                    If strOptionP4 <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & "-OP-" & strOptionP4
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "CKS"
                    'オプション判定
                    strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "W"
                                bolOptionW = True
                            Case "Y1"
                                bolOptionY1 = True
                            Case "Y2"
                                bolOptionY2 = True
                        End Select
                    Next

                    '基本価格キー
                    Select Case True
                        Case bolOptionW = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "W"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'スイッチ加算価格キー
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(5).Trim)
                        End If
                    End If

                    '小爪加算価格キー
                    Select Case selectedData.Symbols(1).Trim
                        Case "08CS"
                            strNumber = "910"
                        Case "12CS"
                            strNumber = "920"
                        Case "16CS"
                            strNumber = "930"
                        Case "20CS"
                            strNumber = "940"
                        Case "25CS"
                            strNumber = "950"
                        Case "32CS"
                            strNumber = "960"
                    End Select

                    Select Case True
                        Case bolOptionY1 = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "Y1" & MyControlChars.Hyphen & strNumber
                            decOpAmount(UBound(decOpAmount)) = 3
                        Case bolOptionY2 = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "Y2" & MyControlChars.Hyphen & strNumber
                            decOpAmount(UBound(decOpAmount)) = 3
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
