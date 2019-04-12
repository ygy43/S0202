'************************************************************************************
'*  ProgramID  ：KHPriceC8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*
'*  概要       ：残圧排出弁／圧力スイッチ　白色シリーズ
'*             ：Ｖ１０００－Ｗ／Ｖ３０００－Ｗ／Ｖ３０１０－Ｗ
'*　　　　　　 ：Ｐ４０００－Ｗ／Ｐ４１００－Ｗ／Ｐ８１００－Ｗ
'*　　　　　　 ：ＡＰＳ－Ｗ
'*　　　　　　 ：ＦＸ
'*
'*  更新履歴   ：                       更新日：2008/01/22   更新者：NII A.Takahashi
'*               ・APS-W/V1000-W/V3000-W/V3010-W/P8100-Wを追加したため、単価見積りロジック変更
'*  ・受付No：RM0907070  二次電池対応機器　V3010
'*                                      更新日：2009/08/25   更新者：Y.Miura
'*  ・受付No：RM1001045  二次電池対応機器追加 P4100 
'*                                      更新日：2010/02/24   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceC8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                'RM1306001 2013/06/04 追加
                Case "FX1004", "FX1011", "FX1037"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If selectedData.Symbols(2).Trim = "" Then
                        If Left(selectedData.Symbols(4), 1) = "F" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "F"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim

                        End If
                    Else
                        If Left(selectedData.Symbols(4), 1) = "F" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "F"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim

                        End If
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    Select Case selectedData.Symbols(5)
                        Case "Z"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-Z"

                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "M"
                            If Left(selectedData.Symbols(4), 1) = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-CM"

                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-M"

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "M1"
                            If Left(selectedData.Symbols(4), 1) = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-CM1"

                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-M1"

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                    End Select

                    'RM1807033_FP機種追加、二次電池機種価格不具合修正
                    Select Case selectedData.Series.key_kataban
                        Case "4", "F"

                            '二次電池、食品製造加算価格キー
                            If selectedData.Symbols(7).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'アタッチメント加算価格キー
                            If selectedData.Symbols(8).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                        Case Else

                            'アタッチメント加算価格キー
                            If selectedData.Symbols(7).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                    End Select

                Case "P4000"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    'RM1612036 T8オプション追加による変更  2016/12/19 追加 松原
                    If selectedData.Symbols(3).IndexOf("T8") >= 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim

                    ElseIf selectedData.Symbols(3).IndexOf("T") >= 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "T"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "T"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    'アタッチメント加算価格キー
                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "V1000", "V3000", "V3010", "V6010"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アタッチメント加算価格キー
                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '食品製造工程向け商品
                    Select Case selectedData.Symbols(4)
                        Case "FP1"
                            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '食品製造工程向け商品
                    Select Case selectedData.Series.series_kataban
                        Case "V1000", "V3000"
                            If selectedData.Series.key_kataban = "F" Then
                                strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                                For intLoopCnt = 0 To strOpArray.Length - 1
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case ""
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Next
                            End If
                        Case Else
                    End Select

                    'RM0907070 2009/08/25 Y.Miura　二次電池対応
                    strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    'RM1001045 2010/02/24 Y.Miura 二次電池追加
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    '二次電池用
                    If selectedData.Series.key_kataban = "X" Then
                        '二次電池加算価格キー
                        If selectedData.Symbols(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                    End If
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                    'Case "P1100", "P4100", "P8100"
                Case "P1100", "P8100"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Series.key_kataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    ''添付品価格キー
                    'If selectedData.Symbols(4).Trim <> "" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 1) & CdCst.Sign.Asterisk & _
                    '                                               Mid(selectedData.Series.series_kataban.Trim, 3, 5) & _
                    '                                               selectedData.Symbols(4).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                    Select Case selectedData.Series.key_kataban
                        Case "F"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'RM1305XXX 2013/05/02
                            '添付品価格キー
                            strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Next

                            'リード線価格キー
                            If selectedData.Symbols(6).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                            'RM1001045 2010/02/24 Y.Miura 二次電池追加

                        Case Else
                            'RM1305XXX 2013/05/02
                            '添付品価格キー
                            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Next

                            'リード線価格キー
                            If selectedData.Symbols(5).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                            'RM1001045 2010/02/24 Y.Miura 二次電池追加

                    End Select


                Case "P4100"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Series.key_kataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算（二次電池）
                    Select Case selectedData.Symbols(4).Trim
                        Case "P4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算（食品製造向け商品）
                        Case "FP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1


                    End Select
                    ''添付品価格キー
                    'If selectedData.Symbols(5).Trim <> "" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 1) & CdCst.Sign.Asterisk & _
                    '                                               Mid(selectedData.Series.series_kataban.Trim, 3, 5) & _
                    '                                               selectedData.Symbols(5).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                    'RM1305XXX 2013/05/02
                    '添付品価格キー
                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next

                    'リード線価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "APS"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                    '                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                    '                                           selectedData.Series.key_kataban.Trim
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                    decOpAmount(UBound(decOpAmount)) = 1

                    'リード線価格キー
                    If selectedData.Symbols(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '添付品加算価格キー
                    strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    '二次電池用
                    If selectedData.Series.key_kataban = "Z" Then
                        '二次電池加算価格キー
                        If selectedData.Symbols(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                    End If
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                Case "P1100-UN", "P4100-UN"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 8) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & "-W"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'リード線価格キー
                    If selectedData.Symbols(5).Trim <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
