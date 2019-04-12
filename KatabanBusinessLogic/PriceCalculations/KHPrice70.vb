'************************************************************************************
'*  ProgramID  ：KHPrice70
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド電磁弁付バルブブロック単品
'*             ：Ｎ３ＧＡ１・２／Ｎ３ＧＢ１・２／Ｎ４ＧＡ１・２／Ｎ４ＧＢ１・２
'*             ：Ｎ３ＧＤ１・２／Ｎ３ＧＥ１・２／Ｎ４ＧＤ１・２／Ｎ４ＧＥ１・２
'*
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　N3GA1/N3GA2/N4GA1/N4GA2/N3GB1/N3GB2/N4GB1/N4GB2にオプションボックス追加
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice70

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intQuantity As Integer

        Dim strKiriIchikbn As String = ""   '切換位置区分
        Dim strSosakbn As String = ""       '操作区分
        Dim strKokei As String = ""         '接続口径
        Dim strCable As String = ""         'ケーブル長さ
        Dim strTanshi As String = ""        '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strSyudoSochi As String = ""    '手動装置
        Dim strDensen As String = ""        '電線接続
        Dim strOption As String = ""        'オプション
        Dim strDenatsu As String = ""       '電圧
        Dim strCleanShiyo As String = ""    'クリーン仕様
        Dim strOptionFP1 As String = ""     '食品製造工程向け商品 RM1610013

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case selectedData.Series.key_kataban.Trim
                'Case "R"
                Case "R", "S" 'RM1610013
                    If selectedData.Series.series_kataban.Contains("GD") Or _
                       selectedData.Series.series_kataban.Contains("GE") Then
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strCable = selectedData.Symbols(6).Trim                'ケーブル長さ
                        strTanshi = selectedData.Symbols(7).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                        strOption = selectedData.Symbols(8).Trim              'オプション
                        strDenatsu = selectedData.Symbols(9).Trim             '電圧
                        If UBound(selectedData.Symbols.ToArray()) >= 10 Then
                            strCleanShiyo = selectedData.Symbols(10).Trim          'クリーン仕様
                        End If
                    Else
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strCable = selectedData.Symbols(6).Trim                'ケーブル長さ
                        strTanshi = selectedData.Symbols(7).Trim               '端子･ｺﾈｸﾀﾋﾟﾝ配列
                        strSyudoSochi = selectedData.Symbols(8).Trim          '手動装置
                        strOption = selectedData.Symbols(9).Trim              'オプション
                        strDenatsu = selectedData.Symbols(10).Trim             '電圧
                        If UBound(selectedData.Symbols.ToArray()) >= 11 Then
                            strCleanShiyo = selectedData.Symbols(11).Trim          'クリーン仕様
                        End If
                        If UBound(selectedData.Symbols.ToArray()) >= 13 Then              'RM1610013 Start
                            strOptionFP1 = selectedData.Symbols(13).Trim        '食品製造工程向け 
                        End If                                                                   'RM1610013 End
                    End If
                Case Else
                    strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                    strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                    strKokei = selectedData.Symbols(3).Trim               '接続口径
                    strDensen = selectedData.Symbols(4).Trim              '電線接続
                    strCable = selectedData.Symbols(5).Trim                'ケーブル長さ
                    strSyudoSochi = selectedData.Symbols(6).Trim          '手動装置
                    strOption = selectedData.Symbols(7).Trim              'オプション
                    strDenatsu = selectedData.Symbols(8).Trim             '電圧
                    strCleanShiyo = selectedData.Symbols(9).Trim          'クリーン仕様
            End Select

            '数量設定
            Select Case strKiriIchikbn
                Case "1"
                    intQuantity = 1
                Case "11"
                    intQuantity = 1
                Case "66", "67", "76", "77"
                    intQuantity = 2
                Case "2"
                    intQuantity = 2
                Case "3"
                    intQuantity = 2
                Case "4"
                    intQuantity = 2
                Case "5"
                    intQuantity = 2
            End Select

            Select Case Left(selectedData.Series.series_kataban.Trim, 4)
                Case "N4GE", "N3GE"
                    'If selectedData.Series.key_kataban.Trim = "R" Then
                    If selectedData.Series.key_kataban.Trim = "R" Or selectedData.Series.key_kataban.Trim = "S" Then 'RM1610013
                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R" & MyControlChars.Hyphen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        If (Left(selectedData.Series.series_kataban.Trim, 4) = "N4GE" And strKiriIchikbn = "1") Or _
                           (Left(selectedData.Series.series_kataban.Trim, 4) = "N3GE" And strKiriIchikbn = "1") Then
                            If Left(strKokei, 2) = "CL" Then
                                If Left(strDensen, 1) = "A" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-A2N" & _
                                                                               strCable & "-CL"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-CL"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                If Left(strDensen, 1) = "A" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-A2N" & _
                                                                               strCable & "-C"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                               strKiriIchikbn & _
                                                                               strSosakbn & "-C"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            End If
                        Else
                            If Left(strDensen, 1) = "A" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                           strKiriIchikbn & _
                                                                           strSosakbn & "-A2N" & _
                                                                           strCable
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & _
                                                                           strKiriIchikbn & _
                                                                           strSosakbn
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If
                Case Else
                    'If selectedData.Series.key_kataban.Trim = "R" Then
                    If selectedData.Series.key_kataban.Trim = "R" Or selectedData.Series.key_kataban.Trim = "S" Then 'RM1610013
                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R" & MyControlChars.Hyphen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else

                        '基本価格キー
                        If strDensen = "A2N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn & MyControlChars.Hyphen & _
                                                                       strDensen & _
                                                                       strCable
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       strKiriIchikbn & _
                                                                       strSosakbn
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            'クリーン仕様加算価格キー
            If strCleanShiyo <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           strKiriIchikbn & _
                                                           strSosakbn & MyControlChars.Hyphen & _
                                                           strCleanShiyo
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '接続口径(継手エルボ)加算価格キー
            'If selectedData.Series.key_kataban.Trim = "R" Then
            If selectedData.Series.key_kataban.Trim = "R" Or selectedData.Series.key_kataban.Trim = "S" Then 'RM1610013
                Select Case True
                    Case Left(strKokei, 2) = "CL" Or _
                         Left(strKokei, 2) = "CD" Or _
                         Left(strKokei, 2) = "CF" Or _
                         Left(strKokei, 3) = "C18" Or _
                         Right(strKokei, 1) = "N" Or _
                         Right(strKokei, 1) = "G"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        If InStr(selectedData.Series.series_kataban.Trim, "3G") <> 0 And _
                           (InStr(strKiriIchikbn, "1") <> 0 Or _
                            InStr(strKiriIchikbn, "11") <> 0) Then

                            If strDensen.Trim = "A2N" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & _
                                           strKokei & MyControlChars.Hyphen & "S-A2N"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & _
                                           strKokei & MyControlChars.Hyphen & "S"
                            End If

                        Else
                            If Right(Left(selectedData.Series.series_kataban.Trim, 4), 1) = "B" Then
                                If strDensen.Trim = "A2N" Then
                                    If strCable = "" Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei & "-A2N"
                                    ElseIf strKiriIchikbn = "1" Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei & "-A2N21"
                                    Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei & "-A2N2"
                                    End If
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei
                                End If
                            Else
                                If strDensen.Trim = "A2N" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei & "-A2N"
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-" & strKokei
                                End If
                            End If
                        End If
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Else

                If Left(selectedData.Series.series_kataban.Trim, 4) <> "N4GE" Then

                    Select Case True
                        Case Left(strKokei, 2) = "CL" Or _
                             Left(strKokei, 2) = "CD" Or _
                             Left(strKokei, 2) = "CF" Or _
                             Left(strKokei, 3) = "C18"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If InStr(selectedData.Series.series_kataban.Trim, "3G") <> 0 And _
                               (InStr(strKiriIchikbn, "1") <> 0 Or _
                                InStr(strKiriIchikbn, "11") <> 0) Then
                                If InStr(strKokei, "N") <> 0 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               Left(strKokei, InStr(strKokei, "N") - 1) & _
                                                                               MyControlChars.Hyphen & "S"
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               strKokei & MyControlChars.Hyphen & "S"
                                End If
                            Else
                                If InStr(strKokei, "N") <> 0 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & Left(strKokei, InStr(strKokei, "N") - 1)
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-" & strKokei
                                End If
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                End If
            End If
            '電線接続加算価格キー
            If strDensen <> "" Then
                If strDensen <> "A2N" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               strDensen
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                End If
            End If

            '端子・ｺﾈｸﾀﾋﾟﾝ配列
            If strTanshi <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strTanshi
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション　加算価格キー
            strOpArray = Split(strOption, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "F"
                        Select Case strKiriIchikbn
                            Case "66", "67", "76", "77"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & "DUAL"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "S", "E"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        If strKiriIchikbn <> "1" And strKiriIchikbn <> "11" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "H"
                        'If Not selectedData.Series.key_kataban.Trim = "R" Then
                        If Not selectedData.Series.key_kataban.Trim = "R" And Not selectedData.Series.key_kataban.Trim = "S" Then 'RM1610013
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'If selectedData.Series.key_kataban.Trim = "R" Then
            If selectedData.Series.key_kataban.Trim = "R" Or selectedData.Series.key_kataban.Trim = "S" Then 'RM1610013
                If Not strOption.Contains("H") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-H"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '電圧
            Select Case selectedData.Series.key_kataban
                'Case "R"
                Case "R", "S" 'RM1610013
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select

            '食品製造工程向け商品 RM1610013
            If selectedData.Series.key_kataban.Trim = "S" Then
                If strOptionFP1.Contains("FP1") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
