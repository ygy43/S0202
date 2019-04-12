'************************************************************************************
'*  ProgramID  ：KHPrice54
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：2009/04/03   更新者：T.Yagyu
'*
'*  概要       ：シリンダバルブ
'*             ：ＮＡＢ／ＧＮＡＢ／ＮＳＢ
'*
'*  2009/04/03  RM0902085 T.Y オプション追加対応
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice54

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "NAB"
                    If (selectedData.Symbols(3).Trim = "8" Or selectedData.Symbols(3).Trim = "10") And _
                       (selectedData.Symbols(4).Trim = "" Or selectedData.Symbols(4).Trim = "0") Then
                        If selectedData.Symbols(2).Trim = "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        Select Case True
                            Case selectedData.Symbols(2).Trim <> "" And selectedData.Symbols(4).Trim <> ""
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case selectedData.Symbols(2).Trim = "" And selectedData.Symbols(4).Trim <> ""
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case selectedData.Symbols(2).Trim <> "" And selectedData.Symbols(4).Trim = ""
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "GNAB"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If selectedData.Symbols(5).Trim = "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                    End If
                    If selectedData.Symbols(4).Trim = "0" Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(4).Trim)
                    End If
                Case "NSB"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "NAD"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If selectedData.Symbols(2).Trim = "V" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*V-" & selectedData.Symbols(4).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*-" & selectedData.Symbols(4).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    '二次電池対応仕様
                    If selectedData.Symbols(6).Trim = "P4" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '食品製造工程向け対応
                    If selectedData.Series.key_kataban.Trim = "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-FP2"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "GNAD"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If selectedData.Symbols(2).Trim = "V" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*V-" & selectedData.Symbols(5).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "*-" & selectedData.Symbols(5).Trim
                    End If
                    If selectedData.Symbols(4).Trim = "0" Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(4).Trim)
                    End If

                    '二次電池対応仕様
                    If selectedData.Symbols(6).Trim = "P4" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & selectedData.Symbols(6).Trim & _
                             MyControlChars.Hyphen & selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '食品製造工程向け対応
                    If selectedData.Series.key_kataban.Trim = "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-FP2"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

            'マニホールドベース加算価格キー
            If selectedData.Series.series_kataban.Trim = "GNAB" Then
                If selectedData.Symbols(5).Trim = "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "BS" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "BS" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If
            If selectedData.Series.series_kataban.Trim = "GNAD" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'コイルオプション加算価格キー
            If selectedData.Series.series_kataban.Trim = "NSB" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'その他オプション加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "NAB"
                    strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "S"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "NSB"
                    strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "NAD"
                    If selectedData.Symbols(5).Trim = "B" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            'スイッチリード線長さ加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "NAB"
                    If selectedData.Series.key_kataban.Trim = "A" Then
                        If selectedData.Symbols(6).Trim <> "" Then
                            If selectedData.Symbols(6).Trim = "X" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim & _
                                                                           selectedData.Symbols(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                            End If
                        End If
                    End If
                Case "NSB"
                    If selectedData.Symbols(7).Trim <> "" Then
                        If selectedData.Symbols(7).Trim = "X" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8).Trim & _
                                                                       selectedData.Symbols(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        End If
                    End If
            End Select

            '電圧加算価格キー
            If selectedData.Series.series_kataban.Trim = "NSB" Then
                '電圧取得
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(10).Trim)
                Select Case strStdVoltageFlag
                    Case Divisions.VoltageDiv.Standard
                    Case Divisions.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OPT"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Divisions.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            '接続口径 NAB or GNAB
            If selectedData.Series.series_kataban.Trim = "NAB" Or selectedData.Series.series_kataban.Trim = "GNAB" Then
                'If selectedData.Symbols(6).Trim <> "" Then
                If Left(selectedData.Symbols(6).Trim, 1) <> Space(1) Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0902085 T.Yagyu
                    If selectedData.Series.series_kataban.Trim = "GNAB" Then
                        If CDec(selectedData.Symbols(4).Trim) <> 0 Then
                            decOpAmount(UBound(decOpAmount)) = CDec(selectedData.Symbols(4).Trim)
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
