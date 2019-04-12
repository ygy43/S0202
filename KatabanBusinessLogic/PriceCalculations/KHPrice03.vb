'************************************************************************************
'*  ProgramID  ：KHPrice03
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：直動式２ポート弁　ＡＢ／ＡＧ
'*
'*  ・受付No：RM0907070  二次電池対応機器　
'*                                      更新日：2009/09/08   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　KHOptionCtl.vb
'*                                      更新日：2010/02/22   更新者：Y.Miura
'*  ・受付No：RM0808112  異電圧対応
'*                                      更新日：2010/08/11   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice03

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strStdVoltageFlag As String = Divisions.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim strPort As String
        Dim bolScrew As Boolean
        Dim intStationQty As Integer = 0
        Dim bolOptionZ As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '基本価格キー
            If InStr(selectedData.Symbols(1).Trim, "G") <> 0 Or _
               InStr(selectedData.Symbols(1).Trim, "N") <> 0 Then
                strPort = "0" & Left(selectedData.Symbols(1).Trim, 1)
                bolScrew = True
            Else
                strPort = Left(selectedData.Symbols(1).Trim, 2)
                bolScrew = False
            End If
            If selectedData.Series.series_kataban.Trim = "AB71" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strPort & MyControlChars.Hyphen & _
                                                           Left(selectedData.Symbols(2).Trim, 2)
                decOpAmount(UBound(decOpAmount)) = 1
            Else

                If Left(selectedData.Symbols(3).Trim, 1) <> "" And _
                   Left(selectedData.Symbols(3).Trim, 1) <> "0" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               strPort & MyControlChars.Hyphen & _
                                                               Left(selectedData.Symbols(2).Trim, 1) & MyControlChars.Hyphen & _
                                                               Left(selectedData.Symbols(3).Trim, 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               strPort & MyControlChars.Hyphen & _
                                                               Left(selectedData.Symbols(2).Trim, 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'コイルハウジング加算価格キー
            If selectedData.Series.series_kataban.Trim = "AB71" Then
                If selectedData.Symbols(3).Trim <> "" Then
                    If Left(selectedData.Symbols(3).Trim, 1) <> "B" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_03" & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                If selectedData.Symbols(4).Trim <> "" Then
                    If Left(selectedData.Symbols(4).Trim, 1) = "2" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'RM0907070 2009/09/08 Y.Miura 二次電池対応
                        'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_04" & _
                        '                                           Left(selectedData.Symbols(4).Trim, 2) & _
                        '                                           Left(selectedData.Symbols(9).Trim, 2)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_04" & _
                                                                   Left(selectedData.Symbols(4).Trim, 2) & _
                                                                   Left(selectedData.Symbols(10).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_04" & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

            '手動操作加算価格キー
            If selectedData.Series.series_kataban.Trim = "AB71" Then
                If selectedData.Symbols(4).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_04" & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If selectedData.Symbols(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_05" & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '取付板加算価格キー
            If selectedData.Series.series_kataban.Trim = "AB71" Then
                If selectedData.Symbols(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_05" & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If selectedData.Symbols(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_06" & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'ケーブルグランド・コンジット加算価格キー
            If selectedData.Series.series_kataban.Trim = "AB71" Then
                If selectedData.Symbols(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_06" & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If selectedData.Symbols(7).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_07" & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'オプション加算価格キー
            If selectedData.Series.series_kataban.Trim <> "AB71" Then
                strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case "S"
                            If Left(selectedData.Symbols(4).Trim, 2) = "" Or _
                               Left(selectedData.Symbols(4).Trim, 2) = "00" Or _
                               Left(selectedData.Symbols(4).Trim, 2) = "3A" Or _
                               Left(selectedData.Symbols(4).Trim, 2) = "4A" Or _
                               Left(selectedData.Symbols(4).Trim, 2) = "6C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_08S0"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_08" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_08" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next

                'オプション２　Ｐ４加算
                'RM0907070 2009/09/08 Y.Miura　二次電池対応
                strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else       '"P4", "P40"を含む
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                Next

            End If

            '電圧加算価格キー
            If selectedData.Series.series_kataban.Trim <> "AB71" Then
                'RM0907070 2009/09/08 Y.Miura 二次電池対応
                'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                '                                               selectedData.Symbols(9).Trim)
                'RM0808112　異電圧対応
                'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                '                                               selectedData.Symbols(10).Trim)
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(10).Trim, strCountryCd, strOfficeCd)
                Select Case strStdVoltageFlag
                    Case Divisions.VoltageDiv.Standard
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'RM0907070 2009/09/08 Y.Miura 二次電池対応
                        'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_09" & _
                        '                                           Left(selectedData.Symbols(9).Trim, 2)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_09" & _
                                                                   Left(selectedData.Symbols(10).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'ねじ加算価格キー
            'If bolScrew Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "MULTI-SCREW-" & Right(selectedData.Symbols(1).Trim, 1)
            '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            '    Select Case Left(selectedData.Series.series_kataban.Trim, 2)
            '        Case "AB"
            '            decOpAmount(UBound(decOpAmount)) = 2
            '        Case "AG"
            '            decOpAmount(UBound(decOpAmount)) = 3
            '    End Select
            'End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
