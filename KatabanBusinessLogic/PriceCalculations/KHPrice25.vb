'************************************************************************************
'*  ProgramID  ：KHPrice25
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*
'*  概要       ：スーパーロッドレスシリンダ　ＳＲＬ２
'*             ：ブレーキ付ロッドレスシリンダ　ＳＲＢ２
'*             ：ガイド付ロッドレスシリンダ　ＳＲＧ
'*             ：ガイド付ロッドレスシリンダ　ＳＲＧ３
'*
'*  更新履歴   ：                       更新日：2009/02/05   更新者：T.Yagyu
'*               ・RM0811134:SRG3機種追加
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice25

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)



        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        'RM0811134:SRG3 T.Y
        Dim strMountingStyle As String '支持形式
        Dim strBoreSize As String 'チューブ内径
        Dim strPipeThreadType As String '配管ねじ種類 ソース内では未使用
        Dim strCushion As String 'クッション
        Dim strStrokeLen As String 'ストローク
        Dim strSwModelNo As String 'スイッチ形番
        Dim strLeadWireLen As String 'リード線長さ
        Dim strSwQuantity As String 'スイッチ数
        Dim strOption As String 'オプション
        Dim bolC5Flag As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'RM0811134:SRG3 T.Y
            'SRG3のときだけstrOpSbl3（配管ねじ種類）が発生する
            'プログラムを共通化するためにselectedData.Symbolsの値を変数にセットし利用する
            Select Case selectedData.Series.series_kataban.Trim
                Case "SRG3"
                    strMountingStyle = selectedData.Symbols(1)
                    strBoreSize = selectedData.Symbols(2)
                    strPipeThreadType = selectedData.Symbols(3)
                    strCushion = selectedData.Symbols(4)
                    strStrokeLen = selectedData.Symbols(5)
                    strSwModelNo = selectedData.Symbols(6)
                    strLeadWireLen = selectedData.Symbols(7)
                    strSwQuantity = selectedData.Symbols(8)
                    strOption = selectedData.Symbols(9)
                Case Else
                    strMountingStyle = selectedData.Symbols(1)
                    strBoreSize = selectedData.Symbols(2)
                    strCushion = selectedData.Symbols(3)
                    strStrokeLen = selectedData.Symbols(4)
                    strSwModelNo = selectedData.Symbols(5)
                    strLeadWireLen = selectedData.Symbols(6)
                    strSwQuantity = selectedData.Symbols(7)
                    strOption = selectedData.Symbols(8)
            End Select

            'RM1306001 2013/06/06
            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(strBoreSize.Trim), _
                                                  CInt(strStrokeLen.Trim))

            '基本価格キー
            If Mid(selectedData.Series.series_kataban, 6, 1) = "Q" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & _
                                                           strBoreSize.Trim & MyControlChars.Hyphen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & MyControlChars.Hyphen & _
                                                           strBoreSize.Trim & MyControlChars.Hyphen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            'バリエーションQ加算価格キー
            If Mid(selectedData.Series.series_kataban, 6, 1) = "Q" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & MyControlChars.Hyphen & _
                                                           strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '支持形式加算価格キー
            If strMountingStyle.Trim <> "00" Then
                Select Case True
                    Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                   strMountingStyle.Trim & MyControlChars.Hyphen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Mid(selectedData.Series.series_kataban, 5, 1) = MyControlChars.Hyphen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strMountingStyle.Trim & MyControlChars.Hyphen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                   strMountingStyle.Trim & MyControlChars.Hyphen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'スイッチ加算価格キー
            If strSwModelNo.Trim <> "" Then
                Select Case True
                    Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                    Case Mid(selectedData.Series.series_kataban, 5, 1) = MyControlChars.Hyphen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                End Select

                'リード線長さ加算価格キー
                If strLeadWireLen.Trim <> "" Then
                    Select Case True
                        Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                            End Select
                        Case Mid(selectedData.Series.series_kataban, 5, 1) = MyControlChars.Hyphen
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                            End Select
                        Case Else
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               strLeadWireLen.Trim & MyControlChars.Hyphen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwQuantity.Trim)
                            End Select
                    End Select
                End If
            End If

            'オプション・付属品価格キー
            strOpArray = Split(strOption, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        If Left(strOpArray(intLoopCnt).Trim, 1) = "L" Or Left(strOpArray(intLoopCnt).Trim, 1) = "N" Then
                            Select Case True
                                Case Mid(selectedData.Series.series_kataban, 6, 1) = "Q" And Left(strOpArray(intLoopCnt).Trim, 1) = "A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Mid(selectedData.Series.series_kataban, 5, 1) = MyControlChars.Hyphen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                            End Select

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            decOpAmount(UBound(decOpAmount)) = Val(Mid(strOpArray(intLoopCnt).Trim, 2, 1))

                        Else
                            Select Case True
                                Case Mid(selectedData.Series.series_kataban, 6, 1) = "Q" And Left(strOpArray(intLoopCnt).Trim, 1) = "A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & _
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Mid(selectedData.Series.series_kataban, 5, 1) = MyControlChars.Hyphen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & _
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban & _
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               strBoreSize.Trim
                            End Select

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                End Select
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
