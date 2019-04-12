'************************************************************************************
'*  ProgramID  ：KHPriceC7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：マグネット式スーパーロッドレスシリンダ　ＭＲＬ２
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceC7

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim bolOptionP4 As Boolean = False      'RM1001045 2010/02/23 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
            If selectedData.Symbols.Count > 8 Then
                strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "P4", "P40"
                            bolOptionP4 = True
                    End Select
                Next
            End If

            'RM1306001 2013/06/05 追加
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(4).Trim))

            '基本価格キー
            Select Case True
                Case Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'RM1306001 2013/06/05 追加
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            'マグネット内蔵(L)加算キー
            If Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                Select Case selectedData.Symbols(2).Trim
                    Case "6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR200"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "10", "16"
                        Select Case True
                            Case CInt(selectedData.Symbols(4).Trim) <= 200
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR200"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case CInt(selectedData.Symbols(4).Trim) >= 201
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR201"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "20", "25", "32"
                        Select Case True
                            Case CInt(selectedData.Symbols(4).Trim) <= 200
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR200"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case CInt(selectedData.Symbols(4).Trim) >= 201 And _
                                 CInt(selectedData.Symbols(4).Trim) <= 500
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR201"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case CInt(selectedData.Symbols(4).Trim) >= 501
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR501"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            End If

            '微速(F)加算価格キー
            Select Case selectedData.Symbols(1).Trim
                Case "F"
                    Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                        Case "G", "W"
                            Select Case selectedData.Symbols(2).Trim
                                Case "6"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR101"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "10"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 151
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR151"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 250
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR250"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 251
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR251"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "20", "25", "32"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 400
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR400"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 401
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR401"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                        Case Else
                            Select Case selectedData.Symbols(2).Trim
                                Case "6"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR101"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "10"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 151
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR151"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "16"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 250
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR250"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 251
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR251"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "20", "25", "32"
                                    Select Case True
                                        Case CInt(selectedData.Symbols(4).Trim) <= 400
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR400"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(selectedData.Symbols(4).Trim) >= 401
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
                                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR401"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select
                Case Else
            End Select

            'RM1306001 2013/06/05 追加
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            'ゴムエアクッション(C)加算キー
            Select Case selectedData.Symbols(3).Trim
                Case "C"
                    Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                        Case "W"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "GOMC" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "G" & MyControlChars.Hyphen & "GOMC" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
            End Select

            If Mid(selectedData.Series.series_kataban.Trim, 6, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
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
                            Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                                 "T3V", "T3YH", "T3YV", "T1H", "T1V", _
                                 "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "SW1" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                            Case "T2YFH", "T2YFV", "T2YMH", "T2YMV", "T3YFH", _
                                 "T3YFV", "T3YMH", "T3YMV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & "SW2" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                        End Select
                    End If

                    'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                    'P4加算
                    If bolOptionP4 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-SW-P4"
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                    End If
                End If
            End If

            'オプション加算キー
            '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) START--->
            Dim isC As Boolean = False
            Dim strP4 As String = ""
            '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) <---END

            strOpArray = Split(selectedData.Symbols(8), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "C"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1306001 2013/06/05 追加
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If

                        '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) START--->
                        isC = True
                        '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) <---END

                    Case "S"
                        Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                            Case "G", "W"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                           "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM1306001 2013/06/05 追加
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM1306001 2013/06/05 追加
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                        End Select
                    Case "R"
                        Select Case selectedData.Symbols(2).Trim
                            Case "10", "16"
                                Select Case True
                                    Case CInt(selectedData.Symbols(4).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        'RM1306001 2013/06/05 追加
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                        End If
                                    Case CInt(selectedData.Symbols(4).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR201"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        'RM1306001 2013/06/05 追加
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                        End If
                                End Select
                            Case "20", "25", "32"
                                Select Case True
                                    Case CInt(selectedData.Symbols(4).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case CInt(selectedData.Symbols(4).Trim) >= 201 And _
                                         CInt(selectedData.Symbols(4).Trim) <= 500
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR201"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case CInt(selectedData.Symbols(4).Trim) >= 501
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "STR501"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        'RM1306001 2013/06/05 追加
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                        End If
                                End Select
                        End Select
                    Case "P4", "P40"        'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '2011/03/04 ADD RM1103016(4月VerUP：MRL2-G,W P4※シリーズ) START--->
                        Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                            Case "G", "W"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & "-OP-" & strOpArray(intLoopCnt).Trim
                            Case Else
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-OP-" & strOpArray(intLoopCnt).Trim
                        End Select
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & "-OP-" & strOpArray(intLoopCnt).Trim
                        '2011/03/04 ADD RM1103016(4月VerUP：MRL2-G,W P4※シリーズ) <---END
                        decOpAmount(UBound(decOpAmount)) = 1

                        '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) START--->
                        strP4 = strOpArray(intLoopCnt).Trim
                        '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) <---END
                        'RM1306001 2013/06/05 追加
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If

                    Case "P72"
                        Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                            Case "G", "W"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                           "G" & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM1306001 2013/06/05 追加
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM1306001 2013/06/05 追加
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                        End Select
                End Select

            Next


            '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) START--->
            'ショックアブソーバ付（Ｃ)オプションと二次電池加算を併用した場合、価格加算
            If isC AndAlso Len(strP4) > 0 Then

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                           "C" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim & _
                                                           MyControlChars.Hyphen & strP4
                decOpAmount(UBound(decOpAmount)) = 1

            End If
            '2011/04/04 ADD RM1104022(5月VerUP:MRL2 P4シリーズ) <---END

            'クリーン仕様加算価格キー
            If selectedData.Symbols(9).Trim <> "" Then
                Select Case Mid(selectedData.Series.series_kataban.Trim, 6, 1)
                    Case "G", "W"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
