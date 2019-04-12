'************************************************************************************
'*  ProgramID  ：KHPrice35
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＢ（防爆）／ＡＧ（防爆）
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice35

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Mid(selectedData.Series.series_kataban.Trim, 5, 2) = "EX" Then
                Select Case True
                    Case Left(selectedData.Symbols(3).Trim, 1) = "H"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "03"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "0" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "J"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "B3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "B" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "K"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "C3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "C" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "L"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "D3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "D" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "M"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "E3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "E" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "N"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "F3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & "F" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Or Left(selectedData.Symbols(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & selectedData.Symbols(3).Trim & "3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & selectedData.Symbols(3).Trim & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Else
                Select Case True
                    Case Left(selectedData.Symbols(3).Trim, 1) = "H"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "03"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "0" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "J"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "B3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "B" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "K"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "C3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "C" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "L"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "D3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "D" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "M"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "E3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "E" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(selectedData.Symbols(3).Trim, 1) = "N"
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & "F" & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        If Left(selectedData.Symbols(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & selectedData.Symbols(3).Trim & "3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & selectedData.Symbols(3).Trim & selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            End If
            '支持形式加算価格キー
            Select Case Left(selectedData.Symbols(1).Trim, 2)
                Case "CA", "TC", "TF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                        selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'コイル別電圧加算価格キー
            If Mid(selectedData.Series.series_kataban.Trim, 5, 2) = "EX" Then
                If Left(selectedData.Symbols(4).Trim, 1) <> "3" And _
                 Left(selectedData.Symbols(4).Trim, 1) <> "4" Or _
                 selectedData.Symbols(8).Trim <> Divisions.PowerSupply.Const1 And _
                 selectedData.Symbols(8).Trim <> Divisions.PowerSupply.Const2 Then
                    If selectedData.Series.series_kataban.Trim <> "AG41E4" Then
                        If KatabanUtility.GetVoltageIsStandard(selectedData.Symbols(8).Trim, strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & selectedData.Symbols(4).Trim & _
                                                                                Left(selectedData.Symbols(8).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        'コイルオプション加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & selectedData.Symbols(4).Trim & _
                                                                            Left(selectedData.Symbols(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                If Left(selectedData.Symbols(4).Trim, 1) <> "3" And _
                   Left(selectedData.Symbols(4).Trim, 1) <> "4" Or _
                   selectedData.Symbols(7).Trim <> Divisions.PowerSupply.Const1 And _
                   selectedData.Symbols(7).Trim <> Divisions.PowerSupply.Const2 Then
                    If selectedData.Series.series_kataban.Trim <> "AG41E4" Then
                        '2010/08/26 ADD RM0808112(異電圧対応) START--->

                        If KatabanUtility.GetVoltageIsStandard(selectedData.Symbols(7).Trim, strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & selectedData.Symbols(4).Trim & _
                                                                                Left(selectedData.Symbols(7).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & selectedData.Symbols(4).Trim & _
                        '                                                    Left(selectedData.Symbols(7).Trim, 2)
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/08/26 ADD RM0808112(異電圧対応) <--- END
                    Else
                        'コイルオプション加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & selectedData.Symbols(4).Trim & _
                                                                            Left(selectedData.Symbols(7).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If
         
            '外部導線引込方式加算価格キー
            Select Case Left(selectedData.Symbols(5).Trim, 1)
                Case "L", "M", "N", "P"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & MyControlChars.Hyphen & selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション・付属品価格
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
