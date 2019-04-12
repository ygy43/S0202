'************************************************************************************
'*  ProgramID  ：KHPrice17
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*
'*  概要       ：ＡＰ／ＡＤ
'*
'*  修正履歴   ：
'*                                      更新日：2008/03/27   更新者：NII A.Takahashi
'*  ・G/NPTねじ追加により、ロジック変更(ねじ加算対応)
'*                                      更新日：2008/07/20      更新者：T.Sato
'*  ・受付No：RM0806072　AP11,AP21,AD11,AD21 コイルハウジング低ワット電圧追加
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice17

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim strOption As String
        Dim strScrewType As String
        Dim bolScrew As Boolean
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'ねじ判定
            If InStr(selectedData.Symbols(1).Trim, "G") <> 0 Or _
               InStr(selectedData.Symbols(1).Trim, "N") <> 0 Then
                strScrewType = Right(selectedData.Symbols(1).Trim, 1)
                bolScrew = True
            Else
                strScrewType = ""
                bolScrew = False
            End If

            '基本価格キー
            Select Case True
                Case Left(selectedData.Symbols(2).Trim, 1) = "H"
                    strOption = "0"
                Case Left(selectedData.Symbols(2).Trim, 1) = "J"
                    strOption = "B"
                Case Left(selectedData.Symbols(2).Trim, 1) = "K"
                    strOption = "C"
                Case Left(selectedData.Symbols(2).Trim, 1) = "L"
                    strOption = "D"
                Case Left(selectedData.Symbols(2).Trim, 1) = "M"
                    strOption = "E"
                Case Left(selectedData.Symbols(2).Trim, 1) = "N"
                    strOption = "F"
                Case Else
                    strOption = selectedData.Symbols(2).Trim
            End Select

            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            If bolScrew Then
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           Left(selectedData.Symbols(1).Trim, (InStr(1, selectedData.Symbols(1).Trim, strScrewType)) - 1) & _
                                                           MyControlChars.Hyphen & strOption
            Else
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & strOption
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'ボディシール加算
            Select Case selectedData.Symbols(2).Trim
                Case "0", "B", "C", "D", "E", "F"
                Case "H"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "HK"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'コイルハウジング加算
            If Left(selectedData.Symbols(3).Trim, 1) = "2" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           Left(selectedData.Symbols(3).Trim, 2) & _
                                                           Left(selectedData.Symbols(5).Trim, 2)
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           selectedData.Symbols(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション価格
            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "S"
                        Select Case selectedData.Symbols(3).Trim
                            Case "2C", "3A", "4A", "6C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "S0"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算価格キー
            If selectedData.Symbols(5).Trim <> "" Then
                '電圧取得
                '2010/08/27 ADD RM0808112(異電圧対応) START--->
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(5).Trim, _
                                                               strCountryCd, strOfficeCd)
                'strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                '                                               selectedData.Symbols(5).Trim)
                '2010/08/27 ADD RM0808112(異電圧対応) <--- END
                If strStdVoltageFlag <> Divisions.VoltageDiv.Standard Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "AC"
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & Left(selectedData.Symbols(5).Trim, 2)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'ねじ加算価格キー
            'If bolScrew Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "MULTI-SCREW-" & Right(selectedData.Symbols(1).Trim, 1)
            '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            '    decOpAmount(UBound(decOpAmount)) = 2
            'End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
