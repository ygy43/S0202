'************************************************************************************
'*  ProgramID  ：KHPrice49
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ジャスフィットバルブ
'*             ：Ｆ＊Ｂ／Ｆ＊Ｇ／ＧＦ＊Ｂ／ＧＦ＊Ｇ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice49

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strSeriesKataban As String
        Dim intValveQty As Integer
        Dim intMaskingQty As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'シリーズ形番設定
            If Left(selectedData.Series.series_kataban.Trim, 1) = "F" Then
                strSeriesKataban = selectedData.Series.series_kataban.Trim
            Else
                strSeriesKataban = Mid(selectedData.Series.series_kataban.Trim, 2, 3)
            End If

            '電磁弁＆マスキングプレート数設定
            If Left(selectedData.Series.series_kataban.Trim, 1) = "G" Then
                If selectedData.Symbols(4).Trim = "X" Then
                    intValveQty = CInt(selectedData.Symbols(9).Trim)
                    intMaskingQty = CInt(selectedData.Symbols(10).Trim)
                Else
                    If CInt(selectedData.Symbols(4).Trim) = 0 Then
                        intValveQty = 1
                    Else
                        intValveQty = CInt(selectedData.Symbols(4).Trim)
                    End If
                    intMaskingQty = 0
                End If
            Else
                intValveQty = 1
                intMaskingQty = 0
            End If

            '基本価格キー
            If Left(selectedData.Series.series_kataban.Trim, 1) = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           selectedData.Symbols(1).Trim & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = intValveQty
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = intValveQty
            End If

            'コイルオプション加算価格キー
            Select Case selectedData.Symbols(6).Trim
                Case "2G", "4A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
            End Select

            '手動装置加算価格キー
            If selectedData.Symbols(7).Trim <> "" Then
                If selectedData.Symbols(7).Trim = "A" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                End If
            End If

            'その他オプション加算価格キー
            If Left(selectedData.Series.series_kataban.Trim, 1) <> "G" Then
                If selectedData.Symbols(8).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(8).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                End If
            End If

            '電圧加算価格キー
            If Left(selectedData.Series.series_kataban.Trim, 1) = "G" Then
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(8).Trim)
                Select Case strStdVoltageFlag
                    Case Divisions.VoltageDiv.Standard
                    Case Divisions.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OPT" & MyControlChars.Hyphen & Left(selectedData.Symbols(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Divisions.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH" & MyControlChars.Hyphen & Left(selectedData.Symbols(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Else
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(9).Trim)
                Select Case strStdVoltageFlag
                    Case Divisions.VoltageDiv.Standard
                    Case Divisions.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OPT" & MyControlChars.Hyphen & Left(selectedData.Symbols(9).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Divisions.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH" & MyControlChars.Hyphen & Left(selectedData.Symbols(9).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'マスキングプレート加算価格キー
            If intMaskingQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                           selectedData.Symbols(1).Trim & "-MP-" & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = intMaskingQty
            End If

            'マニホールドベース加算価格キー
            If Left(selectedData.Series.series_kataban.Trim, 1) = "G" Then
                If selectedData.Symbols(4).Trim <> "0" Then
                    If selectedData.Symbols(4).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   selectedData.Symbols(1).Trim & "-BS-" & _
                                                                   (intMaskingQty + intValveQty).ToString & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   selectedData.Symbols(1).Trim & "-BS-" & _
                                                                   selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
