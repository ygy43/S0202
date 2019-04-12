'************************************************************************************
'*  ProgramID  ：KHPrice10
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＲＳＶ
'*             ：ＲＳＶ－Ｗ／Ｈ
'*             ：ＧＳＶ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice10

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

            Select Case selectedData.Series.series_kataban.Trim
                Case "RSV"
                    Select Case selectedData.Series.key_kataban.Trim
                        Case ""
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'コイルハウジング加算価格キー
                            If selectedData.Symbols(4).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'その他電圧加算価格キー
                            If selectedData.Symbols(5).Trim <> "" Then
                                '電圧取得
                                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                               selectedData.Symbols(5).Trim)
                                Select Case strStdVoltageFlag
                                    Case Divisions.VoltageDiv.Standard
                                    Case Divisions.VoltageDiv.Options
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                        Case Else
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'コイルハウジング加算価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & "W" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'その他電圧加算価格キー
                            If selectedData.Symbols(8).Trim <> "" Then
                                '電圧取得
                                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                               selectedData.Symbols(8).Trim)
                                Select Case strStdVoltageFlag
                                    Case Divisions.VoltageDiv.Standard
                                    Case Divisions.VoltageDiv.Options
                                    Case Else
                                        If Left(selectedData.Symbols(8).Trim, 2) = "AC" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-W-AC-OTH"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-W-DC-OTH"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                End Select
                            End If
                    End Select
                Case "GSV"
                    '基本価格キー
                    If selectedData.Symbols(3).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'その他電圧加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        '電圧取得
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(4).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If

                Case "GSV2"
                    '基本価格キー
                    If selectedData.Symbols(2).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'その他電圧加算価格キー
                    If selectedData.Symbols(4).Trim <> "" Then
                        '電圧取得
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(4).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If

            End Select


            'オプション価格
            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "_04" & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
