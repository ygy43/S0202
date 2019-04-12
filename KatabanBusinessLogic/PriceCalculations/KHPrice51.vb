'************************************************************************************
'*  ProgramID  ：KHPrice51
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/24   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：モータバルブ　ＭＳＢ／ＭＳＧ／ＭＸＢ／ＭＸＧ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice51

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
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim
            If selectedData.Symbols(1).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(1).Trim
            End If
            If selectedData.Symbols(2).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(2).Trim
            End If
            If selectedData.Symbols(3).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(3).Trim
            End If
            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen
            If selectedData.Symbols(5).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(5).Trim
            End If
            If selectedData.Symbols(6).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(6).Trim
            End If
            If selectedData.Symbols(7).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & _
                                                           selectedData.Symbols(7).Trim
            End If
            strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut(strOpRefKataban(UBound(strOpRefKataban)))
            If selectedData.Symbols(8).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(8).Trim
            End If
            If selectedData.Symbols(9).Trim.IndexOf("K") >= 0 Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & "-K"
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'その他オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "", "K"
                    Case Else
                        If strOpArray(intLoopCnt).Trim = "E" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

            '食品製造工程向
            Select Case selectedData.Series.key_kataban
                Case "F"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(10).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '電圧加算価格キー
            Select Case selectedData.Series.key_kataban
                Case "F"

                    If selectedData.Symbols(11).Trim <> "" Then
                        '電圧取得
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(11).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                                Select Case selectedData.Symbols(11).Trim
                                    Case "1", "2", "AC110V", "AC120V", "AC220V", "AC240V" 'RM1609026 2016/09/13
                                        'Case "1", "2"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OPT-AC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OPT-DC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Case Divisions.VoltageDiv.Other
                                Select Case selectedData.Symbols(10).Trim
                                    Case "1", "2"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTH-AC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTH-DC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                        End Select
                    End If

                Case Else

                    If selectedData.Symbols(10).Trim <> "" Then
                        '電圧取得
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(10).Trim)
                        Select Case strStdVoltageFlag
                            Case Divisions.VoltageDiv.Standard
                            Case Divisions.VoltageDiv.Options
                                Select Case selectedData.Symbols(10).Trim
                                    Case "1", "2", "AC110V", "AC120V", "AC220V", "AC240V" 'RM1609026 2016/09/13
                                        'Case "1", "2"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OPT-AC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OPT-DC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Case Divisions.VoltageDiv.Other
                                Select Case selectedData.Symbols(10).Trim
                                    Case "1", "2"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTH-AC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTH-DC"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                        End Select
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        Finally
        End Try

    End Sub

End Module
