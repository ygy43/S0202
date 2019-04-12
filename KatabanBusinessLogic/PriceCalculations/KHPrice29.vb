'************************************************************************************
'*  ProgramID  ：KHPrice29
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：扁平シリンダ　ＦＣＤ／ＦＣＨ／ＦＣＳ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice29

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(1).Trim), _
                                                  CInt(selectedData.Symbols(3).Trim))

            '基本価格キー
            Select Case True
                Case Mid(selectedData.Series.series_kataban, 4, 2) = MyControlChars.Hyphen & "L"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Mid(selectedData.Series.series_kataban, 4, 1) = MyControlChars.Hyphen
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Mid(selectedData.Series.series_kataban, 6, 1) = "L"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 4) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 6) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'マグネット(L)内蔵加算価格キー
            If Mid(selectedData.Series.series_kataban, 5, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban, 6, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "FC*" & MyControlChars.Hyphen & "L" & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'スイッチ加算価格キー
                If selectedData.Symbols(4).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "FC*" & selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)

                    'リード線長さ加算価格キー
                    If selectedData.Symbols(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "FC*" & selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                    End If
                End If
            End If

            'オプション加算価格キー
            If Mid(selectedData.Series.series_kataban, 5, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban, 6, 1) = "L" Then
                strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
            Else
                strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            End If
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "FC*" & strOpArray(intLoopCnt).Trim
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "R"
                                Select Case selectedData.Symbols(1).Trim
                                    Case "32"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "25"
                                    Case "50"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "40"
                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(1).Trim
                                End Select
                            Case "M"
                                strOpRefKataban(UBound(strOpRefKataban)) = "FC*" & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                           intStroke.ToString
                        End Select

                        If (selectedData.Series.series_kataban = "FCD-D" Or selectedData.Series.series_kataban = "FCD-DL") And _
                           strOpArray(intLoopCnt).Trim = "M" Then
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
