'************************************************************************************
'*  ProgramID  ：KHPriceK4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ペンシルシリンダ　ＳＣＰＤ２－＊Ｃ／ＳＣＰＤ２－Ｌ－＊Ｃ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceK4

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(4).Trim))
            Select selectedData.Series.series_kataban
                Case "SCPG2", "SCPG2-L"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション(2)加算価格価格キー
                    If Mid(selectedData.Series.series_kataban, 7, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                        'リード線長さ加算価格価格キー
                        If selectedData.Symbols(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        End If
                    End If

                    'オプション・付属品価格キー
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                Case Else
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション(2)加算価格価格キー
                    If Mid(selectedData.Series.series_kataban, 7, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    If selectedData.Symbols(1).Trim = "CB" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'スイッチ加算価格キー
                    If selectedData.Symbols(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                        'リード線長さ加算価格価格キー
                        If selectedData.Symbols(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                        End If
                    End If

                    'オプション・付属品価格キー
                    strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select
        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
