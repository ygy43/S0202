'************************************************************************************
'*  ProgramID  ：KHPrice88
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空エジェクタユニット単体
'*             ：真空切替ユニット単体
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice88

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            ' 機種毎に価格キーを設定する
            Select Case selectedData.Series.series_kataban.Trim
                Case "VSK"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & "**" & _
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & "**" & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSJ"
                    If selectedData.Symbols(8).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "***" & MyControlChars.Hyphen & "**" & _
                                                                   selectedData.Symbols(6).Trim & MyControlChars.Hyphen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "***" & MyControlChars.Hyphen & "**" & _
                                                                   selectedData.Symbols(6).Trim & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "VSN"
                    ' 真空センサ仕様 
                    If Len(selectedData.Symbols(8).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-**-****-" & _
                                                                   selectedData.Symbols(7).Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-**-***" & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "VSJP"
                    If selectedData.Symbols(6).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "****" & MyControlChars.Hyphen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "****" & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "VSNP"
                    ' 真空センサ仕様 
                    If Len(selectedData.Symbols(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-***-" & _
                                                                   selectedData.Symbols(4).Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-***-" & _
                                                                   selectedData.Symbols(4).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "VSX"
                    '基本キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "***" & MyControlChars.Hyphen & "**" & _
                                                               selectedData.Symbols(6).Trim & MyControlChars.Hyphen & "*"

                    '真空センサ仕様
                    If selectedData.Symbols(8).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim
                    End If

                    '取付方法
                    If selectedData.Symbols(9).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9).Trim
                    End If

                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1806035_二次電池機種追加対応
                    If selectedData.Series.key_kataban.Trim = "P" Then

                        '基本キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(10).Trim

                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "VSXP"
                    '基本キー
                    '基本キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & "***" & MyControlChars.Hyphen & "*"

                    '真空センサ仕様
                    If selectedData.Symbols(6).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                    End If

                    '取付方法
                    If selectedData.Symbols(7).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7).Trim
                    End If

                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSQ"
                    ' 基本キー
                    Select Case Left(selectedData.Symbols(1).Trim, 1)
                        Case "T"
                            If selectedData.Symbols(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "T**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "T**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "D"
                            If selectedData.Symbols(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "D**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "D**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            If selectedData.Symbols(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "**" & MyControlChars.Hyphen & "**" & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "VSQP"
                    If selectedData.Symbols(5).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "***" & MyControlChars.Hyphen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "***" & MyControlChars.Hyphen & "*" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    'RM1806035_二次電池機種追加対応
                Case "VSFU"

                    If selectedData.Series.key_kataban.Trim = "P" Then

                        '基本価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Ｐ４加算価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
