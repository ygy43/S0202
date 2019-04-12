Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Namespace Managers
    Public Class RodEndManager
        ''' <summary>
        '''     ロッド先端情報の取得
        ''' </summary>
        ''' <param name="info"></param>
        ''' <returns></returns>
        Public Shared Function GetRodEndInfo(info As SelectedInfo) As List(Of RodEndInfo)

            Dim series = info.Series.series_kataban
            Dim keyKataban = info.Series.key_kataban

            Using client As New DbAccessServiceClient

                'ロッド先端情報を取得
                Dim rodEndInfos = client.SelectRodEndInfo(series, keyKataban)

                '特殊形番の設定
                If series = "JSC3" AndAlso keyKataban = "1" AndAlso info.Symbols(4).Contains("FA") Then
                    'ロッド先端特注画面イメージ
                    rodEndInfos(0).url = String.Empty
                    rodEndInfos(1).url = "../KHImage/JSC3RodN13(FA).gif"
                    rodEndInfos(2).url = "../KHImage/JSC3RodN15(FA).gif"
                    rodEndInfos(3).url = "../KHImage/JSC3RodN11(FA).gif"
                    rodEndInfos(4).url = "../KHImage/JSC3RodN1(FA).gif"
                    rodEndInfos(5).url = "../KHImage/JSC3RodN12(FA).gif"
                    rodEndInfos(6).url = "../KHImage/JSC3RodN14(FA).gif"
                    rodEndInfos(7).url = "../KHImage/JSC3RodN3(FA).gif"
                    rodEndInfos(8).url = "../KHImage/JSC3RodN31(FA).gif"
                    rodEndInfos(9).url = "../KHImage/JSC3RodN2(FA).gif"
                    rodEndInfos(10).url = "../KHImage/JSC3RodN21(FA).gif"
                    rodEndInfos(11).url = String.Empty

                    'ロッド先端特注パターン記号
                    rodEndInfos(0).rod_pattern_symbol = String.Empty
                    rodEndInfos(1).rod_pattern_symbol = "N13"
                    rodEndInfos(2).rod_pattern_symbol = "N15"
                    rodEndInfos(3).rod_pattern_symbol = "N11"
                    rodEndInfos(4).rod_pattern_symbol = "N1"
                    rodEndInfos(5).rod_pattern_symbol = "N12"
                    rodEndInfos(6).rod_pattern_symbol = "N14"
                    rodEndInfos(7).rod_pattern_symbol = "N3"
                    rodEndInfos(8).rod_pattern_symbol = "N31"
                    rodEndInfos(9).rod_pattern_symbol = "N2"
                    rodEndInfos(10).rod_pattern_symbol = "N21"
                    rodEndInfos(11).rod_pattern_symbol = "Other"
                End If

                Return rodEndInfos
            End Using
        End Function

        ''' <summary>
        '''     ロッド先端外径寸法などの情報の取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="boreSize">口径</param>
        ''' <returns></returns>
        Public Shared Function GetRodEndExternalFormInfo(series As String, keyKataban As String, boreSize As Integer) _
            As List(Of RodEndExternalFormInfo)

            Using client As New DbAccessServiceClient

                'ロッド先端外径寸法などの情報を取得
                Return client.SelectRodEndExternalFormInfo(series, keyKataban, boreSize)

            End Using
        End Function

        ''' <summary>
        '''     ロッド先端ユニット選択可否の判断
        ''' </summary>
        ''' <param name="pattern">ロッド先端パタン</param>
        ''' <param name="info">引当情報</param>
        ''' <returns></returns>
        Public Shared Function IsRodEndUnitEnable(pattern As String,
                                                  info As SelectedInfo,
                                                  boreSize As Integer) As Boolean
            Dim result = True

            If pattern = "Other" Then
                'その他寸法エリア設定
                Select Case info.Series.series_kataban
                    Case "SSD"
                        Select Case info.Series.key_kataban
                            Case "D"
                                result = False
                        End Select
                    Case "SCA2"
                        Select Case info.Series.key_kataban
                            Case "D"
                                result = False
                        End Select
                    Case "SCS"
                        Select Case info.Series.key_kataban
                            Case "D"
                                result = False
                        End Select
                    Case "CMK2"
                        Select Case info.Series.key_kataban
                            Case "D"
                                result = False
                        End Select
                End Select
            Else
                Select Case info.Series.series_kataban
                    Case "SSD"

                        '口径が12,16の場合N3/N31/N2/N21選択不可
                        Select Case boreSize
                            Case 12, 16
                                Select Case pattern
                                    Case Divisions.RodEndPatternDiv.N3, RodEndPatternDiv.N31,
                                        RodEndPatternDiv.N2, RodEndPatternDiv.N21
                                        result = False
                                End Select
                        End Select

                        Select Case info.Series.key_kataban
                            Case ""
                                '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                                Select Case info.Symbols(21)
                                    Case "I", "Y", "I2", "Y2"
                                        Select Case pattern
                                            Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                            Case Else
                                                result = False
                                        End Select
                                End Select
                                'N13-N11/N11-N13選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                        result = False
                                End Select
                            Case "K"
                                '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                                Select Case info.Symbols(19)
                                    Case "I", "Y", "I2", "Y2"
                                        Select Case pattern
                                            Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                            Case Else
                                                result = False
                                        End Select
                                End Select
                                'バリエーション「U」を含んでいる場合N12/N14選択不可
                                If InStr(info.Symbols(1), "U") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N12, RodEndPatternDiv.N14
                                            result = False
                                    End Select
                                End If
                                'N13-N11/N11-N13選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                        result = False
                                End Select
                            Case "D"
                                If InStr(info.Symbols(1), "Q") <> 0 Or
                                   info.Symbols(7).Trim = "R" Then
                                Else
                                    Select Case pattern
                                        Case RodEndPatternDiv.N11N13
                                            result = False
                                    End Select
                                End If
                                '中間ストロークの場合N11-N13選択可
                                Select Case info.Symbols(6).Trim
                                    Case "5", "10", "15", "20", "25", "30", "40", "50",
                                        "60", "70", "80", "90", "100", "110", "120",
                                        "130", "140", "150", "160", "170", "180", "190",
                                        "200", "210", "220", "230", "240", "250", "260",
                                        "270", "280", "290", "300"
                                        Select Case pattern
                                            Case RodEndPatternDiv.N11N13
                                                result = False
                                        End Select
                                End Select
                                '支持金具に「FA」を含む場合N11-N13選択可
                                If InStr(info.Symbols(12), "FA") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N11N13
                                            result = False
                                    End Select
                                End If
                                'N13-N11/N11-N13以外選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                    Case Else
                                        result = False
                                End Select
                        End Select
                    Case "JSC3"
                        '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                        If InStr(info.Symbols(14), "I") <> 0 Or
                           InStr(info.Symbols(14), "Y") <> 0 Then
                            Select Case pattern
                                Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                Case Else
                                    result = False
                            End Select
                        End If
                    Case "JSC4"
                        '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                        If InStr(info.Symbols(14), "I") <> 0 Or
                           InStr(info.Symbols(14), "Y") <> 0 Then
                            Select Case pattern
                                Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                Case Else
                                    result = False
                            End Select
                        End If
                    Case "SCA2"

                        '付属品初期値設定
                        Dim position = 0

                        '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                        Select Case info.Series.key_kataban
                            Case "", "V"
                                position = 14
                            Case "B"
                                position = 18
                            Case "D"
                                position = 13
                            Case "2"
                                position = 15
                            Case "C"
                                position = 19
                            Case "E"
                                position = 14
                        End Select
                        If InStr(info.Symbols(position), "I") <> 0 Or
                           InStr(info.Symbols(position), "Y") <> 0 Then
                            Select Case pattern
                                Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                Case Else
                                    result = False
                            End Select
                        End If

                        Select Case info.Series.key_kataban
                            Case "", "V", "B", "2", "C"
                                'N13-N11/N11-N13を選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                        result = False
                                End Select
                            Case "D", "E"
                                ' バリエーションに「Q」を含み、落下防止機構で「HR」を選択しない場合は「N11-N13」は選択可
                                If InStr(info.Symbols(1), "Q") <> 0 And
                                   InStr(info.Symbols(8), "HR") = 0 Then
                                Else
                                    Select Case pattern
                                        Case RodEndPatternDiv.N11N13
                                            result = False
                                    End Select
                                End If
                                'N13-N11/N11-N13以外を選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                    Case Else
                                        result = False
                                End Select
                        End Select
                    Case "SCS"
                        Select Case info.Series.key_kataban
                            Case "", "B"
                                'オプション「A2」を含む、もしくは付属品「I」「Y」を含む場合はN13/N15以外を選択不可
                                If InStr(info.Symbols(17), "A2") <> 0 Or
                                   InStr(info.Symbols(18), "I") <> 0 Or
                                   InStr(info.Symbols(18), "Y") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                        Case Else
                                            result = False
                                    End Select
                                End If
                                'N13-N11は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11
                                        result = False
                                End Select
                            Case "D"
                                'N13-N11以外は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11
                                    Case Else
                                        result = False
                                End Select
                        End Select
                    Case "SCS2"
                        Select Case info.Series.key_kataban
                            Case "", "B", "F"
                                'N3,N31,N2,N21は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N3, RodEndPatternDiv.N31, RodEndPatternDiv.N2,
                                        RodEndPatternDiv.N21
                                        result = False
                                End Select
                                ' オプション「A2」が選択されていた場合は「N13」「N15」以外は非表示
                                If InStr(info.Symbols(18), "A2") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                        Case Else
                                            result = False
                                    End Select
                                End If
                                ' 付属品「I」「Y」が選択されていた場合は「N13」「N15」以外は非表示
                                If InStr(info.Symbols(19), "I") <> 0 Or
                                   InStr(info.Symbols(19), "Y") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                        Case Else
                                            result = False
                                    End Select
                                End If
                                'N13-N11は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11
                                        result = False
                                End Select
                            Case "D", "G"
                                'N3,N31,N2,N21は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N3, RodEndPatternDiv.N31, RodEndPatternDiv.N2,
                                        RodEndPatternDiv.N21
                                        result = False
                                End Select
                                'N13-N11以外は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11
                                    Case Else
                                        result = False
                                End Select
                        End Select

                    Case "CMK2"

                        Select Case info.Series.key_kataban
                            Case ""
                                '付属品「I」「Y」を含む場合はN13/N15以外を選択不可
                                If InStr(info.Symbols(16), "I") <> 0 Or
                                   InStr(info.Symbols(16), "Y") <> 0 Then
                                    Select Case pattern
                                        Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                                        Case Else
                                            result = False
                                    End Select
                                End If
                                'N13-N11,N11-N13は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11, RodEndPatternDiv.N11N13
                                        result = False
                                End Select
                            Case "D"
                                'N13-N11以外は選択不可
                                Select Case pattern
                                    Case RodEndPatternDiv.N13N11
                                    Case RodEndPatternDiv.N11N13
                                        'バリエーションに「Q」を含む場合(この場合「DQ」のみ)は「N11-N13」は選択可
                                        If InStr(info.Symbols(1), "Q") = 0 Then
                                            result = False
                                        End If
                                    Case Else
                                        result = False
                                End Select
                        End Select
                End Select
            End If

            Return result
        End Function

        ''' <summary>
        ''' WF最大値を取得
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <param name="boreSize">口径</param>
        ''' <returns></returns>
        Public Shared Function GetWfMaxValue(series As String, keyKataban As String, boreSize As Integer) As Double
            Using client As New DbAccessServiceClient

                'WF最大値を取得
                Dim wfMaxValue = client.SelectWfMaxValue(series, keyKataban, boreSize)

                If String.IsNullOrEmpty(wfMaxValue) Then
                    Return 0
                Else
                    Return CType(wfMaxValue, Double)
                End If

            End Using
        End Function

        ''' <summary>
        '''     ロッド先端メッセージ１表示判断
        ''' </summary>
        ''' <param name="pattern"></param>
        ''' <param name="series"></param>
        ''' <returns></returns>
        Public Shared Function IsRodEndUnitShowMessage(pattern As String, series As String) As Boolean
            Select Case series
                Case "SSD", "SCA2", "CMK2"
                    'メッセージ表示
                    Return pattern = RodEndPatternDiv.N13N11 Or pattern = RodEndPatternDiv.N11N13
                Case Else
                    Return False
            End Select
        End Function

#Region "検証"

#End Region
    End Class
End Namespace