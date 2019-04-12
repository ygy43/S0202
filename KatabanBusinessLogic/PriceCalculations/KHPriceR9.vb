'************************************************************************************
'*  ProgramID  �FKHPriceR9
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2012/09/27   �쐬�ҁFY.Tachi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�y���V���V�����_�@�r�b�o���R
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceR9

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intIndex As Integer
        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            ''�X�g���[�N�擾
            'intStroke = KHKataban.fncGetStrokeSize(selectedData.Series.series_kataban, _
            '                                      selectedData.Series.key_kataban, _
            '                                      CInt(selectedData.Symbols(2).Trim), _
            '                                      CInt(selectedData.Symbols(3).Trim))

            ''��{���i�L�[
            'Select Case True
            '    Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "D" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "K" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "M" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "O" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "T" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "V" Or _
            '         Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "Z"
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
            '                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            '    Case Mid(selectedData.Series.series_kataban.Trim, 5, 1) = ""
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
            '                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            '    Case Else
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
            '                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            'End Select

            '��{���i�L�[
            Select Case True
                Case CInt(selectedData.Symbols(3).Trim) <= 10
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                    'RM1305005 2013/06/14 �C��
                    Select Case selectedData.Symbols(2).Trim
                        Case "6"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-105"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case CInt(selectedData.Symbols(3).Trim) >= 106 And CInt(selectedData.Symbols(3).Trim) <= 120
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 121 And CInt(selectedData.Symbols(3).Trim) <= 135
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-135"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 136 And CInt(selectedData.Symbols(3).Trim) <= 150
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-150"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 151 And CInt(selectedData.Symbols(3).Trim) <= 165
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-165"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 166 And CInt(selectedData.Symbols(3).Trim) <= 180
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-180"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 181 And CInt(selectedData.Symbols(3).Trim) <= 195
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-195"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 196 And CInt(selectedData.Symbols(3).Trim) <= 210
                    'RM1305005 2013/06/14 �C��
                    Select Case selectedData.Symbols(2).Trim
                        Case "10"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-210"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case CInt(selectedData.Symbols(3).Trim) >= 211 And CInt(selectedData.Symbols(3).Trim) <= 225
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-225"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 226 And CInt(selectedData.Symbols(3).Trim) <= 240
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-240"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 241 And CInt(selectedData.Symbols(3).Trim) <= 255
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-255"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 256 And CInt(selectedData.Symbols(3).Trim) <= 260
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-260"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(selectedData.Symbols(3).Trim) >= 261
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-270"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select


            'ϸ�ȯĉ��Z���i�L�[
            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3-L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�x���`�����Z���i�L�[
            If Left(selectedData.Symbols(1).Trim, 2) = "CB" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "FA" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "LB" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "LS" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                intIndex = 5
                If selectedData.Symbols(intIndex).Trim <> "" Then
                    '�X�C�b�`���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(intIndex).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)

                    If selectedData.Symbols(intIndex + 1).Trim <> "" Then
                        '���[�h���������Z���i�L�[
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(intIndex + 1).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)
                    End If
                End If
                intIndex = intIndex + 2
            End If

            '�I�v�V�����E�t���i���Z���i�L�[
            intIndex = 0
            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                If Len(selectedData.Symbols(8)) <> 0 Then
                    intIndex = 8
                End If
            Else
                If Len(selectedData.Symbols(5)) <> 0 Then
                    intIndex = 5
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(selectedData.Symbols(intIndex), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    '�H�i�����H���������i
                    Select Case selectedData.Series.key_kataban
                        Case "F"
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End Select
                Next
            End If

            '�I�v�V�����E�t���i���Z���i�L�[
            intIndex = 0
            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                If selectedData.Series.key_kataban.Trim = "F" Then
                    If Len(selectedData.Symbols(9)) <> 0 Then
                        intIndex = 9
                    End If
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(selectedData.Symbols(intIndex), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

            '�񎟓d�r���Z���i�L�[
            intIndex = 0
            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                If selectedData.Series.key_kataban.Trim = "4" Then
                    If Len(selectedData.Symbols(9)) <> 0 Then
                        intIndex = 9
                    End If
                End If
            Else
                If Len(selectedData.Symbols(6)) <> 0 Then
                    intIndex = 6
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(selectedData.Symbols(intIndex), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
