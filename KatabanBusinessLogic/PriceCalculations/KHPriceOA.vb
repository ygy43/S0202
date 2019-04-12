Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceOA
    '************************************************************************************
    '*  ProgramID  �FKHPriceOA (KHPriceO0�����)
    '*  Program��  �F�P���v�Z�T�u���W���[��
    '*
    '*                                      �쐬���F2007/02/27   �쐬�ҁFNII K.Sudoh
    '*                                      �X�V���F             �X�V�ҁF
    '*
    '*  �T�v       �F�������^�C�v�@���j�A�X���C�h�V�����_�@�k�b�q�^�k�b�q�|�p
    '*
    '*  �X�V����   �F                        
    '*�@�E��tNo�FRM1002067  KHPriceO0�@����@�k�b�q�^�k�b�q�|�p�𕪗�
    '*                      �o���G�[�V�����A�I�v�V�����ǉ��Ή�
    '*                                      �X�V���F2010/04/07   �X�V�ҁFY.Miura
    '************************************************************************************
    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim bolC5Flag As Boolean
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionP4 As Boolean = False
        Dim strOptionP4 As String = String.Empty

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '�I�v�V�������Z���i�L�[
            strOpArray = Split(selectedData.Symbols(12), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                        strOptionP4 = strOpArray(intLoopCnt).Trim
                End Select
            Next

            'C5�`�F�b�N
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData)

            '�X�g���[�N�ݒ�
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(4).Trim))

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            '�o���G�[�V�������Z���i�L�[
            Select Case selectedData.Series.series_kataban.Trim
                Case "LCG-Q", "LCR-Q"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If
            End Select

            'RM1003086 2010/04/07 Y.Miura �ǉ�
            If selectedData.Symbols(1).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAR-" & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If
            End If

            '�X�C�b�`���Z���i�L�[
            If selectedData.Symbols(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)

                '���[�h���������Z���i�L�[
                If selectedData.Symbols(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                End If

                'RM0906034 2009/08/18 Y.Miura�@�񎟓d�r�Ή�
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                End If
            End If


            '�I�v�V����(1)���Z���i�L�[
            If selectedData.Symbols(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                Select Case selectedData.Symbols(8).Trim
                    Case "W1", "W2", "W3", "W4", "W5", "W6"
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                    Case Else
                        If selectedData.Symbols(10).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8).Trim & _
                                                                       selectedData.Symbols(10).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                        End If
                End Select
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

                '�񎟓d�r�Ή�
                If bolOptionP4 Then
                    Select Case selectedData.Symbols(8).Trim
                        Case "A1", "A2", "A3", "A4" '�V���b�N�L���[�t
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                      "A" & MyControlChars.Hyphen & _
                                                                      strOptionP4 & MyControlChars.Hyphen & _
                                                                      selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "A5", "A6"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                      "A" & MyControlChars.Hyphen & _
                                                                      strOptionP4 & MyControlChars.Hyphen & _
                                                                      selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
                End If
            End If


            '�X�g���[�N�����͈͉��Z
            If selectedData.Symbols(9).Trim <> "" Then
                'If selectedData.Symbols(8).ToString.PadRight(1, " ").Substring(0, 1) = "C" And selectedData.Symbols(9).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-STR-" & _
                                                          selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                          selectedData.Symbols(9).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                'Else
                'End If
            End If

            '�I�v�V����(3)���Z���i�L�[
            If selectedData.Symbols(11).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                Select Case selectedData.Symbols(8).Trim
                    Case "W1", "W2", "W3", "W4", "W5", "W6"
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).ToString.PadRight(1, " ").Substring(0, 1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                    Case Else
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                End Select
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

            End If

            '�N���[���d�l���Z���i�L�[
            strOpArray = Split(selectedData.Symbols(12), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "U"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    Case "P4", "P40"    '�񎟓d�r
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "P72", "P73"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
