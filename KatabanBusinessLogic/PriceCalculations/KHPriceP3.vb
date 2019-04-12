'************************************************************************************
'*  ProgramID  �FKHPriceP3
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/12/18   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F�G�A���j�b�g�@�b�w�t�V���[�Y
'*�@�ύX       �FRM0911XXX 2009/11/11 Y.Miura �@��ǉ�
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP3

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer
        Dim strOpArray() As String

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Dim fullKataban = PriceManager.GetFullKataban(selectedData)

            Select Case True
                Case InStr(fullKataban, "CXU10-GFAB3") <> 0 Or _
                     InStr(fullKataban, "CXU30-GFAB4U") <> 0
                    '��{���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�@��I�v�V�������i�L�[
                    strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next

                    '�R�C���I�v�V�������i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case InStr(fullKataban, "CXU30-M4G2") <> 0
                    '��{���i�L�[
                    If Len(selectedData.Symbols(4).Trim) = 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   "M4G2-2"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   "M4G2-4"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�d���ڑ�(1�`4�A)�L�[
                    For intLoopCnt = 2 To 5
                        If Len(selectedData.Symbols(intLoopCnt).Trim) <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If Len(selectedData.Symbols(8).Trim) <> 0 And selectedData.Symbols(intLoopCnt).Trim <> "0" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           Mid(selectedData.Symbols(1).Trim, 2, 3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(8).Trim
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           Mid(selectedData.Symbols(1).Trim, 2, 3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(intLoopCnt).Trim
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Next

                    '�ڑ����a���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�@��I�v�V�������i�L�[
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next
                Case InStr(fullKataban, "CXU10-FAB3") <> 0 Or _
                     InStr(fullKataban, "CXU30-FAB4U") <> 0
                    '��{���i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    '�R�C���I�v�V�������i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case InStr(fullKataban, "CXU30-FAD") <> 0
                    '��{���i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                  selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                  selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                  selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�R�C���I�v�V�������i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case InStr(fullKataban, "CXU30-4G2") <> 0
                    '��{���i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    '�d���ڑ�(1�`2�A)�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        For intLoopCnt = 3 To 4
                            If Len(selectedData.Symbols(intLoopCnt).Trim) <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                If Len(selectedData.Symbols(6).Trim) <> 0 And selectedData.Symbols(intLoopCnt).Trim <> "0" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(6).Trim
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(intLoopCnt).Trim
                                End If
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Next
                    Else
                        For intLoopCnt = 2 To 3
                            If Len(selectedData.Symbols(intLoopCnt).Trim) <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                If Len(selectedData.Symbols(5).Trim) <> 0 And selectedData.Symbols(intLoopCnt).Trim <> "0" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(5).Trim
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(intLoopCnt).Trim
                                End If
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Next
                    End If

                    '�ڑ����a���i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�T�C�����T���i�L�[
                    If selectedData.Symbols(2).Trim = "UN" Then
                        If Len(selectedData.Symbols(7).Trim) <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        If Len(selectedData.Symbols(6).Trim) <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(6).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                    'RM0911XXX 2009/11/11 Y.Miura �@��ǉ�
                Case InStr(fullKataban, "CXU30-ADK") <> 0
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 15)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If selectedData.Symbols(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If selectedData.Symbols(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 12)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If selectedData.Symbols(3).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If selectedData.Symbols(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                    'Case InStr(fullKataban, "CXU10-CHV") <> 0
                Case InStr(fullKataban, "CXU10-CHV-00") <> 0
                    If selectedData.Symbols(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 12) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 12)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(4).Trim <> "" Then     '�I�v�V�����L��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 12) & "-OP-" & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case InStr(fullKataban, "CXU10-CHV-UN") <> 0
                    If selectedData.Symbols(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 15) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 15)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(5).Trim <> "" Then     '�I�v�V�����L��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & "-00-OP-" & _
                                                                   selectedData.Symbols(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(6).Trim <> "" Then     '�I�v�V�����L��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 9) & "-00-OP-" & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    'RM1003086 2010/03/26 Y.Miura �@��ǉ�
                Case InStr(fullKataban, "CXU10-EXA") <> 0
                    '��{���i
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    '�R�C���I�v�V�������i
                    If selectedData.Symbols(2).Trim = "UN" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        Select Case selectedData.Symbols(4)
                            Case "2C"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim
                            Case Else
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                          selectedData.Symbols(4).Trim
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        Select Case selectedData.Symbols(3)
                            Case "2C"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                            Case Else
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(fullKataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                          selectedData.Symbols(3).Trim
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case InStr(fullKataban, "CXU10-GEXA") <> 0
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�@��I�v�V�������i�L�[
                    strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next

                    '�R�C���I�v�V�������i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case selectedData.Symbols(5)
                        Case "2C"
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(6).Trim
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                    '��RM1502045
                Case InStr(fullKataban, "CXU30-VE") <> 0
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���̓X�C�b�`�I�v�V�������i�L�[
                    If selectedData.Symbols(3).Trim.Length > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '��������I�v�V�������i�L�[
                    If selectedData.Symbols(3).Trim.Length > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '�d���ڑ��I�v�V�������Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�T�C�����T�I�v�V�������i�L�[
                    If selectedData.Symbols(3).Trim.Length > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '�d���I�v�V�������i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '��RM1306001 2013/06/05 �ǉ�
                    '��RM1711*** 2017/11/20 �ύX
                Case InStr(fullKataban, "WXU-H") <> 0
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���ʃZ���T�o�͉��Z(�A�������Z)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim & _
                                                               selectedData.Symbols(6).Trim

                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(1).Trim
                Case InStr(fullKataban, "WXU-J") <> 0
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���ʃZ���T�o�͉��Z(�A�������Z)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim & _
                                                               selectedData.Symbols(8).Trim

                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(1).Trim
                Case InStr(fullKataban, "WXU-P") <> 0
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���ʃZ���T�o�͉��Z(�A�������Z)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10).Trim & _
                                                               selectedData.Symbols(11).Trim

                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(2).Trim
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
