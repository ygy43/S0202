'************************************************************************************
'*  ProgramID  �FKHPriceQ9
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/12/18   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FEXA�V���[�Y  (���k��C�p�@�p�C���b�g���Q�|�[�g�d���� ���`�G�A�u���[�o���u)
'*             �FGEXA�V���[�Y (���k��C�p�@�p�C���b�g���Q�|�[�g�d���� �}�j�z�[���h)
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceQ9

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If selectedData.Series.series_kataban.Trim <> "GEXA" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '�V�[���ގ����Z
                If selectedData.Symbols(2).Trim <> "0" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2)
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

                '�R�C���I�v�V�������Z
                '�P�[�u������
                Select Case selectedData.Series.key_kataban.Trim
                    'RM1612033 ���i�ǉ��Ή��̂��߁Acase����Ɂu3�v��ǉ�  2016/12/19 �ǉ� ����
                    Case "1", "F", "3"
                        Select Case selectedData.Symbols(3).Trim
                            Case ""
                            Case "2C"
                                If selectedData.Symbols(5).Trim.Equals("1") Then
                                    'AC100V�̂݉��Z
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V �͉��Z����
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                        '���̑��I�v�V�������Z
                        If Not selectedData.Symbols(4).Trim.Equals("") Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "2"
                        Select Case selectedData.Symbols(3).Trim
                            Case ""
                            Case "2C"
                                If selectedData.Symbols(4).Trim.Equals("1") Then
                                    'AC100V�̂݉��Z
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V �͉��Z����
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select

                ''���̑��I�v�V�������Z
                'If Not selectedData.Symbols(4).Trim.Equals("") Then
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                '                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                '                                               selectedData.Symbols(4)
                '    decOpAmount(UBound(decOpAmount)) = 1
                'End If

                '�H�i�����H������
                If selectedData.Series.key_kataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(6)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

            End If

            If selectedData.Series.series_kataban.Trim = "GEXA" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1) & _
                                                           selectedData.Symbols(2) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4)

                decOpAmount(UBound(decOpAmount)) = 1

                If selectedData.Symbols(5).Trim = "2C" Then
                    If selectedData.Symbols(6).Trim = "1" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5)
                        decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(3)
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5)
                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(3)
                End If

            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
