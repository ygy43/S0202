'************************************************************************************
'*  ProgramID  �FKHPriceQ8
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/08/11   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FESSD,ELCR�V���[�Y  (�d���A�N�`���G�[�^)
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceQ8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            'RM1312084 2013/12/25
            'RM1402099 2014/02/25 ETS�V���[�Y�ǉ�
            Select Case selectedData.Series.series_kataban.Trim

                Case "ETV"
                    'RM1410045
                    If selectedData.Series.key_kataban = "T" Or selectedData.Series.key_kataban = "U" Or _
                        selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                        'TOYO�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "T" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        If selectedData.Symbols(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                        "T" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(10) <> "N" And _
                            selectedData.Symbols(10) <> "D" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                        "T" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                            If selectedData.Symbols(1) <> "05" And _
                               selectedData.Symbols(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                            '                                           selectedData.Symbols(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        If selectedData.Symbols(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(10) <> "N" And _
                            selectedData.Symbols(10) <> "D" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                            If selectedData.Symbols(1) <> "05" And _
                               selectedData.Symbols(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                            '                                           selectedData.Symbols(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ECS"
                    If selectedData.Series.key_kataban = "T" Or selectedData.Series.key_kataban = "U" Or _
                        selectedData.Series.key_kataban = "V" Or selectedData.Series.key_kataban = "W" Or _
                        selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                        'TOYO�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "T" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "T" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If selectedData.Symbols(8) <> "N" Then
                            '���_�Z���T
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       "T" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(10) <> "N" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       "T" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                            If selectedData.Symbols(1) <> "05" And _
                               selectedData.Symbols(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "V" Or selectedData.Series.key_kataban = "W" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If selectedData.Symbols(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If selectedData.Symbols(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                            '�H�i
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If selectedData.Symbols(4) <> "E" And _
                            selectedData.Symbols(4) <> "B" Then
                            '���[�^��t���@
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(8) <> "N" Then
                            '���_�Z���T
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(10) <> "N" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Symbols(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                            If selectedData.Symbols(1) <> "05" And _
                               selectedData.Symbols(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "4" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If selectedData.Symbols(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If selectedData.Symbols(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ETS"
                    'RM1402053
                    If selectedData.Series.key_kataban = "T" Or selectedData.Series.key_kataban = "U" Or _
                        selectedData.Series.key_kataban = "V" Or selectedData.Series.key_kataban = "W" Or _
                         selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@�ƃX�g���[�N�Ō��Z(�{�f�B�T�C�Y�F13,14,17)
                        If selectedData.Symbols(1) = "13" Or _
                            selectedData.Symbols(1) = "14" Then

                            If selectedData.Symbols(4) = "D" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If
                        If selectedData.Symbols(1) = "17" Then

                            If selectedData.Symbols(4) = "D" Or _
                            selectedData.Symbols(4) = "R" Or _
                            selectedData.Symbols(4) = "L" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If

                        '���_�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '�O���[�X�j�b�v��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                        '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                        If selectedData.Symbols(1) <> "05" And _
                           selectedData.Symbols(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Series.key_kataban = "V" Or selectedData.Series.key_kataban = "W" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If selectedData.Symbols(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            If selectedData.Symbols(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                            '�H�i
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    ElseIf selectedData.Series.key_kataban = "A" Or selectedData.Series.key_kataban = "B" Or _
                           selectedData.Series.key_kataban = "C" Or selectedData.Series.key_kataban = "D" Then
                        '���{Multi Axis�V���[�Y
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & Left(selectedData.Symbols(2), 1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3) & selectedData.Symbols(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�E���~�b�g�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7)
                        decOpAmount(UBound(decOpAmount)) = 1

                    ElseIf selectedData.Series.key_kataban = "I" Or selectedData.Series.key_kataban = "J" Or _
                     selectedData.Series.key_kataban = "K" Or selectedData.Series.key_kataban = "L" Or _
                     selectedData.Series.key_kataban = "M" Or selectedData.Series.key_kataban = "N" Or _
                     selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "P" Then
                        '���{Multi Axis�V���[�Y
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & Left(selectedData.Symbols(2), 1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3) & selectedData.Symbols(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�E���~�b�g�Z���T
                        If selectedData.Symbols(1).Trim = "210" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7) & MyControlChars.Hyphen & "210"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(7)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�W���i
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '�O���[�X�j�b�v��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                        '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                        If selectedData.Symbols(1) <> "05" And _
                           selectedData.Symbols(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If selectedData.Series.key_kataban = "4" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If selectedData.Symbols(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If selectedData.Symbols(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If selectedData.Series.key_kataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        'RM1802016  ���^���X�g�p�d�l�ǉ�
                        '�{�f�B�T�C�Y���u12�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                        If selectedData.Series.key_kataban = "Z" Then
                            If selectedData.Symbols(1) = "12" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12) & "-12"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(12) & "-06"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                Case "ECV"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '���_�Z���T
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�O���[�X�j�b�v��
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(10)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '�ΏۊO�ƂȂ��Ă���L�[�^�Ԃɂ��Ă�ΏۂƂȂ邽�ߏC��  2017/03/22 �C��
                    'If selectedData.Series.key_kataban = "T" Or selectedData.Series.key_kataban = "U" _
                    '   Or selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                    
                    '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                    '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ������
                    If selectedData.Symbols(1) <> "05" And _
                           selectedData.Symbols(1) <> "06" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11) & "-10"
                        decOpAmount(UBound(decOpAmount)) = 1

                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(11) & "-05"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    
                    If selectedData.Series.key_kataban = "F" Or _
                         selectedData.Series.key_kataban = "X" Or selectedData.Series.key_kataban = "Y" Then
                        '�H�i
                        '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(12)
                        'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        '                                           selectedData.Symbols(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESM"
                    Select Case selectedData.Symbols(1)
                        Case "HDU", "TTU", "CA", "SE", "PP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "VC"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "ST"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(selectedData.Symbols(2) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                       intST
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "B"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(selectedData.Symbols(3) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1)
                            decOpAmount(UBound(decOpAmount)) = intST
                    End Select

                Case "ERL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & "S-" & _
                                                               selectedData.Symbols(4) & _
                                                               selectedData.Symbols(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ERL2"
                    '��{���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & "E-" & _
                                                               selectedData.Symbols(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t�������Z�L�[
                    If selectedData.Symbols(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                             selectedData.Symbols(1) & selectedData.Symbols(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�u���[�L���Z���i�L�[
                    If selectedData.Symbols(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If selectedData.Symbols(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(9).Trim = "N" Then
                    Else

                        Select Case selectedData.Symbols(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                    End If

                Case "ESD"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & "S-" & _
                                                               selectedData.Symbols(4) & _
                                                               selectedData.Symbols(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(10).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(10)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESD2"
                    '��{���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & "E-" & _
                                                               selectedData.Symbols(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t�������Z�L�[
                    If selectedData.Symbols(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                             selectedData.Symbols(1) & selectedData.Symbols(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�u���[�L���Z���i�L�[
                    If selectedData.Symbols(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If selectedData.Symbols(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If selectedData.Symbols(9).Trim = "N" Then
                    Else

                        Select Case selectedData.Symbols(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                           MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1

                        End Select

                    End If

                    If selectedData.Symbols(10).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(10) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'RM1803042_EBS�EEBR�ǉ�
                Case "EBS"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(4) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "EBR"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(11)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1804032_EKS�ǉ�
                Case "EKS"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1) & _
                                                                selectedData.Symbols(2) & _
                                                                selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t���@
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                                selectedData.Symbols(8) & _
                                                                selectedData.Symbols(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    If selectedData.Symbols(5).Trim = "005" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                    selectedData.Symbols(5) & MyControlChars.Hyphen & _
                                                                    selectedData.Symbols(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                    selectedData.Symbols(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & "BASE" & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   MyControlChars.Hyphen & "OP" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(5)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

