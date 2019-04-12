'************************************************************************************
'*  ProgramID  �FKHPriceP5
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/01/09   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F�X�[�p�[�R���p�N�g�V�����_�@�r�r�c�Q
'*
'*�y�C������z
'*                                      �X�V���F2008/05/07   �X�V�ҁFT.Sato
'*  �E��tNo�FRM0802088�Ή��@�o���G�[�V�����i'�c','�l','�p','�w','�x'�j�ǉ��ɔ����C��
'* �@�@�@�@�@�@�@�@�@�@�@�@�@���Ɂi'�p'�j�̓{�b�N�X���P�����_��l�����ďC��
'*  �E��tNo�FRM0906034  �񎟓d�r�Ή��@��@SSD2
'*                                      �X�V���F2009/08/04   �X�V�ҁFY.Miura
'*  �E��tNo�FRM1001043  �񎟓d�r�Ή��@�� �`�F�b�N�敪�ύX 3��2�@
'*                                      �X�V���F2010/02/22   �X�V�ҁFY.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP5

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim intStrokeS1 As Integer = 0      'RM1010017 ADD 
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean
        Dim intOpAmount As Integer
        Dim intOpAmountBW As Integer
        Dim bolOpP4 As Boolean              'RM0906034 2009/08/04 Y.Miura�@�񎟓d�r�Ή�

        Dim strVariation As String          '�o���G�[�V����
        Dim strSwitchAttached As String     '�X�C�b�`
        Dim strBoreSize As String           '���a
        Dim strCushion As String            '�z�ǂ˂��A�N�b�V����
        Dim strStroke As String             '�X�g���[�N
        Dim strPositionLocking As String    '�����h�~�ʒu
        Dim strSwitchModel As String        '�X�C�b�`
        Dim strLeadWireLen As String        '���[�h������
        Dim strSwitchQty As String          '��
        Dim strLod As String                '���b�h��[

        '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
        Dim strStrokeS1 As String           '�X�g���[�N(S1)
        Dim strPositionLockingS1 As String  '�����h�~�ʒu(S1)
        Dim strSwitchModelS1 As String      '�X�C�b�`(S1)
        Dim strLeadWireLenS1 As String      '���[�h������(S1)
        Dim strSwitchQtyS1 As String        '��(S1)
        '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

        Dim strOption As String             '�I�v�V����
        Dim strFP1 As String                '�H�i��������
        Dim strMountingBracket As String    '�x�����
        Dim strAccessory As String          '�t���i

        Try



            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case selectedData.Series.key_kataban
                '2010/11/01 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                Case "", "K", "L", "4"
                    ''2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                    'Case ""
                    '2010/11/01 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                    strVariation = selectedData.Symbols(1).Trim         '�o���G�[�V�����@
                    strSwitchAttached = selectedData.Symbols(2).Trim    '�o���G�[�V�����A(�X�C�b�`)
                    strBoreSize = selectedData.Symbols(4).Trim          '���a
                    strCushion = selectedData.Symbols(5).Trim           '�z�ǂ˂��A�N�b�V����
                    strStrokeS1 = selectedData.Symbols(7).Trim          '�r�P�F�X�g���[�N
                    strPositionLockingS1 = selectedData.Symbols(8).Trim '�r�P�F�����h�~�ʒu
                    strSwitchModelS1 = selectedData.Symbols(9).Trim     '�r�P�F�X�C�b�`
                    strLeadWireLenS1 = selectedData.Symbols(10).Trim    '�r�P�F���[�h������
                    strSwitchQtyS1 = selectedData.Symbols(11).Trim      '�r�P�F��
                    strLod = selectedData.Symbols(12).Trim              '�r�P�F���b�h��[
                    strStroke = selectedData.Symbols(14).Trim           '�r�Q�F�X�g���[�N
                    strPositionLocking = selectedData.Symbols(15).Trim  '�r�Q�F�����h�~�ʒu
                    strSwitchModel = selectedData.Symbols(16).Trim      '�r�Q�F�X�C�b�`
                    strLeadWireLen = selectedData.Symbols(17).Trim      '�r�Q�F���[�h������
                    strSwitchQty = selectedData.Symbols(18).Trim        '�r�Q�F��
                    strOption = selectedData.Symbols(19).Trim           '�I�v�V����
                    strMountingBracket = selectedData.Symbols(20).Trim  '�x�����
                    strAccessory = selectedData.Symbols(21).Trim        '�t���i
                    strFP1 = ""
                    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
                    '2010/11/01 DEL RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    'Case "Q"
                    '    strVariation = selectedData.Symbols(1).Trim        '�o���G�[�V����
                    '    strSwitchAttached = selectedData.Symbols(2).Trim   '�X�C�b�`
                    '    strBoreSize = selectedData.Symbols(3).Trim         '���a
                    '    strCushion = selectedData.Symbols(4).Trim          '�z�ǂ˂��A�N�b�V����
                    '    strStroke = selectedData.Symbols(5).Trim           '�X�g���[�N
                    '    strPositionLocking = selectedData.Symbols(6).Trim  '�����h�~�ʒu
                    '    strSwitchModel = selectedData.Symbols(7).Trim      '�X�C�b�`
                    '    strLeadWireLen = selectedData.Symbols(8).Trim      '���[�h������
                    '    strSwitchQty = selectedData.Symbols(9).Trim        '��
                    '    strOption = selectedData.Symbols(10).Trim          '�I�v�V����
                    '    strMountingBracket = selectedData.Symbols(11).Trim '�x�����
                    '    strAccessory = selectedData.Symbols(12).Trim       '�t���i
                    '2010/11/01 DEL RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                Case "7", "N"
                    strVariation = selectedData.Symbols(1).Trim         '�o���G�[�V�����@
                    strSwitchAttached = selectedData.Symbols(2).Trim    '�o���G�[�V�����A(�X�C�b�`)
                    strBoreSize = selectedData.Symbols(4).Trim          '���a
                    strCushion = selectedData.Symbols(5).Trim           '�z�ǂ˂��A�N�b�V����
                    strStrokeS1 = selectedData.Symbols(7).Trim          '�r�P�F�X�g���[�N
                    strPositionLockingS1 = selectedData.Symbols(8).Trim '�r�P�F�����h�~�ʒu
                    strSwitchModelS1 = selectedData.Symbols(9).Trim     '�r�P�F�X�C�b�`
                    strLeadWireLenS1 = selectedData.Symbols(10).Trim    '�r�P�F���[�h������
                    strSwitchQtyS1 = selectedData.Symbols(11).Trim      '�r�P�F��
                    strLod = selectedData.Symbols(12).Trim              '�r�P�F���b�h��[
                    strStroke = selectedData.Symbols(14).Trim           '�r�Q�F�X�g���[�N
                    strPositionLocking = selectedData.Symbols(15).Trim  '�r�Q�F�����h�~�ʒu
                    strSwitchModel = selectedData.Symbols(16).Trim      '�r�Q�F�X�C�b�`
                    strLeadWireLen = selectedData.Symbols(17).Trim      '�r�Q�F���[�h������
                    strSwitchQty = selectedData.Symbols(18).Trim        '�r�Q�F��
                    strOption = selectedData.Symbols(19).Trim           '�I�v�V����
                    strFP1 = selectedData.Symbols(20).Trim              '�H�i��������
                    strMountingBracket = selectedData.Symbols(21).Trim  '�x�����
                    strAccessory = selectedData.Symbols(22).Trim        '�t���i
                Case "F"
                    strVariation = selectedData.Symbols(1).Trim        '�o���G�[�V����
                    strSwitchAttached = selectedData.Symbols(2).Trim   '�X�C�b�`
                    strBoreSize = selectedData.Symbols(3).Trim         '���a
                    strCushion = selectedData.Symbols(4).Trim          '�z�ǂ˂��A�N�b�V����
                    strStroke = selectedData.Symbols(5).Trim           '�X�g���[�N
                    strPositionLocking = ""                                             '�����h�~�ʒu
                    strSwitchModel = selectedData.Symbols(6).Trim      '�X�C�b�`
                    strLeadWireLen = selectedData.Symbols(7).Trim      '���[�h������
                    strSwitchQty = selectedData.Symbols(8).Trim        '��
                    strOption = selectedData.Symbols(9).Trim           '�I�v�V����
                    strFP1 = selectedData.Symbols(10).Trim             '�H�i��������
                    strMountingBracket = selectedData.Symbols(11).Trim '�x�����
                    strAccessory = selectedData.Symbols(12).Trim
                Case Else
                    strVariation = selectedData.Symbols(1).Trim        '�o���G�[�V����
                    strSwitchAttached = selectedData.Symbols(2).Trim   '�X�C�b�`
                    strBoreSize = selectedData.Symbols(3).Trim         '���a
                    strCushion = selectedData.Symbols(4).Trim          '�z�ǂ˂��A�N�b�V����
                    strStroke = selectedData.Symbols(5).Trim           '�X�g���[�N
                    strPositionLocking = ""                                             '�����h�~�ʒu
                    strSwitchModel = selectedData.Symbols(6).Trim      '�X�C�b�`
                    strLeadWireLen = selectedData.Symbols(7).Trim      '���[�h������
                    strSwitchQty = selectedData.Symbols(8).Trim        '��
                    strOption = selectedData.Symbols(9).Trim          '�I�v�V����
                    strMountingBracket = selectedData.Symbols(10).Trim '�x�����
                    strAccessory = selectedData.Symbols(11).Trim       '�t���i
                    strFP1 = ""
            End Select

            'RM0906034 2009/08/04 Y.Miura�@�ǉ�����
            '�I�v�V�������񎟓d�r�Ή������f����
            bolOpP4 = False
            strOpArray = Split(strOption, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOpP4 = True
                End Select
            Next
            'RM0906034 2009/08/04 Y.Miura�@�ǉ�����

            '���ʐݒ�
            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
            intOpAmount = 1
            intOpAmountBW = 1
            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

            Select Case selectedData.Series.key_kataban
                Case "D", "E", "F"
                    intOpAmount = 2
                    '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                    '��2013/09/20 ���[�J���łƂ̍��ُC��
                Case "", "4", "7"

                    Select Case Left(strVariation.Trim, 1)
                        Case "B", "W"
                            intOpAmountBW = 2
                    End Select
                    'Case Else
                    '    intOpAmount = 1
                    '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
            End Select

            'C5�`�F�b�N
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)

            'C5�`�F�b�N
            Select Case selectedData.Series.key_kataban
                Case "L", "4", "6", "E", "7", "F", "N"
                    bolC5Flag = True
                    '��RM1306001 2013/06/06 �ǉ�
                Case "", "K"
                    If selectedData.Symbols(22).Trim = "SX" Then
                        bolC5Flag = True
                    End If
                Case "D"
                    If selectedData.Symbols(12).Trim = "SX" Then
                        bolC5Flag = True
                    End If
            End Select

            '�X�g���[�N�ݒ�
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(strBoreSize), _
                                                  CInt(strStroke))


            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            Select Case selectedData.Series.key_kataban
                '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                Case "", "4", "7"
                    Select Case Left(strVariation.Trim, 1)
                        Case "B", "W"
                            '�X�g���[�N�ݒ�(S1)
                            intStrokeS1 = KatabanUtility.GetStrokeSize(selectedData, _
                                                                    CInt(strBoreSize), _
                                                                    CInt(IIf(strStrokeS1.Equals(String.Empty), 0, strStrokeS1)))
                            'S1
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       "BASE" & MyControlChars.Hyphen & strBoreSize & MyControlChars.Hyphen & intStrokeS1.ToString

                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    End Select

                    'S2
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "BASE" & MyControlChars.Hyphen & strBoreSize & MyControlChars.Hyphen & intStroke.ToString
                    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
                Case "D", "K"
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "BASE" & MyControlChars.Hyphen & selectedData.Series.key_kataban & MyControlChars.Hyphen & _
                                                               strBoreSize & MyControlChars.Hyphen & intStroke.ToString
                Case "L", "N"        'RM0906034 2009/08/04 Y.Miura�@�ǉ�
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "BASE" & MyControlChars.Hyphen & "K" & MyControlChars.Hyphen & _
                                                               strBoreSize & MyControlChars.Hyphen & intStroke.ToString
                Case "E", "F"
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "BASE" & MyControlChars.Hyphen & "D" & MyControlChars.Hyphen & _
                                                               strBoreSize & MyControlChars.Hyphen & intStroke.ToString
                Case Else
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "BASE" & MyControlChars.Hyphen & strBoreSize & MyControlChars.Hyphen & intStroke.ToString
            End Select
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
            End If

            '�o���G�[�V�������Z���i�L�[
            Select Case selectedData.Series.key_kataban.Trim
                '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                Case "", "4", "7"
                    Select Case strVariation
                        '2010/11/01 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                        Case "T1", "T1L", "O", "B", "W", "G", "G1", "G4", "G5", "M", "Q"
                            'Case "T1", "T1L", "O", "B", "W", "G", "G1", "G4", "G5"
                            '2010/11/01 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & strVariation & MyControlChars.Hyphen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If

                        Case "G2", "G3"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & strVariation & MyControlChars.Hyphen & strBoreSize & _
                                                                MyControlChars.Hyphen & intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select
                    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
                    '2010/11/01 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                Case "K", "L"
                    Select Case strVariation
                        Case "KU", "KG5"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & strVariation & MyControlChars.Hyphen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case "KG1", "KG4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & Right(strVariation, 2) & MyControlChars.Hyphen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case "KG2", "KG3"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & strVariation & MyControlChars.Hyphen & strBoreSize & _
                                                                MyControlChars.Hyphen & intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select
                Case "D"
                    Select Case strVariation
                        Case "DG1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & Right(strVariation, 2) & MyControlChars.Hyphen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If

                        Case "DG4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-" & strVariation & MyControlChars.Hyphen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                        Case "DM"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                "-VAL-M-" & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                    End Select
                    '2010/11/01 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                    '2010/11/01 DEL RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    'Case "M"
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAL-M-" & strBoreSize
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    '    If bolC5Flag = True Then
                    '        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    '    End If
                    'Case "Q"
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-VAL-Q-" & strBoreSize
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    '    If bolC5Flag = True Then
                    '        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    '    End If
                    '2010/11/01 DEL RM1011020(12��VerUP:SSD2�V���[�Y) <---END
            End Select

            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
            '�o���G�[�V�����B
            '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            If (selectedData.Series.key_kataban.Trim = "" OrElse _
                selectedData.Series.key_kataban.Trim = "K" OrElse _
                selectedData.Series.key_kataban.Trim = "L" OrElse _
                selectedData.Series.key_kataban.Trim = "4") _
            AndAlso selectedData.Symbols(3).Trim = "F" Then
                'If selectedData.Series.key_kataban.Trim = "" AndAlso _
                'selectedData.Symbols(3).Trim = "F" Then
                '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                '���a����
                Select Case strBoreSize
                    Case "12", "16", "20"
                        '�X�g���[�N(S2)
                        Select Case True
                            Case strStroke <= 15
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-15"
                            Case strStroke >= 16 And strStroke <= 30
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-16-30"
                            Case strStroke >= 31
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-31-50"
                        End Select
                    Case "25", "32", "40", "50", "63", "80", "100"
                        '�X�g���[�N(S2)
                        Select Case True
                            Case strStroke <= 25
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-25"
                            Case strStroke >= 26 And strStroke <= 50
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-26-50"
                            Case strStroke >= 51 And strStroke <= 75
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-51-75"
                            Case strStroke >= 76
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-76-100"
                        End Select
                    Case "125", "140", "160"
                        '�X�g���[�N(S2)
                        Select Case True
                            Case strStroke <= 50
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-50"
                            Case strStroke >= 51 And strStroke <= 100
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-51-100"
                            Case strStroke >= 101 And strStroke <= 200
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-101-200"
                            Case strStroke >= 201
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                        "-F-" & strBoreSize & "-201-300"
                        End Select

                End Select

                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

            End If
            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

            '�X�C�b�`���Z���i�L�[
            Select Case selectedData.Series.key_kataban.Trim
                Case "", "K", "L", "4", "E", "6", "D", "2", "7", "F", "N"
                    If strSwitchAttached <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "SW" & MyControlChars.Hyphen & strSwitchAttached & MyControlChars.Hyphen & strBoreSize
                        '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                        decOpAmount(UBound(decOpAmount)) = intOpAmountBW
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                        End If
                    End If
            End Select

            '�N�b�V�������Z���i�L�[
            '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
            Select Case selectedData.Series.key_kataban.Trim
                Case "", "K", "L", "4", "E", "D", "7", "F", "N"
                    Select Case strCushion
                        Case "D", "GD", "ND"
                            'If strCushion <> "" Then
                            '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       "OP" & MyControlChars.Hyphen & Right(strCushion, 1) & MyControlChars.Hyphen & strBoreSize
                            'strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                            '                                           "OP" & MyControlChars.Hyphen & strCushion & MyControlChars.Hyphen & strBoreSize
                            '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                        Case "C", "GC", "NC"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim _
                                                                    & MyControlChars.Hyphen & "K-*C" & MyControlChars.Hyphen & strBoreSize

                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            End If
                            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                    End Select
            End Select

            '�X�C�b�`���Z���i�L�[
            '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            Select Case selectedData.Series.key_kataban.Trim
                Case "", "K", "L", "4", "7", "N"
                    ''2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                    'If selectedData.Series.key_kataban.Trim = "" Then
                    '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                    If strSwitchModelS1 <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   "SW" & MyControlChars.Hyphen & strSwitchModelS1
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQtyS1)

                        '��2013/09/20 ���[�J���łƍ��ُC��
                        If bolOpP4 Then  'P4
                            '�X�C�b�`���Z
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       "SW" & MyControlChars.Hyphen & "P4"
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQty)
                        End If
                    End If
            End Select
            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

            If strSwitchModel <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           "SW" & MyControlChars.Hyphen & strSwitchModel
                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQty)

                'RM0906034 2009/08/04 Y.Miura�@�񎟓d�r�Ή��ǉ�����
                If bolOpP4 Then  'P4
                    '�X�C�b�`���Z
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               "SW" & MyControlChars.Hyphen & "P4"
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQty)
                End If
                'RM0906034 2009/08/04 Y.Miura�@�񎟓d�r�Ή��ǉ�����
            End If

            '���[�h���������Z���i�L�[
            '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            Select Case selectedData.Series.key_kataban.Trim
                Case "", "K", "L", "4", "7", "N"
                    ''2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                    'If selectedData.Series.key_kataban.Trim = "" Then
                    '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                    If strSwitchModelS1 <> "" AndAlso strLeadWireLenS1 <> "" Then

                        '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                        Dim strKataban As String = ""
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                        Select Case strSwitchModelS1
                            'RM1307003 2013/07/04�ǉ�(F2S,F3S)
                            Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                 "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                 "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                 "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "SWLW(1)" & MyControlChars.Hyphen & strLeadWireLenS1
                            Case "T2YD"
                                strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "SWLW(2)" & MyControlChars.Hyphen & strLeadWireLenS1
                            Case "T2YDT"
                                strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "SWLW(3)" & MyControlChars.Hyphen & strLeadWireLenS1
                            Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "SWLW(7)" & MyControlChars.Hyphen & strLeadWireLenS1
                            Case "V0", "V7"
                                strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "SWLW(8)" & MyControlChars.Hyphen & strLeadWireLenS1
                        End Select

                        '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                        If strKataban.Trim.Length > 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban
                            decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQtyS1)

                        End If
                        '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                    End If
            End Select
            '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

            If strSwitchModel <> "" Then
                If strLeadWireLen <> "" Then
                    '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    Dim strKataban As String = ""
                    'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    Select Case selectedData.Series.key_kataban.Trim
                        Case "", "K"
                            ''2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                            'If selectedData.Series.key_kataban.Trim = "" Then
                            '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                            Select Case strSwitchModel
                                'RM1307003 2013/07/04�ǉ�(F2S,F3S)
                                Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                     "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                     "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                     "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(1)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "T2YD"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(2)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "T2YDT"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(3)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(7)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "V0", "V7"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(8)" & MyControlChars.Hyphen & strLeadWireLen
                            End Select

                        Case Else
                            Select Case strSwitchModel
                                'RM1307003 2013/07/04�ǉ�(F2S,F3S)
                                Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                     "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                     "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                     "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(1)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "T2YD"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(2)" & MyControlChars.Hyphen & strLeadWireLen
                                Case "T2YDT"
                                    strKataban = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               "SWLW(3)" & MyControlChars.Hyphen & strLeadWireLen
                            End Select
                    End Select

                    '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

                    '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    If strKataban.Trim.Length > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban

                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(strSwitchQty)
                    End If
                    '2010/11/17 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                End If
            End If

            '�I�v�V�������Z���i�L�[
            strOpArray = Split(strOption, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                Select Case selectedData.Series.key_kataban.Trim
                    Case "", "K", "L", "4", "7", "N"
                        ''2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                        'If selectedData.Series.key_kataban.Trim = "" Then
                        '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                            Case "M"
                                Select Case strBoreSize
                                    Case "12", "16", "20", "25"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   strBoreSize & MyControlChars.Hyphen & intStroke.ToString

                                        decOpAmount(UBound(decOpAmount)) = intOpAmountBW
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                        End If
                                    Case "32", "40", "50", "63", "80", "100", "125", "140", "160"
                                        Select Case Left(strVariation.Trim, 1)
                                            Case "B", "W"
                                                'S1
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                           strBoreSize & MyControlChars.Hyphen & intStrokeS1.ToString

                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                                End If

                                        End Select

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                                   strBoreSize & MyControlChars.Hyphen & intStroke.ToString

                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                        End If
                                End Select
                            Case "P5", "P51", "P7", "P71"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & Left(strOpArray(intLoopCnt).Trim, 2) & "*" & _
                                                                           MyControlChars.Hyphen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If

                                '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                            Case "M0", "M1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                                '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                                '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) START--->
                            Case "S"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If

                                '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) <---END
                            Case "P4", "P40"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount

                        End Select

                    Case Else
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                'RM0912XXX 2009/12/09 Y.Miura�@�񎟓d�rC5���Z�s�v
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                'End If
                                'RM0906034 2009/08/04 Y.Miura�@�񎟓d�r�Ή��ǉ�����
                            Case "P4", "P40"
                                Select Case selectedData.Series.key_kataban
                                    Case "E"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   "OP-D" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & strBoreSize
                                        decOpAmount(UBound(decOpAmount)) = intOpAmount
                                End Select
                                'RM0912XXX 2009/12/09 Y.Miura�@�񎟓d�rC5���Z�s�v
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                'End If
                                'RM0906034 2009/08/04 Y.Miura�@�񎟓d�r�Ή��ǉ�����
                            Case "M"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                           strBoreSize & MyControlChars.Hyphen & intStroke.ToString
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                            Case "M0", "M1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           "OP" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                                End If
                        End Select
                End Select
                '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
            Next

            'FP���Z���i�L�[
            If strFP1 <> "" Then

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                           strFP1 & MyControlChars.Hyphen & strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                End If

            End If

            '�x�������Z���i�L�[
            If strMountingBracket <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strMountingBracket & MyControlChars.Hyphen & strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                'RM0912XXX 2009/12/09 Y.Miura�@�񎟓d�rC5���Z�s�v
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                'End If
            End If

            '�t���i���Z���i�L�[
            If strAccessory <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strAccessory & MyControlChars.Hyphen & strBoreSize

                'RMXXXXXXX 2009/09/11 Y.Miura �t���i�̐��ʂ��[���ɂȂ�s��C��
                decOpAmount(UBound(decOpAmount)) = intOpAmount
                'RM0912XXX 2009/12/09 Y.Miura�@�񎟓d�rC5���Z�s�v
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                'End If

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
