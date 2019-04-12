'************************************************************************************
'*  ProgramID  �FKHPriceP4
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/12/18   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F���W���[���N�[�����g�o���u   �f�b�u�d�Q�E�f�b�u�r�d�Q�V���[�Y
'*�@�X�V����@�@�F
'*�@�@�@�@�@�@�@�@�I�v�V����B�i��t�j�̒ǉ�      RM0912039 2009/12/17 Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP4

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intStation As Integer
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '�A��(���W���[�����Z)���i�L�[
            Select Case selectedData.Symbols(2).Trim
                Case "A", "B"
                    intStation = 1
                Case Else
                    intStation = CInt(selectedData.Symbols(2).Trim)

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(5).Trim
            decOpAmount(UBound(decOpAmount)) = intStation

            '�R�C���I�v�V�������Z���i�L�[
            If Len(selectedData.Symbols(6).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(6).Trim
                decOpAmount(UBound(decOpAmount)) = intStation
            End If

            '���̑��I�v�V�������Z���i�L�[
            'RM0912039 2009/12/17 Y.Miura �I�v�V����B�i��t�j�ǉ�
            'If Len(selectedData.Symbols(7).Trim) <> 0 Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
            '                                               selectedData.Symbols(7).Trim
            '    decOpAmount(UBound(decOpAmount)) = intStation
            'End If
            strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim

                        'If intStation >= 3 And strOpArray(intLoopCnt).Trim.Equals("B") Then
                        '    decOpAmount(UBound(decOpAmount)) = 2
                        'Else
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        'End If

                        '��RM1303003 2013/03/04 Y.Tachi
                        decOpAmount(UBound(decOpAmount)) = 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "B"
                                If intStation >= 3 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "S"
                                If selectedData.Symbols(2).Trim = "A" Or _
                                   selectedData.Symbols(2).Trim = "B" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = selectedData.Symbols(2).Trim
                                End If
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
