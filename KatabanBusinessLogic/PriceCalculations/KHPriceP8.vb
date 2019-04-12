'************************************************************************************
'*  ProgramID  �FKHPriceP8
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/06/10   �쐬�ҁFM.Kojima
'*
'*  �T�v       �F�y�ʃN�����v�V�����_�@CAC�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP8
    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '��{���i�L�[�̐ݒ�
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                selectedData.Symbols(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '�t���i���Z
            If (selectedData.Symbols(8).Trim <> "") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & _
                    MyControlChars.Hyphen & _
                    selectedData.Symbols(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�X�C�b�`���Z���i�L�[
            If selectedData.Symbols(4).Trim.Length <> 0 Then

                'RM1801025_�I�v�V�����ǉ��Ή�
                '�^�C���b�h����
                If selectedData.Symbols(7).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        "TIEROD" & MyControlChars.Hyphen & _
                        selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                        selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '��t���i���Z
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                        selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                        "TIEROD"
                    decOpAmount(UBound(decOpAmount)) = 1

                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                        selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                        selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

                '���[�h���������Z���i�L�[
                If selectedData.Symbols(5).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If selectedData.Symbols(4).Trim = "T2YD" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        selectedData.Series.series_kataban.Trim & _
                        MyControlChars.Hyphen & "T2YD" & MyControlChars.Hyphen & _
                        selectedData.Symbols(5).Trim

                    ElseIf selectedData.Symbols(4).Trim = "T2YDT" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         selectedData.Series.series_kataban.Trim & _
                         MyControlChars.Hyphen & "T2YDT" & MyControlChars.Hyphen & _
                         selectedData.Symbols(5).Trim

                    ElseIf selectedData.Symbols(4).Trim = "T2JH" Or _
                    selectedData.Symbols(4).Trim = "T2JV" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         selectedData.Series.series_kataban.Trim & _
                         MyControlChars.Hyphen & "T2J" & MyControlChars.Hyphen & _
                         selectedData.Symbols(5).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         selectedData.Series.series_kataban.Trim & _
                         MyControlChars.Hyphen & "T" & MyControlChars.Hyphen & _
                         selectedData.Symbols(5).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
