'************************************************************************************
'*  ProgramID  �FKHPriceP7
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/06/10   �쐬�ҁFM.Kojima
'*
'*  �T�v       �F���J�j�J���p���[�V�����_�@MCP�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP7
    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strSuiryoku As String '����
        Dim strStroke As String '�X�g���[�N
        Dim strLead As String '���[�h������

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strSuiryoku = selectedData.Symbols(2).Trim
            strStroke = selectedData.Symbols(3).Trim
            strLead = selectedData.Symbols(5).Trim

            '��{���i�L�[�̐ݒ�
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            If (selectedData.Series.series_kataban.Trim = "MCP-W") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & _
                    MyControlChars.Hyphen & "00" & MyControlChars.Hyphen & _
                    strSuiryoku & MyControlChars.Hyphen & _
                    strStroke
            ElseIf (selectedData.Series.series_kataban.Trim = "MCP-S") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & _
                    MyControlChars.Hyphen & "00" & MyControlChars.Hyphen & _
                    strSuiryoku
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'FA���Z
            If (selectedData.Symbols(1).Trim = "FA") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & _
                    MyControlChars.Hyphen & "FA" & MyControlChars.Hyphen & _
                    strSuiryoku
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '���b�h��[���˂�(N)���Z
            If (selectedData.Symbols(7).Trim.Length <> 0) Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & _
                    MyControlChars.Hyphen & _
                    strSuiryoku & _
                    MyControlChars.Hyphen & _
                    selectedData.Symbols(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�X�C�b�`���Z���i�L�[
            If selectedData.Symbols(4).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                    strSuiryoku & MyControlChars.Hyphen & _
                    selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                    selectedData.Symbols(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '���[�h���������Z���i�L�[
                If selectedData.Symbols(5).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(selectedData.Series.series_kataban.Trim, 3) & _
                        MyControlChars.Hyphen & _
                        strLead
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
