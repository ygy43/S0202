'************************************************************************************
'*  ProgramID  �FKHPriceQ5
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/02/02   �쐬�ҁFT.Yagyu
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FLAD�V���[�Y
'*
'* �ύX
'*              �񎟓d�r�Ή�             RM1004012 2010/04/23 Y.Miura 
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceQ5

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOptionKataban As String = ""
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            If Len(selectedData.Symbols(2).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '2010/08/23 MOD RM1008009(9��VerUP) START--->
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                'selectedData.Symbols(2).Trim
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                selectedData.Symbols(3).Trim
                '2010/08/23 MOD RM1008009(9��VerUP) <--- END
                decOpAmount(UBound(decOpAmount)) = 1
            End If
            '�I�v�V�������Z���i�L�[
            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        Select Case selectedData.Symbols(2).Trim
                            Case "10A", "15A"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                                "S-B"
                            Case "20A", "25A"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                                "L-B"
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                        selectedData.Symbols(2).Trim & "-1"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '�񎟓d�r���Z
            'RM1004012 2010/04/23 Y.Miura
            If UBound(selectedData.Symbols.ToArray()) >= 5 Then
                If selectedData.Symbols(5) <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-OP-" & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If


        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

