Namespace Models
    ''' <summary>
    '''     ���O�C�����[�U�[���
    ''' </summary>
    <DataContract>
    Public Class UserInfo
        Public Sub New()
            user_id = String.Empty
            base_cd = String.Empty
            country_cd = String.Empty
            office_cd = String.Empty
            person_cd = String.Empty
            mail_address = String.Empty
            language_cd = String.Empty
            currency_cd = String.Empty
            edit_div = String.Empty
            user_class = String.Empty
            price_disp_lvl = 0
            add_information_lvl = 0
            use_function_lvl = 0
            current_datetime = String.Empty
        End Sub

        '<summary>���[�U�[ID</summary>
        <DataMember>
        Public Property user_id As String

        '<summary>���_�R�[�h</summary>
        <DataMember>
        Public Property base_cd As String

        '<summary>���R�[�h</summary>
        <DataMember>
        Public Property country_cd As String

        '<summary>�c�Ə��R�[�h</summary>
        <DataMember>
        Public Property office_cd As String

        '<summary>�S���҃R�[�h</summary>
        <DataMember>
        Public Property person_cd As String

        '<summary>���[���A�h���X</summary>
        <DataMember>
        Public Property mail_address As String

        '<summary>����R�[�h</summary>
        <DataMember>
        Public Property language_cd As String

        '<summary>�ʉ݃R�[�h</summary>
        <DataMember>
        Public Property currency_cd As String

        '<summary>�ҏW�敪</summary>
        <DataMember>
        Public Property edit_div As String

        '<summary>���[�U�[���</summary>
        <DataMember>
        Public Property user_class As String

        '<summary>���i�\�����x��</summary>
        <DataMember>
        Public Property price_disp_lvl As Integer

        '<summary>�t����񃌃x��</summary>
        <DataMember>
        Public Property add_information_lvl As Integer

        '<summary>���p�@�\���x��</summary>
        <DataMember>
        Public Property use_function_lvl As Integer

        '<summary>�X�V��</summary>
        <DataMember>
        Public Property current_datetime As String
    End Class
End Namespace