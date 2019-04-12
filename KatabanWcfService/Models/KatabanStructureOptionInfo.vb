Namespace Models
    ''' <summary>
    '''     形番構成オプション情報
    ''' </summary>
    <DataContract>
    Public Class KatabanStructureOptionInfo

        '<summary>候補</summary>
        <DataMember>
        Public Property option_symbol As String

        '<summary>候補名称</summary>
        <DataMember>
        Public Property option_nm As String

        '<summary>デフォルト候補名称</summary>
        <DataMember>
        Public Property default_option_nm As String

        '<summary>表示順</summary>
        <DataMember>
        Public Property disp_seq_no As Integer

        '<summary>未使用</summary>
        <DataMember>
        Public Property price_acc_div As String

        '<summary>生産レベル</summary>
        <DataMember>
        Public Property place_lvl As String

    End Class
End Namespace