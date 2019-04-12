Namespace Models
    ''' <summary>
    '''     形番構成オプション情報(全SeqNo)
    ''' </summary>
    <DataContract>
    Public Class KatabanStructureOptionInfoAllSeqNo
        '<summary>候補</summary>
        <DataMember>
        Public Property option_symbol As String

        '<summary>SeqNo</summary>
        <DataMember>
        Public Property ktbn_strc_seq_no As String

        '<summary>生産レベル</summary>
        <DataMember>
        Public Property place_lvl As String
    End Class
End Namespace