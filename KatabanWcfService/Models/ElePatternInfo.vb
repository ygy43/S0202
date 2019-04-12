Namespace Models
    ''' <summary>
    '''     構成オプション検証条件
    ''' </summary>
    <DataContract>
    Public Class ElePatternInfo
        Public Sub New()
            Me.search_seq_no = String.Empty
            Me.option_symbol = String.Empty
            Me.condition_cd = String.Empty
            Me.condition_seq_no = String.Empty
            Me.condition_seq_no_br = String.Empty
            Me.cond_option_symbol = String.Empty
        End Sub

        '<summary>データ区分</summary>
        <DataMember>
        Public Property search_seq_no As String

        '<summary>検証用構成値</summary>
        <DataMember>
        Public Property option_symbol As String

        '<summary>検証条件</summary>
        <DataMember>
        Public Property condition_cd As String

        '<summary>検証条件番号</summary>
        <DataMember>
        Public Property condition_seq_no As String

        '<summary>検証条件枝番</summary>
        <DataMember>
        Public Property condition_seq_no_br As String

        '<summary>表示構成値</summary>
        <DataMember>
        Public Property cond_option_symbol As String
    End Class
End Namespace