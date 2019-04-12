Namespace Models
    ''' <summary>
    '''     形番構成情報
    ''' </summary>
    <DataContract>
    Public Class KatabanStructureInfo
        Public Sub New()
            'Me.SelectedValue = String.Empty
            Me.ktbn_strc_seq_no = String.Empty
            Me.element_div = String.Empty
            Me.structure_div = String.Empty
            Me.addition_div = String.Empty
            Me.hyphen_div = String.Empty
            Me.default_nm = String.Empty
            Me.ktbn_strc_nm = String.Empty

            Me.PluralGroupData = String.Empty
        End Sub

#Region "DB項目"

        '<summary>オプション番号</summary>
        <DataMember>
        Public Property ktbn_strc_seq_no As String

        '<summary>オプション区分</summary>
        <DataMember>
        Public Property element_div As String

        '<summary>オプション入力区分</summary>
        <DataMember>
        Public Property structure_div As String

        '<summary>付加情報区分</summary>
        <DataMember>
        Public Property addition_div As String

        '<summary>継続ハイフンフラグ</summary>
        <DataMember>
        Public Property hyphen_div As String

        '<summary>構成デフォルト名称</summary>
        <DataMember>
        Public Property default_nm As String

        '<summary>構成名称</summary>
        <DataMember>
        Public Property ktbn_strc_nm As String

#End Region

#Region "追加した項目"

        '<summary>複数選択可能オプションのグループ情報</summary>
        <DataMember>
        Public Property PluralGroupData As String

        '<summary>テキストボックスの幅</summary>
        <DataMember>
        Public Property Width As Integer

#End Region
    End Class
End Namespace