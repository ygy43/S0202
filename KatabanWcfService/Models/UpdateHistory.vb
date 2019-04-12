Namespace Models
    ''' <summary>
    '''     メニュー画面の更新履歴
    ''' </summary>
    <DataContract>
    Public Class UpdateHistory
        Public Sub New()
            language_cd =string.Empty
            message = String.Empty
            seq_no = string.Empty
        End Sub

        '<summary>言語コード</summary>
        <DataMember>
        Public Property language_cd As String

        '<summary>更新メッセージ</summary>
        <DataMember>
        Public Property message As String

        '<summary>メッセージ番号</summary>
        <DataMember>
        Public Property seq_no As String

    End Class
End NameSpace