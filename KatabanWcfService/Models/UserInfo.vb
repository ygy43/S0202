Namespace Models
    ''' <summary>
    '''     ログインユーザー情報
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

        '<summary>ユーザーID</summary>
        <DataMember>
        Public Property user_id As String

        '<summary>拠点コード</summary>
        <DataMember>
        Public Property base_cd As String

        '<summary>国コード</summary>
        <DataMember>
        Public Property country_cd As String

        '<summary>営業所コード</summary>
        <DataMember>
        Public Property office_cd As String

        '<summary>担当者コード</summary>
        <DataMember>
        Public Property person_cd As String

        '<summary>メールアドレス</summary>
        <DataMember>
        Public Property mail_address As String

        '<summary>言語コード</summary>
        <DataMember>
        Public Property language_cd As String

        '<summary>通貨コード</summary>
        <DataMember>
        Public Property currency_cd As String

        '<summary>編集区分</summary>
        <DataMember>
        Public Property edit_div As String

        '<summary>ユーザー種別</summary>
        <DataMember>
        Public Property user_class As String

        '<summary>価格表示レベル</summary>
        <DataMember>
        Public Property price_disp_lvl As Integer

        '<summary>付加情報レベル</summary>
        <DataMember>
        Public Property add_information_lvl As Integer

        '<summary>利用機能レベル</summary>
        <DataMember>
        Public Property use_function_lvl As Integer

        '<summary>更新日</summary>
        <DataMember>
        Public Property current_datetime As String
    End Class
End Namespace