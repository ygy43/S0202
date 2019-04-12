Namespace Models

    Public Class SeriesInfoFullKataban
        '<summary>表示順</summary>
        <DataMember>
        Public Property sort_key As String

        '<summary>シリーズ形番</summary>
        <DataMember>
        Public Property series_kataban As String

        '<summary>キー形番</summary>
        <DataMember>
        Public Property key_kataban As String

        '<summary>ハイフン</summary>
        <DataMember>
        Public Property hyphen_div As String

        '<summary>表示形番</summary>
        <DataMember>
        Public Property disp_kataban As String

        '<summary>検索区分</summary>
        <DataMember>
        Public Property division As String
        
        '<summary>検索区分</summary>
        <DataMember>
        Public Property kataban_check_div As String

        '<summary>部品名称</summary>
        <DataMember>
        Public Property parts_nm As String
        
        '<summary>モデル名称</summary>
        <DataMember>
        Public Property model_nm As String

        '<summary>表示名称</summary>
        <DataMember>
        Public Property disp_name As String

        '<summary>価格番号</summary>
        <DataMember>
        Public Property price_no As String

        '<summary>仕様種類番号</summary>
        <DataMember>
        Public Property spec_no As String

        '<summary>通貨コード</summary>
        <DataMember>
        Public Property currency_cd As String

        ''<summary>販売国国コード（ログインユーザー国コード）</summary>
        '<DataMember>
        'Public Property country_cd As String
    End Class
End NameSpace