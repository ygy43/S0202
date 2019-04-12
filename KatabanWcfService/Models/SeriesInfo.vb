Namespace Models
    ''' <summary>
    '''     シリーズ形番検索結果
    ''' </summary>
    <DataContract>
    Public Class SeriesInfo
        Public Sub New()
            sort_key = String.Empty
            series_kataban = String.Empty
            key_kataban = String.Empty
            hyphen_div = String.Empty
            disp_kataban = String.Empty
            division = String.Empty
            disp_name = String.Empty
            price_no = String.Empty
            spec_no = String.Empty
            order_no = String.Empty
            currency_cd = String.Empty
            country_cd = String.Empty
        End Sub
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

        '<summary>表示名称</summary>
        <DataMember>
        Public Property disp_name As String

        '<summary>価格番号</summary>
        <DataMember>
        Public Property price_no As String

        '<summary>仕様種類番号</summary>
        <DataMember>
        Public Property spec_no As String

        '<summary>順番</summary>
        <DataMember>
        Public Property order_no As String

        '<summary>通貨コード</summary>
        <DataMember>
        Public Property currency_cd As String

        '<summary>販売国国コード（ログインユーザー国コード）</summary>
        <DataMember>
        Public Property country_cd As String
    End Class
End Namespace