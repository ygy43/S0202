Namespace Models
    ''' <summary>
    '''     価格情報
    ''' </summary>
    <DataContract>
    Public Class PriceInfo
        Public Sub New()

            kataban_check_div = String.Empty
            place_cd = String.Empty
            currency_cd = String.Empty
            country_group_cd = string.Empty
            country_cd = String.Empty
            ls_price = 0
            rg_price = 0
            ss_price = 0
            bs_price = 0
            gs_price = 0
            ps_price = 0
        End Sub

        '<summary>形番</summary>
        <DataMember>
        Public Property kataban As String

        '<summary>チェック区分</summary>
        <DataMember>
        Public Property kataban_check_div As String

        '<summary>プラント</summary>
        <DataMember>
        Public Property place_cd As String

        '<summary>通貨</summary>
        <DataMember>
        Public Property currency_cd As String

#Region "フル形番情報のみ使用"

        '<summary>国グループ</summary>
        <DataMember>
        Public Property country_group_cd As String

        '<summary>国コード</summary>
        <DataMember>
        Public Property country_cd As String

#End Region

        '<summary>定価</summary>
        <DataMember>
        Public Property ls_price As Decimal

        '<summary>登録店価格</summary>
        <DataMember>
        Public Property rg_price As Decimal

        '<summary>SS店価格</summary>
        <DataMember>
        Public Property ss_price As Decimal

        '<summary>BS店価格</summary>
        <DataMember>
        Public Property bs_price As Decimal

        '<summary>GS店価格</summary>
        <DataMember>
        Public Property gs_price As Decimal

        '<summary>PS店価格</summary>
        <DataMember>
        Public Property ps_price As Decimal
    End Class
End Namespace