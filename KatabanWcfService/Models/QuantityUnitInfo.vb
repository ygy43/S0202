Namespace Models
    ''' <summary>
    '''     販売数量単位情報
    ''' </summary>
    Public Class QuantityUnitInfo
        Public Sub New()
            qty_unit_nm = String.Empty
            default_unit_nm = String.Empty
            sales_unit = String.Empty
            sap_base_unit = String.Empty
            quantity_per_sales_unit = String.Empty
            order_lot = String.Empty
        End Sub

        ''' <summary>
        '''     販売数量単位名称
        ''' </summary>
        Public Property qty_unit_nm As String

        ''' <summary>
        '''     販売数量単位名称（デフォルト）
        ''' </summary>
        Public Property default_unit_nm As String

        ''' <summary>
        '''     販売数量単位
        ''' </summary>
        Public Property sales_unit As String

        ''' <summary>
        '''     SAP単位
        ''' </summary>
        Public Property sap_base_unit As String

        ''' <summary>
        '''     単位数量
        ''' </summary>
        Public Property quantity_per_sales_unit As String

        ''' <summary>
        '''     ロット
        ''' </summary>
        Public Property order_lot As String
    End Class
End Namespace