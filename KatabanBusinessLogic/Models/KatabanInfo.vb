
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Managers
Imports KatabanCommon.Constants

Namespace Models
    ''' <summary>
    '''     形番情報
    ''' </summary>
    Public Class KatabanInfo
        Public Sub New(selectedData As SelectedInfo, userData As UserInfo)
            Me.SelectedData = selectedData

            'フル形番の作成
            Me.FullKataban = PriceManager.GetFullKataban(Me.SelectedData)

            '出荷場所の作成
            Me.ShipPlaces = New List(Of String) From {"JPN"}

            '価格情報の取得
            Me.Prices = PriceManager.GetPricesInfo(Me.FullKataban, Me.Currency, selectedData, userData)

            '現地定価の取得
            Me.LocalPrice = PriceManager.GetLocalPrice(Me.Currency,
                                                       userData.currency_cd,
                                                       userData.country_cd,
                                                       Me.FullKataban.Split(MyControlChars.Hyphen)(0),
                                                       Me.CheckDiv,
                                                       Prices.ls_price,
                                                       Prices.gs_price)

            '購入価格の取得
            Me.FobPrice = PriceManager.GetFobPrice(Me.DataType,
                                                   Me.Currency,
                                                   userData.country_cd,
                                                   Me.FullKataban.Split(MyControlChars.Hyphen)(0),
                                                   ShipPlaces.FirstOrDefault(),
                                                   Prices.gs_price)

            '販売数量単位情報の取得
            Me.SalesUnitInfo = PriceManager.GetQuantityUnitInfo(Me.FullKataban, "ja")

        End Sub

#Region "形番情報"

        ''' <summary>
        '''     データ種類
        ''' </summary>
        Public ReadOnly Property DataType As String
            Get
                Return Me.SelectedData.Series.division
            End Get
        End Property

        ''' <summary>
        '''     フル形番
        ''' </summary>
        Public ReadOnly Property FullKataban As String

        ''' <summary>
        '''     形番表示名称
        ''' </summary>
        Public ReadOnly Property DisplayName As String
            Get
                Return Me.SelectedData.Series.disp_name
            End Get
        End Property

        ''' <summary>
        '''     機種
        ''' </summary>
        Public ReadOnly Property Series As String
            Get

                Return Me.SelectedData.Series.series_kataban
            End Get
        End Property

        ''' <summary>
        '''     キー形番
        ''' </summary>
        Public ReadOnly Property KeyKataban As String
            Get
                Return Me.SelectedData.Series.key_kataban
            End Get
        End Property

        ''' <summary>
        '''     通貨
        ''' </summary>
        Public ReadOnly Property Currency As String
            Get
                Return Me.SelectedData.Series.currency_cd
            End Get
        End Property

        ''' <summary>
        '''     チェック区分
        ''' </summary>
        Public ReadOnly Property CheckDiv As String
            Get
                If CostCalculateNo = AccumulatePriceDiv.C5 Then
                    Return "Z" & Me.Prices.kataban_check_div & "(C5)"
                Else
                    Return "Z" & Me.Prices.kataban_check_div
                End If
            End Get
        End Property

        ''' <summary>
        '''     原価積算No.
        ''' </summary>
        Public ReadOnly Property CostCalculateNo As String
            Get
                Dim result As String = String.Empty

                If Left(Me.FullKataban, 3) = "JSG" Then
                    If Me.SelectedData.Symbols(6) = "T2YDU" Then
                        result = AccumulatePriceDiv.C5
                    End If
                End If
                try

                If Me.FullKataban.EndsWith("-FP1") Then
                    If Me.FullKataban.Contains("4GA") Or
                       Me.FullKataban.Contains("4GB") Or
                       Me.FullKataban.Contains("3GA") Or
                       Me.FullKataban.Contains("3GB") Then
                        Return result
                    End If
                End If

                Catch ex As Exception
                    dim test =ex.Message
                End Try
                'If KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc) = True Or
                '   Left(objKtbnStrc.strcSelection.strFullKataban, 10) = "CAC3-T2YDU" Then
                '    result = AccumulatePriceDiv.C5

                'End If

                'RM14070XX 2014/07/11 SWのC5対応
                If Me.FullKataban.StartsWith("SW-") And Me.Prices.kataban_check_div = "3" Then
                    If Me.FullKataban = "SW-T2YDU" Then
                    Else
                        result = AccumulatePriceDiv.C5
                    End If
                End If

                Return result
            End Get
        End Property

        ''' <summary>
        '''     評価タイプ
        ''' </summary>
        Public ReadOnly Property EvaluationType As String

        ''' <summary>
        '''     出荷場所
        ''' </summary>
        Public ReadOnly Property Plant As String

        ''' <summary>
        '''     保管場所
        ''' </summary>
        Public ReadOnly Property StorageLocation As String

        ''' <summary>
        '''     生産国
        ''' </summary>
        Public ReadOnly Property MadeCountry As String

        ''' <summary>
        '''     販売数量単位関連情報
        ''' </summary>
        Public ReadOnly Property SalesUnitInfo As QuantityUnitInfo

        ''' <summary>
        '''     販売単位
        ''' </summary>
        Public ReadOnly Property SalesUnit As String
            Get
                Return SalesUnitInfo.sales_unit
            End Get
        End Property

        ''' <summary>
        '''     販売単位名称
        ''' </summary>
        Public ReadOnly Property SalesUnitName As String
            Get
                Return _
                    IIf(SalesUnitInfo.qty_unit_nm = String.Empty, SalesUnitInfo.default_unit_nm,
                        SalesUnitInfo.qty_unit_nm)
            End Get
        End Property

        ''' <summary>
        '''     SAP基本単位
        ''' </summary>
        Public ReadOnly Property SapBaseUnit As String
            Get
                Return SalesUnitInfo.sap_base_unit
            End Get
        End Property

        ''' <summary>
        '''     販売数量
        ''' </summary>
        Public ReadOnly Property QuantityPerSalesUnit As String
            Get
                Return SalesUnitInfo.quantity_per_sales_unit
            End Get
        End Property

        ''' <summary>
        '''     ロット
        ''' </summary>
        Public ReadOnly Property OrderLot As String
            Get
                Return SalesUnitInfo.order_lot
            End Get
        End Property

        ''' <summary>
        '''     標準納期
        ''' </summary>
        Public ReadOnly Property StandardNouki As String

        ''' <summary>
        '''     適用個数
        ''' </summary>
        Public ReadOnly Property Kosuu As String

        ''' <summary>
        '''     EL区分
        ''' </summary>
        Public ReadOnly Property ElFlag As String
            Get
                Return PriceManager.GetElFlag(Me.FullKataban, ElDiv.IsEl)
            End Get
        End Property

        ''' <summary>
        '''     在庫情報
        ''' </summary>
        Public ReadOnly Property Stock As String
            Get
                Return PriceManager.GetStock(Me.FullKataban, "ja", Me.ShipPlaces.FirstOrDefault())
            End Get
        End Property

        ''' <summary>
        '''     出荷場所
        ''' </summary>
        Public ReadOnly Property ShipPlaces As List(Of String)

#End Region

#Region "価格情報"

        ''' <summary>
        '''     価格情報
        ''' </summary>
        Private Property Prices As PriceInfo

        ''' <summary>
        '''     定価
        ''' </summary>
        Public ReadOnly Property ListPrice As Decimal
            Get
                Return Prices.ls_price
            End Get
        End Property

        ''' <summary>
        '''     登録店価格
        ''' </summary>
        Public ReadOnly Property RegisterPrice As Decimal
            Get
                Return Prices.rg_price
            End Get
        End Property

        ''' <summary>
        '''     SS店価格
        ''' </summary>
        Public ReadOnly Property SsPrice As Decimal
            Get
                Return Prices.ss_price
            End Get
        End Property

        ''' <summary>
        '''     BS店価格
        ''' </summary>
        Public ReadOnly Property BsPrice As Decimal
            Get
                Return Prices.bs_price
            End Get
        End Property

        ''' <summary>
        '''     GS店価格
        ''' </summary>
        Public ReadOnly Property GsPrice As Decimal
            Get
                Return Prices.gs_price
            End Get
        End Property

        ''' <summary>
        '''     PS店価格
        ''' </summary>
        Public ReadOnly Property PsPrice As Decimal
            Get
                Return Prices.ps_price
            End Get
        End Property

        ''' <summary>
        '''     現地定価
        ''' </summary>
        Public ReadOnly Property LocalPrice As String

        ''' <summary>
        '''     購入価格
        ''' </summary>
        Public ReadOnly Property FobPrice As String

#End Region

#Region "選択情報"

        ''' <summary>
        '''     選択した情報
        ''' </summary>
        Public Property SelectedData As SelectedInfo

#End Region
    End Class
End Namespace