@ModelType S0202.ViewModels.Prices.PricesIndexViewModel
@Code
    ViewData("Title") = "Index"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<h3>@Model.DisplayName</h3>
<h3>@Model.FullKataban</h3>

<hr />

<div class="row">
    <div class="col-md-2">
    </div>
    <div class="col-md-2">
        @Html.Label("形番チェック")
    </div>
    <div class="col-md-2">
        @Html.Label("プラント")
    </div>
    <div class="col-md-2">
        @Html.Label("標準納期")
    </div>
    <div class="col-md-2">
        @Html.Label("適用個数")
    </div>
    <div class="col-md-2">
        @Html.Label("E/L該当品区分")
    </div>
</div>
<div class="row">
    <div class="col-md-2">
    </div>
    <div class="col-md-2">
        @Html.DisplayFor(Function(model) model.CheckDiv)
    </div>
    <div class="col-md-2">
        @Html.DropDownListFor(Function(model) model.SelectedShipPlace, New SelectList(Model.ShipPlaces), New With {.class = "form-control"})
    </div>
    <div class="col-md-2">
        @Html.DisplayFor(Function(model) model.StandardNouki)
    </div>
    <div class="col-md-2">
        @Html.DisplayFor(Function(model) model.Kosuu)
    </div>
    <div class="col-md-2">
        @Html.DisplayFor(Function(model) model.ElDiv)
    </div>
</div>

<hr />

<div class="row">
    <div class="col-md-5">
        <table id="currentOptions" class="table table-bordered table-autowidth table-hover">
            @*タイトル*@
            <thead>
                <tr class="info">
                    <th>区分</th>
                    <th>単価</th>
                    <th>通貨</th>
                </tr>
            </thead>

            @*価格情報*@
            <tbody>
                <tr class="clickable" onclick="clickPrice('@Model.ListPrice');">
                    <td>定価</td>
                    <td>@Model.ListPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr class="clickable" onclick="clickPrice('@Model.RegisterPrice');">
                    <td>登録店</td>
                    <td>@Model.RegisterPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr class="clickable" onclick="clickPrice('@Model.SsPrice');">
                    <td>SS店</td>
                    <td>@Model.SsPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr class="clickable" onclick="clickPrice('@Model.BsPrice');">
                    <td>BS店</td>
                    <td>@Model.BsPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr class="clickable" onclick="clickPrice('@Model.GsPrice');">
                    <td>GS店</td>
                    <td>@Model.GsPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr class="clickable" onclick="clickPrice('@Model.PsPrice');">
                    <td>PS店</td>
                    <td>@Model.PsPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr>
                    <td>現地定価</td>
                    <td>@Model.LocalPrice</td>
                    <td>@Model.Currency</td>
                </tr>
                <tr>
                    <td>購入価格</td>
                    <td>@Model.FobPrice</td>
                    <td>@Model.Currency</td>
                </tr>
            </tbody>
        </table>
    </div>
    <div class="col-md-7">
        <table class="table table-bordered table-autowidth table-hover">
            @*価格入力情報*@
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">掛率</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.Rate, New With {.class = "form-control", .onchange = "calculatePrice();"})
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">金額</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.TotalWithoutTax, New With {.class = "form-control"})
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">単価</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.Price, New With {.class = "form-control"})
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">消費税</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.Tax, New With {.class = "form-control"})
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">数量</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.Amount, New With {.class = "form-control", .onchange = "calculatePrice();"})
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <div class="input-group-addon">合計</div>
                            @Html.TextBoxFor(Function(model) model.SelectedPrice.TotalWithTax, New With {.class = "form-control"})
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
</div>

<div class="row">
    <button type="submit" class="btn btn-primary">仕様出力</button>
    <button type="submit" class="btn btn-primary">価格詳細</button>
    <button type="submit" class="btn btn-primary">3D CAD</button>
</div>

@Scripts.Render("~/bundles/jquery")

<script>

    //クリックスタイルの設定
    function clickPrice(price) {
        //スタイルの設定
        $(this).addClass("active").siblings().removeClass("active");

        //価格計算
        var floatPrice = parseFloat(price);

        $("#SelectedPrice_Rate").val("1.0000");
        $("#SelectedPrice_Price").val(floatPrice);

        calculatePrice();
    };

    /**価格計算 */
    function calculatePrice() {
        var floatPrice = parseFloat($("#SelectedPrice_Price").val());
        var floatRate = parseFloat($("#SelectedPrice_Rate").val());
        var intAmount = parseInt($("#SelectedPrice_Amount").val());

        //金額
        $("#SelectedPrice_TotalWithoutTax").val(floatPrice * intAmount * floatRate);
        //消費税
        $("#SelectedPrice_Tax").val(floatPrice * intAmount * 0.08 * floatRate);
        //合計
        $("#SelectedPrice_TotalWithTax").val(floatPrice * intAmount * 1.08 * floatRate);
    };

</script>