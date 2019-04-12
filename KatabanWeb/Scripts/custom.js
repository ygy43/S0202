
//クリックスタイルの設定
$("#currentOptions").on("click", ".clickable", function (event) {
    $(this).addClass("active").siblings().removeClass("active");
});
