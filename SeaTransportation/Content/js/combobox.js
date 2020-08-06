//点击table的行选择复选框
function Check(selector, index) {
    if ($(selector).eq(index).find("input[type='checkbox']").attr("checked")) {
        $(selector).eq(index).find("input[type='checkbox']").get(0).removeAttribute("checked");
    } else {
        $(selector).eq(index).find("input[type='checkbox']").attr("checked", "checked");
    }
}

//绑定下拉框
function appendOptions(id, url, name) {
    $.getJSON(url, function (data) {
        $("#" + id).empty();
        $("#" + id).append("<option value=" + 0 + ">" + name + "</option>");
        $.each(data, function (i) {
            $("#" + id).append("<option value=" + data[i].id + ">" + data[i].name + "</option>");
        });
    });
}

function appendOption(id, url) {
    $.getJSON(url, function (data) {
        $("#" + id).empty();
        $("#" + id).append("<option value=" + 0 + ">" + "----请选择----" + "</option>");
        $.each(data, function (i) {
            $("#" + id).append("<option value=" + data[i].id + ">" + data[i].name + "</option>");
        });
    });
}


