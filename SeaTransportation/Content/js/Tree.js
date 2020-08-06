function Tree(A, B, C) {
    var z = 0;
    var x = 0;
    $(A).append("<ul><li class='jstree-open'>所有数据<ul></ul></li></ul>")
    $.ajax({
        type: "post",
        url: B,
        dataType: "json",
        success: function (Tree1) {
            $.each(Tree1, function (i) {
                if (C == "" || C == undefined || C == null) {
                    $("" + A + " ul").eq(0).children().children().append("<li  onclick='Tree1(" + Tree1[i].id + ")'>" + Tree1[i].name + "<ul class='Treex" + Tree1[i].id + "'></ul></li>")
                    z++;
                } else {
                    $("" + A + " ul").eq(0).children().children().append("<li>" + Tree1[i].name + "<ul class='Treex" + Tree1[i].id + "'></ul></li>")
                    z++;
                }
                 if ((C == "" || C == undefined || C == null) && z == Tree1.length) {
                     $("" + A + " ").jstree();
                     setTimeout(function () {
                         $("#j1_1_anchor").attr("onclick", "Tree1(0)")
                     }, 100)
                } else {
                     $.ajax({
                         type: "post",
                         url: C + Tree1[i].id,
                         dataType: "json",
                         success: function (Tree2) {
                             $.each(Tree2, function (m) {
                                 $("" + A + " .Treex" + Tree1[i].id + "").append("<li  onclick='Tree2(" + Tree2[m].id + ")'>" + Tree2[m].name + "</li>")
                             });
                             x++
                             if (x == Tree1.length) {
                                 $("" + A + " ").jstree();
                                 setTimeout(function () {
                                     $("#j1_1_anchor").attr("onclick", "Tree1(0)")
                                 }, 100)
                             }
                         }
                    });
                }
            });
        }
    });
}
