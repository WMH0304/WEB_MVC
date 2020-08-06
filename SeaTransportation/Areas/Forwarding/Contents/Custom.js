//关闭模态框
function Close(A) {
    if (A==undefined||A==null) {
        $(".modal input").val("")
        $(".modal input[type='checkbox']").prop('checked', false)
        $(" input[type='checkbox']").val(true)
        $(".modal input[type='radio']").prop('checked', false)
        $(".modal select").val(0)
        $(".modal textarea").val("")
        //$("#StaffImage").attr("src", "");
    } else {
        $("#"+A+" input").val("")
        $("#" + A + " input[type='checkbox']").prop('checked', false)
        $("#" + A + " input[type='checkbox']").val(true)
        $("#" + A + " input[type='radio']").prop('checked', false)
        $("#" + A + " select").val(0)
        $("#" + A + " textarea").val("")
        //$("#StaffImage").attr("src", "");
    }
}
//切换记录
function Records(m) {
    $(".record").fadeOut(500)
    setTimeout(function () {
        $(".record").eq(m).fadeIn(500)
    }, 500);
}
//全屏
function requestFullScreen(element) {
    var requestMethod = element.requestFullScreen || //W3C
    element.webkitRequestFullScreen ||    //Chrome等
    element.mozRequestFullScreen || //FireFox
    element.msRequestFullScreen; //IE11
    requestMethod.call(element);
}
//退出全屏
function exitFull()
{
    $("#image").attr("style", "height:500px;display:none;cursor:")
    $(".FullScreen div").attr("class", "")
    var exitMethod = document.exitFullscreen || //W3C
    document.mozCancelFullScreen ||    //Chrome等
    document.webkitExitFullscreen || //FireFox
    document.webkitExitFullscreen; //IE11
    exitMethod.call(document);
    T = 0;
    L = 0;
    R = 0;
    B = 0;
    height = 0;
    trolley = false;
}

//联系电话
function Phone(Event) {
    var phone = /^1[34578]\d{9}$/;
    var Telephone = /^0\d{2,3}-?\d{7,8}$/;
    if (Event.value == "" || Event.value == undefined || Event.value == null) {
        return null;
    } else if (!phone.test(Event.value) && !Telephone.test(Event.value)) {
        Event.value = "";
        layer.msg('您输入的号码不正确，请重新输入', { icon: 0 });
    } 
}

//浮点数
function Float(Event, Begin, Finish) {
    var float = /^(-?\d+)(\.\d+)?$/;
    if (Event.value == "" || Event.value == undefined || Event.value == null) {
        return null;
    } else if (!float.test(Event.value)) {
        Event.value = "";
        layer.msg('您输入的数据不正确，请重新输入', { icon: 0 });
    }
    if (Begin > 0 || Finish > 0) {
        if (Begin > Event.value) {
            Event.value = "";
            layer.msg('您输入的数据超出范围，请重新输入', { icon: 0 });
        } else
        if (Finish < Event.value) {
            Event.value = "";
            layer.msg('您输入的数据超出范围，请重新输入', { icon: 0 });
        }
    }
}

//只能输入数字
function Number(Event) {
    var number = /(^\d+$)/;
    if (Event.value == "" || Event.value == undefined || Event.value ==null) {
        return null;
    }  else if (!number.test(Event.value)) {
        Event.value = "";
        layer.msg('您输入的格式不正确，请重新输入', { icon: 0 });
    }  
}
//只能输入英文（大写）
function English(Event) {
    var English = /^[A-Z]+$/
    if (!English.test(Event.value)) {
        layer.msg('您输入的格式不正确，请重新输入', { icon: 0 });
        Event.value = ""
    }
}
//只能输入英文（大小写）
function Englishning(Event) {
    var English = /^[a-zA-Z]+$/
    if (!English.test(Event.value)) {
        layer.msg('您输入的格式不正确，请重新输入', { icon: 0 });
        Event.value = ""
    }
}
//姓名
function Name(Event) {
    var name = /^[\u4E00-\u9FA5\uf900-\ufa2d·s]{2,20}$/
    if (!name.test(Event.value)) {
        layer.msg('您输入的格式不正确，请重新输入', { icon: 0 });
        Event.value=""
    }
}

//身份证地址
function Address (Event) {
    var name = /^[\u4E00-\u9FA5\uf900-\ufa2d·s]{2,20}$/
    if (!name.test(Event.value)) {
        layer.msg('您输入的身份证地址不正确，请重新输入', { icon: 0 });
        Event.value=""
    }
}

//身份证
var vcity = {
    11: "北京", 12: "天津", 13: "河北", 14: "山西", 15: "内蒙古",
    21: "辽宁", 22: "吉林", 23: "黑龙江", 31: "上海", 32: "江苏",
    33: "浙江", 34: "安徽", 35: "福建", 36: "江西", 37: "山东", 41: "河南",
    42: "湖北", 43: "湖南", 44: "广东", 45: "广西", 46: "海南", 50: "重庆",
    51: "四川", 52: "贵州", 53: "云南", 54: "西藏", 61: "陕西", 62: "甘肃",
    63: "青海", 64: "宁夏", 65: "新疆", 71: "台湾", 81: "香港", 82: "澳门", 91: "国外"
};

//身份证验证
function checkCard(Event) {
    var card = Event.value;
    //是否为空
    if (card == '' || card == undefined || card == null) {
        $("#BasicsInformation input[name='Nativeplace']").val("")
        $("#BasicsInformation input[name='Nuber_Birth']").val("")
        $("#BasicsInformation input[name='Nuber_Age']").val("")
        return null;
    }else
    //校验长度，类型
    if (isCardNo(card) == false) {
        layer.msg('您输入的身份证号码不正确，请重新输入', { icon: 0 });
        Event.value = "";
        $("#BasicsInformation input[name='Nativeplace']").val("")
        $("#BasicsInformation input[name='Nuber_Birth']").val("")
        $("#BasicsInformation input[name='Nuber_Age']").val("")
    }else
    //检查省份
    if (checkProvince(card) == false) {
        layer.msg('您输入的身份证号码不正确,请重新输入', { icon: 0 });
        Event.value = "";
        $("#BasicsInformation input[name='Nativeplace']").val("")
        $("#BasicsInformation input[name='Nuber_Birth']").val("")
        $("#BasicsInformation input[name='Nuber_Age']").val("")
    }else
    //校验生日
    if (checkBirthday(card) == false) {
        layer.msg('您输入的身份证号码生日不正确,请重新输入', { icon: 0 });
        Event.value = "";
        $("#BasicsInformation input[name='Nativeplace']").val("")
        $("#BasicsInformation input[name='Nuber_Birth']").val("")
        $("#BasicsInformation input[name='Nuber_Age']").val("")
    }else
    //检验位的检测
    if (checkParity(card) == false) {
        layer.msg('您的身份证校验位不正确,请重新输入', { icon: 0 });
        Event.value = "";
        $("#BasicsInformation input[name='Nativeplace']").val("")
        $("#BasicsInformation input[name='Nuber_Birth']").val("")
        $("#BasicsInformation input[name='Nuber_Age']").val("")
    } else {
        $("#BasicsInformation input[name='Nativeplace']").val(vcity[card.substr(0, 2)])
        if (card.length == '15') {
            var arr_data = card.match( /^(\d{6})(\d{2})(\d{2})(\d{2})(\d{3})$/);
            var birthday = '19' + arr_data[2] + '-' + arr_data[3] + '-' + arr_data[4];
            var time = parseInt(new Date().getFullYear() - (parseInt(19+arr_data[2]) + parseInt(arr_data[3]) / 12));
            $("#BasicsInformation input[name='Nuber_Birth']").val(birthday)
            if (time < 18 || time > 55) {
                layer.msg('抱歉，年龄录用超出范围', { icon: 0 });
                $("#BasicsInformation input[name='Nativeplace']").val("")
                $("#BasicsInformation input[name='Nuber_Birth']").val("")
                $("#BasicsInformation input[name='Nuber_Age']").val("")
            } else {
                $("#BasicsInformation input[name='Nuber_Age']").val(time)
            }
        }else
        if (card.length == '18') {
            var arr_data = card.match(/^(\d{6})(\d{4})(\d{2})(\d{2})(\d{3})([0-9]|X)$/);
            var birthday = arr_data[2] + '-' + arr_data[3] + '-' + arr_data[4]
            var time = parseInt(new Date().getFullYear() - (parseInt(arr_data[2]) + parseInt(arr_data[3]) / 12));
            $("#BasicsInformation input[name='Nuber_Birth']").val(birthday)
            if (time < 18 || time > 55) {
                layer.msg('抱歉，年龄超出范围', { icon: 0 });
                $("#BasicsInformation input[name='Nativeplace']").val("")
                $("#BasicsInformation input[name='Nuber_Birth']").val("")
                $("#BasicsInformation input[name='Nuber_Age']").val("")
            } else {
                $("#BasicsInformation input[name='Nuber_Age']").val(time)
            }
        }
    }
};

//检查号码是否符合规范，包括长度，类型
function isCardNo(card) {
    //身份证号码为15位或者18位，15位时全为数字，18位前17位为数字，最后一位是校验位，可能为数字或字符X
    var reg = /(^\d{15}$)|(^\d{17}(\d|X)$)/;
    if (reg.test(card) == false) {
        return false;
    }
    return true;
};

//取身份证前两位,校验省份
function checkProvince(card) {
    var province = card.substr(0, 2);
    if (vcity[province] == undefined) {
        return false;
    }
    return true;
};

//检查生日是否正确
function checkBirthday(card) {
    var len = card.length;
    //身份证15位时，次序为省（3位）市（3位）年（2位）月（2位）日（2位）校验位（3位），皆为数字
    if (len == '15') {
        var re_fifteen = /^(\d{6})(\d{2})(\d{2})(\d{2})(\d{3})$/;
        var arr_data = card.match(re_fifteen);
        var year = arr_data[2];
        var month = arr_data[3];
        var day = arr_data[4];
        var birthday = new Date('19' + year + '/' + month + '/' + day);
        return verifyBirthday('19' + year, month, day, birthday);
    }
    //身份证18位时，次序为省（3位）市（3位）年（4位）月（2位）日（2位）校验位（4位），校验位末尾可能为X
    if (len == '18') {
        var re_eighteen = /^(\d{6})(\d{4})(\d{2})(\d{2})(\d{3})([0-9]|X)$/;
        var arr_data = card.match(re_eighteen);
        var year = arr_data[2];
        var month = arr_data[3];
        var day = arr_data[4];
        var birthday = new Date(year + '/' + month + '/' + day);
        return verifyBirthday(year, month, day, birthday);
    }
    return false;
};

//校验日期
function verifyBirthday(year, month, day, birthday) {
    var now = new Date();
    var now_year = now.getFullYear();
    //年月日是否合理
    if (birthday.getFullYear() == year && (birthday.getMonth() + 1) == month && birthday.getDate() == day) {
        //判断年份的范围（3岁到100岁之间)
        var time = now_year - year;
        if (time >= 3 && time <= 100) {
            return true;
        } 
        return false;
    }
    return false;
};

//校验位的检测
function checkParity(card) {
    //15位转18位
    card = changeFivteenToEighteen(card);
    var len = card.length;
    if (len == '18') {
        var arrInt = new Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2);
        var arrCh = new Array('1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2');
        var cardTemp = 0, i, valnum;
        for (i = 0; i < 17; i++) {
            cardTemp += card.substr(i, 1) * arrInt[i];
        }
        valnum = arrCh[cardTemp % 11];
        if (valnum == card.substr(17, 1)) {
            return true;
        }
        return false;
    }
    return false;
};

//15位转18位身份证号
function changeFivteenToEighteen(card) {
    if (card.length == '15') {
        var arrInt = new Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2);
        var arrCh = new Array('1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2');
        var cardTemp = 0, i;
        card = card.substr(0, 6) + '19' + card.substr(6, card.length - 6);
        for (i = 0; i < 17; i++) {
            cardTemp += card.substr(i, 1) * arrInt[i];
        }
        card += arrCh[cardTemp % 11];
        return card;
    }
    return card;
};
