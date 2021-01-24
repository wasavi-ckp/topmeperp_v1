//讀取圖算資料
function getMapItem4Task() {
    $.ajax({
        url: "/ProjectPlan/getMapItem4Task",
        data: $("#formQueryForm").serialize(),
        method: "POST",
        dataType: "html",
        success: function (result) {
            $("#MapItem").html(result);
        }
    })
}

//更新TypeCodeL2
function getTypeCode2() {
    if ("" == $('#TypeCodeL1').val()) {
        $('#TypeCodeL2 option').remove();
        $('#TypeSub option').remove();
        return;
    }
    $.ajax({
        url: "/MaterialManage/getTypeCodeL2",
        data: { typecode1: $('#TypeCodeL1').val() },
        method: "POST",
        dataType: "JSON",
        success: function (result) {
            //更換typecode2 option
            var $el = $('#TypeCodeL2');
            $el.html(' ');
            $.each(result, function (key, value) {
                $el.append($("<option></option>").attr("value", key).text(value));
            });
        },
        error: function (xhr) {
            alert('Ajax request 發生錯誤');
        }
    })
}
//更新Sub Type
function getSubType() {
    $.ajax({
        url: "/MaterialManage/getSubType",
        data: { typecode: $('#TypeCodeL2').val() },
        method: "POST",
        dataType: "JSON",
        success: function (result) {
            //更換typecode2 option
            var $el = $('#TypeSub');
            $el.html(' ');
            $.each(result, function (key, value) {
                $el.append($("<option></option>").attr("value", key).text(value));
            });
        },
        error: function (xhr) {
            alert('Ajax request 發生錯誤');
        }
    })
}


//使用全選
function checkBoxStatus(that, chkMap) {
    var checked_status = $("#"+ that).prop('checked');
    //alert(checked_status);
    var checkBoxName = "input[name='" + chkMap + "']";
    $(checkBoxName).each(function () {
        this.checked = checked_status;
    });
}
//列印HTML
function printpage(printarea) {
    //var newstr = document.all.item(printarea).innerHTML;
    var newstr = document.getElementById(printarea).innerHTML;
    var oldstr = document.body.innerHTML;
    document.body.innerHTML = newstr;
    window.print();
    document.body.innerHTML = oldstr;
    return false;
}