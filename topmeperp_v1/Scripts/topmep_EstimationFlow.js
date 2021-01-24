//審核通過
$("#SendForm").click(function () {
    console.log("SendForm Cilck!!");
    $.ajax({
        url: '/Estimation/SendEstimationForm',
        data: $('#formESTMain').serialize(),
        type: "POST",
        dataType: 'text',
        success: function (msg) {
            alert(msg);
            window.location.reload();
        },
        error: function (xhr, ajaxOptions, thrownError) {
            //console.log(thrownError.)
            alert(xhr.responseText);
        }
    });
}
);
//退件
$("#RejectForm").click(function () {
    $.ajax({
        url: '/Estimation/RejectEstimationForm',
        data: $('#formESTMain').serialize(),
        type: "POST",
        dataType: 'text',
        success: function (msg) {
            alert(msg);
            close();
        },
        error: function (xhr, ajaxOptions, thrownError) {
            alert(thrownError);
        }
    });
});
//中止
$("#CancelForm").click(function () {
    $.ajax({
        url: '/Estimation/CancelEstimationForm',
        data: $('#formESTMain').serialize(),
        type: "POST",
        dataType: 'text',
        success: function (msg) {
            alert(msg);
            close();
        },
        error: function (xhr, ajaxOptions, thrownError) {
            alert(thrownError);
        }
    });
});

