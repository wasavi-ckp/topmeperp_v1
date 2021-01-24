
//由ID 取得資料填入表單
function getPaymentTerm(contractid) {
    //alert(contractid);
    $.ajax({
        url: "/PurchaseForm/getPaymentTerms",
        type: "GET",
        data: { contractid: contractid },
        dataType: "JSON",
        success: function (data) {
            console.log(data);
            $('#project_id').val(data.PROJECT_ID);
            $('#contract_id').val(data.CONTRACT_ID);
            $('#date1').val(data.DATE_1);
            $('#date2').val(data.DATE_2);
            $('#date3').val(data.DATE_3);
            $('#paymenttype').val(data.PAYMENT_TYPE);
            $('#paymentcash').val(data.PAYMENT_CASH);
            $('#payment_date1').val(data.PAYMENT_UP_TO_U_DATE1);
            $('#payment_date2').val(data.PAYMENT_UP_TO_U_DATE2);
            $('#payment_1').val(data.PAYMENT_UP_TO_U_1);
            $('#payment_2').val(data.PAYMENT_UP_TO_U_2);
            $('#paymentadvance').val(data.PAYMENT_ADVANCE_RATIO);
            $('#paymentadvance_cash').val(data.PAYMENT_ADVANCE_CASH_RATIO);
            $('#paymentadvance_1').val(data.PAYMENT_ADVANCE_1_RATIO);
            $('#paymentadvance_2').val(data.PAYMENT_ADVANCE_2_RATIO);
            $('#paymentestimated').val(data.PAYMENT_ESTIMATED_RATIO);
            $('#paymentestimated_cash').val(data.PAYMENT_ESTIMATED_CASH_RATIO);
            $('#paymentestimated_1').val(data.PAYMENT_ESTIMATED_1_RATIO);
            $('#paymentestimated_2').val(data.PAYMENT_ESTIMATED_2_RATIO);
            $('#paymentretention').val(data.PAYMENT_RETENTION_RATIO);
            $('#paymentretention_cash').val(data.PAYMENT_RETENTION_CASH_RATIO);
            $('#paymentretention_1').val(data.PAYMENT_RETENTION_1_RATIO);
            $('#paymentretention_2').val(data.PAYMENT_RETENTION_2_RATIO);
            $('#usancecash').val(data.USANCE_CASH);
            $('#usance_date1').val(data.USANCE_UP_TO_U_DATE1);
            $('#usance_date2').val(data.USANCE_UP_TO_U_DATE2);
            $('#usance_1').val(data.USANCE_UP_TO_U_1);
            $('#usance_2').val(data.USANCE_UP_TO_U_2);
            $('#usanceadvance').val(data.USANCE_ADVANCE_RATIO);
            $('#usanceadvance_cash').val(data.USANCE_ADVANCE_CASH_RATIO);
            $('#usanceadvance_1').val(data.USANCE_ADVANCE_1_RATIO);
            $('#usanceadvance_2').val(data.USANCE_ADVANCE_2_RATIO);
            $('#usancegoods').val(data.USANCE_GOODS_RATIO);
            $('#usancegoods_cash').val(data.USANCE_GOODS_CASH_RATIO);
            $('#usancegoods_1').val(data.USANCE_GOODS_1_RATIO);
            $('#usancegoods_2').val(data.USANCE_GOODS_2_RATIO);
            $('#usancefinished').val(data.USANCE_FINISHED_RATIO);
            $('#usancefinished_cash').val(data.USANCE_FINISHED_CASH_RATIO);
            $('#usancefinished_1').val(data.USANCE_FINISHED_1_RATIO);
            $('#usancefinished_2').val(data.USANCE_FINISHED_2_RATIO);
            $('#usanceretention').val(data.USANCE_RETENTION_RATIO);
            $('#usanceretention_cash').val(data.USANCE_RETENTION_CASH_RATIO);
            $('#usanceretention_1').val(data.USANCE_RETENTION_1_RATIO);
            $('#usanceretention_2').val(data.USANCE_RETENTION_2_RATIO);
            if ((data.PAYMENT_FREQUENCY) == 'O') {
                $("input[name=payfrequency][value='O']").attr('checked', true);
            } else {
                $("input[name=payfrequency][value='T']").attr('checked', true);
            }

            if (data.PAYMENT_TYPE == 'S') {
                $("input[name=payterms][value='S']").attr('checked', true);
            } else {
                $("input[name=payterms][value='P']").attr('checked', true);
            }

            $('#paymentInfo').modal('show'); // show bootstrap modal when complete loaded

        },
        error: function (jqXHR, textStatus, errorThrown) {
            alert('Error get data from ajax');
        }
    });

}

$('#savePayment').click(function () {
    var method1 = $('input[name=payfrequency]:radio:checked').val();
    if (typeof (method1) == "undefined") { // 注意檢查完全沒有選取的寫法，這行是精華
        alert("請選取付款方式，並檢查輸入之日期是否正確！");
        return false;
    }
    if (method1 == "O") {
        alert("您選擇的付款方式為每月付款2次");
    }
    if (method1 == "T") {
        alert("您選擇的付款方式為每月1次");
    }
    var method2 = $('input[name=payterms]:radio:checked').val();
    if (typeof (method2) == "undefined") { // 注意檢查完全沒有選取的寫法，這行是精華
        alert("請選取付款比例與票期，並檢查輸入之資料是否正確！");
        return false;
    }
    if (method2 == "P") {
        if ($("#paymenttype").val() == "") {
            alert("請選取按期估驗的合約類型，例如:連工帶料");
            return false;
        }
        alert("您選擇的付款比例與票期為按期估驗");
    }
    if (method2 == "S") {
        alert("您選擇的付款比例與票期為階段付款");
    }
    var s = $('#formPayment').serialize();
    var URLs = "/PurchaseForm/addPaymentTerms";
    $.ajax({
        url: URLs,
        data: $('#formPayment').serialize(),
        type: "POST",
        dataType: 'text',
        success: function (msg) {
            alert(msg);
            window.location.reload()
        },
        error: function (xhr, ajaxOptions, thrownError) {
            alert(thrownError);
        }
    });
});

