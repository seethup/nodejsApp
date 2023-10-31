$(document).ready(function () {
    $('#tableSelector').change(function () {
        var selectedTable = $(this).val();
        console.log(selectedTable)
        //  $('#tableContainer').empty(); // Clear existing tables
        $('#' + selectedTable).show();



        
        //document.getElementById('successMsg').innerHTML='';
        //document.getElementById('succMsg').innerHTML='';
        var count = $('#tableSelector').children('option').length;
        var length = selectedTable.length;
        console.log("length ====" + length)
        console.log("count ===" + count);
        var subs = selectedTable.substr(5, length)
        console.log("substr ==== " + selectedTable.substr(5, length))
        for (var x = 0; x < count; x++) {
            if (x != subs) {
                $('#table' + x).hide();
            }

        }

        var elementExists = document.getElementById("emailMsg");
        var elementExists1 = document.getElementById("successMsg");
        var elementExists2 = document.getElementById("succMsg");

        if(elementExists){
            document.getElementById('emailMsg').innerHTML='';
        }
        if(elementExists1){
            document.getElementById('successMsg').innerHTML='';
        }
        if(elementExists2){
            document.getElementById('succMsg').innerHTML='';
        }        

    });
    $('#monthSelector').change(function () {
        var month = $(this).val();

        document.getElementById('month').value= month;

        console.log(document.getElementById('month').value)
    });

});

function calculatePrice(ex) {

    console.log(ex.id);
    var idNum = ex.id.substr(2, ex.id.length)
    console.log(idNum)
    var price = $('#price' + idNum).val();
    console.log(price)
    var isGst = $('#gst' + idNum).is(":checked");
    console.log(isGst);
    var finalPrice = "";
    if (isGst) {
        finalPrice = price;
    }
    else {
        finalPrice = (price * 1.12).toFixed(2);
    }

    console.log("Final PRice === " + finalPrice)

    ex.value = finalPrice;
}

const sendMail =() =>{

   // fetch('http://localhost:3000/sendMail').then(function(response) { console.log(response.status())});



    $.ajax({
        url: 'sendMail',
        type: 'GET',
        success: function(response){
            console.log("call success")
            //$('#emailMsg').val("Email Sent Successfully");
            document.getElementById('emailMsg').innerHTML='Email Sent Successfully';
        },
        error: function(error){
            console.log(error);
        }
    });
}