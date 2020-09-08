$("#ctrl" ).change(function() {
  var x = document.getElementById("ctrl");
  var url = x.files[0].path;
  localStorage.setItem("templateFolder",JSON.stringify(url));
  if(url != ""){
    
    $("#template").empty();
    $('#template').append($('<option>', {value: null, text: "Template"}));
    
    changeTemplateOptions();
    if($("#orientation").val() != null && $("#paper_size").val() != null){
      $('#template').removeAttr('disabled');
    }

  var templateOptions = JSON.parse(localStorage.getItem("templateOptions"))
  templateOptions.forEach(element => {
    $('#template').append($('<option>', {
      value: element,
      text: element
    }));
  });
  


    $("#folderSelect" ).hide();
    $("#templateSelect" ).show();
    

  }
})


$("#generate_button").change(function(){
  if(($("#paper_size").val() == null) || ($("#orientation").val() == null) || ($("#template").val() == null) || ($("#template").val() == "Template")){
    alert("Please select all document settings before proceding.")
   
  }
  else{

    var p = document.getElementById("generate_button");
    var url = p.files[0].path;

    url = url.substr(0, url.lastIndexOf("/"));

    localStorage.setItem("outputFolder",JSON.stringify(url));

    generateWordDocument();
    
  }
})




$("#paper_size").change(function(){
  $("#template").empty();
  $('#template').append($('<option>', {value: null, text: "Template"}));
  
  changeTemplateOptions();
  if($("#orientation").val() != null){
    $('#template').removeAttr('disabled');
  }

  var templateOptions = JSON.parse(localStorage.getItem("templateOptions"))
  templateOptions.forEach(element => {
    $('#template').append($('<option>', {
      value: element,
      text: element
    }));
  });
  
  

})


$("#orientation").change(function(){
  $("#template").empty();
  $('#template').append($('<option>', {value: null, text: "Template"}));
  changeTemplateOptions();
  if($("#paper_size").val() != null){
    $('#template').removeAttr('disabled');
  }

  
  var templateOptions = JSON.parse(localStorage.getItem("templateOptions"))
  templateOptions.forEach(element => {
    $('#template').append($('<option>', {
      value: element,
      text: element
    }));
  });
  


})





$("#submitButton" ).click(function() {
  


if($("#csvFiles").val()!=""){
  $("#advance-search").show();
}

var myOpts1 = document.getElementById('searchfield-1').options;
console.log(myOpts1.length)
var myOpts2 = document.getElementById('searchfield-2').options;
var myOpts3 = document.getElementById('searchfield-3').options;
var myOpts4 = document.getElementById('searchfield-4').options;
var myOpts5 = document.getElementById('searchfield-5').options;
var myOpts6 = document.getElementById('searchfield-6').options;


if (!(myOpts1.length > 1 || myOpts2.length > 1 || myOpts3.length > 1 || myOpts4.length > 1 || myOpts5.length > 1 || myOpts6.length > 1)){

  //MODULE CODES 
var moduleCodes = JSON.parse(localStorage.getItem("moduleCodes"));
moduleCodes.forEach(element => {
  $('#searchfield-1').append($('<option>', {
    value: element,
    text: element
  }));
});


//SUPERVISORS
var supervisorList = JSON.parse(localStorage.getItem("supervisorList"));
supervisorList.forEach(element => {
  $('#searchfield-2').append($('<option>', {
    value: element,
    text: element
  }));
  
});

//PROJECT AUTHORS
var authorsList = JSON.parse(localStorage.getItem("authorsList"));
authorsList.forEach(element => {
  $('#searchfield-3').append($('<option>', {
    value: element,
    text: element
  }));
  
});

//CLIENT NAMES
var clientList = JSON.parse(localStorage.getItem("clientList"));
clientList.forEach(element => {
  $('#searchfield-4').append($('<option>', {
    value: element,
    text: element
  }));
  
});

//TECHNOLOGIES
var techList = JSON.parse(localStorage.getItem("techList"));

techList.forEach(element => {
  $('#searchfield-5').append($('<option>', {
    value: element,
    text: element
  }));
  
});

//FIELD AREAS
var fieldAreasList = JSON.parse(localStorage.getItem("fieldAreasList"));
fieldAreasList.forEach(element => {
  $('#searchfield-6').append($('<option>', {
    value: element,
    text: element
  }));
  
});

}


});

$("#search_button").click(function() {
  

    var brochureData = JSON.parse(localStorage.getItem("brochureData"))
    
    // EXTRACT VALUE FOR HTML HEADER. 
    // ('Book ID', 'Book Name', 'Category' and 'Price')
    var col = [];
    for (var i = 0; i < brochureData.length; i++) {
        for (var key in brochureData[i]) {
            if (col.indexOf(key) === -1) {
                col.push(key);
            }
        }
    }

    // CREATE DYNAMIC TABLE.
    var table = document.createElement("table");

    // CREATE HTML TABLE HEADER ROW USING THE EXTRACTED HEADERS ABOVE.

    var tr = table.insertRow(-1);                   // TABLE ROW.

    for (var i = 0; i < col.length - 6; i++) {
        var th = document.createElement("th");      // TABLE HEADER.
        th.innerHTML = col[i];
        tr.appendChild(th);
    }

    // ADD JSON DATA TO THE TABLE AS ROWS.
    for (var i = 0; i < brochureData.length; i++) {

        tr = table.insertRow(-1);

        for (var j = 0; j < col.length -6; j++) {
            var tabCell = tr.insertCell(-1);
            tabCell.innerHTML = brochureData[i][col[j]];
        }
    }

    // FINALLY ADD THE NEWLY CREATED TABLE WITH JSON DATA TO A CONTAINER.
    var divContainer = document.getElementById("showData");
    divContainer.innerHTML = "";
    divContainer.appendChild(table);

  $("#docreate").show();
  $("#showData").show()
  
});


$(document).ready(function() { 
  $('input[type="file"]').change(function() { 
    var x = document.getElementById("csvFiles");
    // console.log("File: " + x.files[0].path);
    
    if (x.value != "") {
      x.disabled = true;
    }
    
    
    var XLSX = require("xlsx");
    
    var url = x.files[0].path;
    var oReq = new XMLHttpRequest();
    oReq.open("GET", url, true);
    oReq.responseType = "arraybuffer";
    
    oReq.onload = function(e) {
      var arraybuffer = oReq.response;
      
      /* convert data to binary string */
      var data = new Uint8Array(arraybuffer);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
      var bstr = arr.join("");
      
      /* Call XLSX */
      var workbook = XLSX.read(bstr, {
        type: "binary"
      });
      
      /* DO SOMETHING WITH workbook HERE */
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      var output = XLSX.utils.sheet_to_json(worksheet, {
        raw: true
      })
      // jsonAllData = output;
      // console.log("data: " + output);
      localStorage.setItem('jsonAllData', JSON.stringify(output));
      
      // return jsonAllData;
      
    }
    oReq.send();  }); 
}); 


