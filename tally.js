$("#form").submit(function (event) {

        //disable the default form submission
		proccess();
        event.preventDefault();
		var excelData = [];
		showDataExcel(event);
	function showDataExcel(event) { 
    var file = event.target[0].files[0];
  var reader = new FileReader();
 reader.onload = function (event) {
 var data = event.target.result;
 var workbook = XLSX.read(data, {
 type: 'binary'
 });

 workbook.SheetNames.forEach(function (sheetName) {
     excelData = [];
   var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
     for (var i = 0; i < XL_row_object.length; i++)
     { excelData.push(XL_row_object[i]);
     };
             
      var json_object = JSON.stringify(XL_row_object);
	    $.ajax({
            type: "POST",
            url: "http://localhost:5800/upload",
			ContentType: "application/json",
			 data: [json_object],
		//cache: false,
			
		//contentType: 'application/vnd.ms-excel',
		enctype: 'multipart/form-data',
		processData: false,
           
            success: function(data) {
			console.log(data)
			
			$("#form1").submit(function clickdownload(){
	
	var pom = document.createElement('a');

	var filename = "file.xml";
	var pom = document.createElement('a');
	var bb = new Blob([data], {type: 'text/plain'});

	pom.setAttribute('href', window.URL.createObjectURL(bb));
	pom.setAttribute('download', filename);

	pom.dataset.downloadurl = ['text/plain', pom.download, pom.href].join(':');
	pom.draggable = true; 
	pom.classList.add('dragout');
	pom.click();
});

$("#form2").submit(function tallytrsn(){
	
    var xhr = new XMLHttpRequest();
	xhr.open("POST", "http://localhost:9000", true);//localhost:9001--is tally server
	//xhr.setRequestHeader('Access-Control-Allow-Origin',' *');
    console.log('ok');
	xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	xhr.withCredentials = true;
	xhr.onload = function ()
	{
	//console.log(xhr.responseText);
	};
	xhr.send(data);     
	alert("Tranfser To Tally Succesfully")
});
	
            },
            error: function(jqXHR, textStatus, err) {
                alert('text status '+textStatus+', err '+err)
            }
        });
	  
     })
};

reader.onerror = function (ex) {
console.log(ex);
};

reader.readAsBinaryString(file);
};
		

});