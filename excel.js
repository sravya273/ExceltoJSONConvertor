
function ExportToTable() {  
   
   var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;  
   /*Checks whether the file is a valid excel file*/  
   if (regex.test($("#excelfile").val().toLowerCase())) {  
       var xlsxflag = false; /*Flag for checking whether excel is .xlsx format*/  
       var xlsflag = false; /*Flag for checking whether excel is .xls format*/
       if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0 ) {  
           xlsxflag = true;  
       } 
      else  if ($("#excelfile").val().toLowerCase().indexOf(".xls") > 0 ) {  
           xlsflag = true;  
       } 
       /*Checks whether the browser supports HTML5*/  
       if (typeof (FileReader) != "undefined") {  
           var reader = new FileReader();  
           reader.onload = function (e) { 
               console.log(reader); 
               var data = e.target.result; 

               /*Converts the excel data in to object*/  
               if (xlsxflag) {  
                   viewXLSX(reader,data);
                    
               }  
               else if (xlsflag) {  
                  viewXLS(reader,data);  
               }  
               
           }  
           if (xlsxflag) {/*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
               reader.readAsArrayBuffer($("#excelfile")[0].files[0]);  
           }  
           else if(xlsflag) {  
               reader.readAsBinaryString($("#excelfile")[0].files[0]);  
           }  
       }  
       else {  
           alert("Sorry! Your browser does not support HTML5!");  
       }  
   }  
   else {  
       alert("Please upload a valid Excel file!");  
   }  
}

    function viewXLSX(reader,data)
    {
        console.log("view");
        var workbook = XLSX.read(data, { type: 'binary' }); 
        var sheet_name_list = workbook.SheetNames;  
        var cnt = 0;
        sheet_name_list.forEach(function (y) 
                { 
                    /*Iterate through all sheets*/    
                    var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]); 
                    if (exceljson.length > 0 && cnt == 0) {  
                        BindTable(exceljson, '#exceltable');  
                        cnt++;  
                        document.getElementById("jsonformat").innerHTML = JSON.stringify(exceljson, undefined, 2);
                    }  
                });  
                $('#exceltable').show();  
                
    }
    function viewXLS(reader,data)
    {
        var workbook = XLS.read(data, { type: 'binary' });  
        var sheet_name_list = workbook.SheetNames;  
        var cnt = 0; 
        sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/    
                     var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);
                    //  console.log(exceljson);
                    // console.log("excel"+exceljson.length);
                      if(exceljson.length > 0 && cnt == 0) {  
                        // console.log("hi");
                        BindTable(exceljson, '#exceltable');  
                        document.getElementById("jsonformat").innerHTML = JSON.stringify(exceljson, undefined, 2);
                        cnt++;  
                    }  
                    //
                });  
               
                $('#exceltable').show();  
                
        
    }

function BindTable(jsondata, tableid) {/*Function used to convert the JSON array to Html Table*/  
    var columns = BindTableHeader(jsondata, tableid); /*Gets all the column headings of Excel*/  
    for (var i = 0; i < jsondata.length; i++) {  
        var row$ = $('<tr/>');  
        for (var colIndex = 0; colIndex < columns.length; colIndex++) {  
            var cellValue = jsondata[i][columns[colIndex]];  
            if (cellValue == null)  
                cellValue = "";  
            row$.append($('<td/>').html(cellValue));  
        }  
        $(tableid).append(row$);  
    }  
}  
function BindTableHeader(jsondata, tableid) {/*Function used to get all column names from JSON and bind the html table header*/  
    var columnSet = [];  
    var headerTr$ = $('<tr/>');  
    for (var i = 0; i < jsondata.length; i++) {  
        var rowHash = jsondata[i];  
        for (var key in rowHash) {  
            if (rowHash.hasOwnProperty(key)) {  
                if ($.inArray(key, columnSet) == -1) {/*Adding each unique column names to a variable array*/  
                    columnSet.push(key);  
                    headerTr$.append($('<th/>').html(key));  
                }  
            }  
        }  
    }  
    $(tableid).append(headerTr$);  
    return columnSet;  
} 
