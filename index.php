<!DOCTYPE html>
<html lang="en-US">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Excel to Database Loader</title>
    </head>
    
    <body>
        <a href="javascript:loadExcelClick();">Load Excel File</a>
        <div id="results" style="display: none;"></div>
        <!-- SCRIPTS -->
        <?php include('inc/php/jsInclude.php');                                     //includes all the javascript files needed ?>
        <script>
            var excelFilePath = "userdata/data_test.xlsx";                               //for debug only. AJAX file upload should assign this value otherwise.
            //The message displayed while Excel file is being
            var loadingMessage = 'Please wait while your file is being loaded.<br><img src="img/loader32.gif"/>';

            function loadExcelClick(){                                              //fires when link is clicked
                $('#results').html(loadingMessage).show();                           //show the loading message and animation
                excelFile.get(excelFilePath);                                       //gets the specified Excel file and returns a JSON string
            }

            //displayExcelData function MUST be implemented
            //do whatever you want with the returned JSON object
            function displayExcelData(data){
                if(data.responseStatus != "error"){                                 //make sure there are now errors
                    var htmlOut = "";
                    for(var i=0;i<data.excelWorksheets.length;i++){
                        var excelWorksheet = data.excelWorksheets[i];                                 //the returned Excel file as a JSON object
                        htmlOut += "<h3>";
                        htmlOut += excelWorksheet.title;
                        htmlOut += "</h3>";
                        
                        htmlOut += "<table>";
                        
                        htmlOut += "<tr>";
                        for (var columnType in excelWorksheet.columnTypes){
                            htmlOut += "<td>";
                            htmlOut += excelWorksheet.columnTypes[columnType];                        //show the cell contents
                            htmlOut += "</td>";
                        }
                        htmlOut += "</tr>";
                        for(var row in excelWorksheet.sheetData){                                      //for each row of data
                            htmlOut += "<tr>";
                            for(var cell in excelWorksheet.sheetData[row]){                            //for each cell in the row
                                htmlOut += "<td>";
                                htmlOut += excelWorksheet.sheetData[row][cell];                        //show the cell contents
                                htmlOut += "</td>"; 
                            }
                            htmlOut += "</tr>";
                        }
                        htmlOut += "</table>";
                    }
                    $('#results').html(htmlOut);
                } else {
                    $('#results').html(data.errorMessage);           //show an error message
                }
            }
        </script>
    </body>
</html>
