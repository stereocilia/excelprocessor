<!DOCTYPE html>
<html lang="en-US">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Excel to Database Loader</title>
    </head>
    
    <body>
        <a id="loadLink" href="javascript:loadExcelClick();">Load Spreadsheet File</a>
        <a id="updateLink" style="display:none;" href="javascript:updateExcelClick();">Update with my changes</a>
        <div id="transactionMessage" style="display: none;"></div>
        <div id="results" style="display: none;"></div>
        <!-- SCRIPTS -->
        <?php include('inc/php/jsInclude.php');                                     //includes all the javascript files needed ?>
        <script type="text/javascript">
            var excelWorkbook = null;                                           //stores the loaded workbook
            var excelFilePath = "userdata/data.csv";                               //for debug only. AJAX file upload should assign this value otherwise.
            //The message displayed while Excel file is being
            var loadingMessage = 'Please wait while your file is being loaded.<br><img src="img/loader32.gif"/>';
            var updatingMessage = 'Please wait while your file is being updated.<br><img src="img/loader32.gif"/>';
            
            function loadExcelClick(){                                              //fires when link is clicked
                $('#loadLink').fadeOut(100, function(){
                   $('#transactionMessage').html(loadingMessage).fadeIn();                           //show the loading message and animation 
                });
                excelFile.get(excelFilePath);                                       //gets the specified Excel file and returns a JSON string
            }
            
            function updateExcelClick(){
                $('#updateLink').fadeOut(100);
                $('#results').fadeOut(100, function(){
                    $('#transactionMessage').html(updatingMessage).fadeIn();                           //show the loading message and animation          
                });
                excelFile.update(excelFilePath);                                       //gets the specified Excel file and returns a JSON string 
            }

            //displayExcelData function MUST be implemented
            //do whatever you want with the returned JSON object
            function displayExcelData(data){
                if(data.responseStatus !== "error"){                                 //make sure there are now errors
                    excelWorkbook = data;
                    $('#transactionMessage').fadeOut(500, function(){
                      $('#updateLink').fadeIn();//maybe don't show this until a change has been made
                      $('#results').fadeIn();
                    });
                    
                    var htmlOut = "";
                    for(var i=0;i<excelWorkbook.excelWorksheets.length;i++){
                        var excelWorksheet = excelWorkbook.excelWorksheets[i];                                 //the returned Excel file as a JSON object
                        htmlOut += "<h3 title='Sheet Name'>";
                        htmlOut += excelWorksheet.title;
                        htmlOut += "</h3>";
                        
                        htmlOut += "<table>";
                        htmlOut += "<tr>";
                        for (var columnType in excelWorksheet.columnTypes){
                            htmlOut += "<td title='Data Type'>";
                            htmlOut += excelWorksheet.columnTypes[columnType];                        //show the cell contents
                            htmlOut += "</td>";
                        }
                        htmlOut += "</tr>";
                        for(var row in excelWorksheet.sheetData){                                      //for each row of data
                            htmlOut += "<tr>";
                            for(var cell in excelWorksheet.sheetData[row]){                            //for each cell in the row
                                htmlOut += "<td title='Cell Data'>";
                                htmlOut += excelWorksheet.sheetData[row][cell];                        //show the cell contents
                                htmlOut += "</td>"; 
                            }
                            htmlOut += "</tr>";
                        }
                        htmlOut += "</table>";
                    }
                    $('#results').html(htmlOut);
                } else {
                    var errorMessage = excelError.errorType[excelWorkbook.errorType].message;
                    errorMessage += "<br>" + excelWorkbook.errorMessage;
                    $('#results').html(errorMessage);           //show an error message
                }
            }
        </script>
        
        <!--//TODO: PRBO - BOOKMARK: The last thing you were doing was create a ui that allows the user to deselect items they do not want in their worksheet. SEE note below.-->
        <!--//NOTE: Worksheets will have check boxes by them that are already selected. 
        If the user does not want a particular worksheet it will be deselected. 
        This will also be the same process for removing columns that are unwanted.
        Also, the user should be allowed to select a different row that is the heading row.
        I want to use jQuery tabs, one per sheet. http://jqueryui.com/tabs/#vertical
        The vertical tabs seems like the would work well. I want each tab to have a check mark by it to include that worksheet
        -->
        <!--//TODO: PRBO - If merged cells detected in DATASET during commit, through error -->
        <!--//TODO: PRBO - Things to think about the future: Mapping from the excel spreadsheet into the database -->
    </body>
</html>
