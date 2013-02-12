<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Excel Process</title>
        <?php 
            include('inc/php/jsInclude.php');                                   //includes all the javascript files needed
        ?>
        <script>
            var excelFilePath = "userdata/data.xlsx";                         //for debug
            var ajaxResults = "";
            function handleLoadXmlClick(){
                processExcelFile(excelFilePath, "load");
            }
            
            //displayData function MUST be implemented
            //do whatever you want with the returned JSON object
            function displayData(data){
                if(data.responseStatus != "error"){
                    var excelData = data.excelData;
                    var htmlOut = "";
                    htmlOut += "<table>";
                    for(var row in excelData){
                        htmlOut += "<tr>";
                        for(var cell in excelData[row]){
                            htmlOut += "<td>";
                            htmlOut += excelData[row][cell];    
                            htmlOut += "</td>"; 
                        }
                        htmlOut += "</tr>";
                    }
                    htmlOut += "</table>";
                    htmlOut += "<div>";
                    htmlOut += data.dataTypes;
                    htmlOut += "</div>";
                    $('.results').html(htmlOut);
                } else {
                    $('.results').html("Error processing your document");
                }
            }
            
        </script>
    </head>
    <body>
        <?php
            echo "hell worlkd<br>";
        ?>
        <a style="cursor: pointer; text-decoration: underline;" onclick="handleLoadXmlClick();">Load XML</a>
        <div class="results"></div>
    </body>
</html>
