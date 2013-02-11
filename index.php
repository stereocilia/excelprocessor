<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Excel Process</title>
        <?php //includes all the javascript files needed ?>
        <?php include('inc/php/jsInclude.php'); ?>
        <script>
            var excelFilePath = "userdata/data.xlsx";
            var ajaxResults = "";
            function handleLoadXmlClick(){
                processExcelFile(excelFilePath, "load");
            }
            
            //displayData function MUST be implemented
            //do whatever you want with the returned JSON object
            function displayData(data){
                //data = JSON.parse(data);
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
