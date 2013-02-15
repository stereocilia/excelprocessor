<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Excel Process</title>
    </head>
    <body>
        <a href="javascript:handleLoadXmlClick();">Load XML</a>
        <div id="results" style="display: none;"></div>
    </body>
    <!-- include at bottom so page appears to load faster -->
        <?php 
            include('inc/php/jsInclude.php');                                   //includes all the javascript files needed
        ?>
        <script type="text/javascript">
            var excelFilePath = "userdata/Idaho_MarshBird_2008.xls";                         //for debug
            var loadingMessage = 'Please wait while your file is being loaded.<br><img src="img/loader32.gif"/>';
            function handleLoadXmlClick(){
                $('#results').html(loadingMessage);
                $('#results').show();   //show the loading icon
                processExcelFile(excelFilePath, "load");
                //excelFile should be singleton
                //excelFile.preview(excelFilePath);
                //excelFile.load(excelFilePath);
                //excelFile.commit(excelFilePath);
            }
            
            //displ ayData function MUST be implemented
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
                    $('#results').html(htmlOut);
                } else {
                    $('#results').html("Error processing your document");
                }
            }
            
        </script>
</html>
