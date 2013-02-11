var columnHeadingIndex = null;

/**
 * Sends AJAX request to retrieve information about the given Excel file
 * 
 * @param string excelFilePath The path to the file that will be processed
 * @param string action The action to perform on the file. Right now its just "load"
 */
function processExcelFile(excelFilePath, action){
    var objRequestData = new excelRequest(excelFilePath, action);
    //objRequestData.action = action;
    //objRequestData.excelFilePath = excelFilePath;
    //objRequestData.showPreview = true;
    //if(objRequestData.excelFilePath == null){
        //throw error
    //}
    if(columnHeadingIndex){ 
        objRequestData.columnHeadingIndex = columnHeadingIndex;
    }
    var strRequestData = "data=" + JSON.stringify(objRequestData);
    //$('.results').html();
    $.ajax({
        url:'processExcel.php', 
        data:strRequestData,
        success: function(data){
           displayData(data); 
        },
        dataType: "json"
    }); 
}

function setColumnHeadingIndex(index){
    columnHeadingIndex = index;
}