/**
 * Contains the functions needed to create requests that will act on the given Excel file and return them.
 * 
 * Creates a transactionRequest object and calls the function that encodes and sends it
 * 
 */
var excelFile = new Object();

/**
 * Send request to transaction page
 * 
 * The data of the request is created from the transactionRequest object that is passed to the function. It is then sent to the excelTransaction page via AJAX get. The information is returned to a function called displayExcelData that receives a signle variable called data.
 * 
 */
excelFile.sendRequest = function(objTransactionRequest){
    var strRequestData = "data=" + JSON.stringify(objTransactionRequest);
    $.ajax({
        url:'excelTransaction.php', 
        data:strRequestData,
        success: displayExcelData,
        dataType: "json"
    });
}


/**
 * Gets the file located at the given path, creates a small sample of the data, and returns it to the displayData function
 */
excelFile.get = function(excelFilePath){
    var tr = new transactionRequest(excelFilePath, "get");
    tr.showPreview = true;  //always returns a sample of the data in the excel file
    this.sendRequest(tr);
}

/**
 * Commits the given file to a mySQL database
 * 
 * NOT YET FUNCTIONAL
 * 
 */
excelFile.commit = function(excelFilePath){
    var er = new transactionRequest(excelFilePath, "commit");
    er.showPreview = true;  //always returns a sample of the data in the excel file
    this.sendRequest(tr);
}


/**
 * Gets the file located at the given path and returns it to the displayData function
 * 
 * NOT YET FUNCTIONAL
 * 
 */
excelFile.getAll = function(excelFilePath){
    var er = new transactionRequest(excelFilePath, "getAll");
    er.showPreview = false;  //always returns a sample of the data in the excel file
    this.sendRequest(tr);
}
