/**
 * The structure for making requests to the excelProcess handler
 * 
 * This structure represents the data that will be stringified into JSON and sent as a request to the server
 * 
 */
function transactionRequest(excelFilePath, action){
    this.excelFilePath = excelFilePath || null;
    this.action = action || "get";    //default action is get
    this.showPreview = true;           //true unless changed, since you will almost always want a preview
}