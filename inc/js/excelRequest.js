/**
 * The structure for making requests to the excelProcess handler
 */
function excelRequest(excelFilePath, action){
    this.excelFilePath = excelFilePath || null;
    this.action = action || "load";    //default action is null
    this.showPreview = true;           //true unless changed, since you will almost always want a preview
}