<?php
require_once ROOT_PATH . '\inc\php\model\modelExcelTransaction.php';
require_once ROOT_PATH . '\inc\php\excelError.php';
require_once ROOT_PATH . '\inc\php\jsonKeys.php';

/**
 * Handles AJAX calls to process excel files
 * 
 * Create a new object and call handleRequest. Gets the request options for the Excel file and sets an object with those options. The object is passed to a loader where it is filled with data. The resulting data is returned as a JSON object. The JSON string is then returned.
 */
class controllerExcelTransaction {
    /**
     * Holds classes that implement the strategy for each type of request action
     * 
     * @var array _strategies Holds classes that implement the strategy for each type of request action
     */
    private $_strategies = array();
    
    /**
     * Set the strategies
     */
    public function __construct() {
        $this->_strategies["get"] = new handleGetExcelRequest();
        $this->_strategies["commit"] = new handleCommitExcelRequest();
        $this->_strategies["getAll"] = new handleGetAllExcelRequest();
    }
    
    /**
     * Process the GET request
     * 
     * Hands off the request data to the strategy associated with the request action
     * 
     * @param a JSON object that represents the data send with the original HTTP GET request
     * @return string JSON string representing the results of the requested action
     */
    public function handleRequest($requestData){
        if(array_key_exists($requestData->action, $this->_strategies)){
            return $this->_strategies[$requestData->action]->handleRequest($requestData);
        } else {
            return json_encode( excelError::createError("The requested action '" . $requestData->action . "' does not exist") );    //error because action does not exist
        }
    }
}

//------------------------------------------------------------------------------
//Strategy handlers for different request types

/**
 * Each strategy must receive and evaluate a request, then return a JSON string
 */
interface IHandleRequestStrategy{
    public function handleRequest($requestData);
}
/**
 * Common functions for request handlers
 */
abstract class handleRequestAbstract {
    /**
     * Prepares the worksheet array to be displayed as HTML
     * @param type $sheet
     */
    protected function makeWorksheetHTMLSafe($sheet){
        foreach($sheet as $rowkey => $row){
            foreach($row as $cellkey => $cell){
                    $sheet[$rowkey][$cellkey] = nl2br($cell);
            }
        }
        return $sheet;
    }
}
/**
 * Handles the request action 'get'
 * 
 * The get request action show a preview of information that will be commit to the database. This will most likely be the most common type of request
 * 
 */
class handleGetExcelRequest extends handleRequestAbstract implements IHandleRequestStrategy{
    /**
     * Call when the JSON object's property action is set to 'get'
     * @return string The string returned is a JSON object that represent the ExcelSheet object that has been loaded
     */
    public function handleRequest($requestData) {
        $loader = new modelExcelTransaction();
        
        if( isset($requestData->showPreview) ){                                 //if the preview option was set in the JSON object passed, limit the amount of returned rows
            if ($requestData->showPreview == FALSE){
                $loader->isPreview = FALSE;
            }
        }
        
        $loader->previewLength = 100;                                             //how many rows will be previewed. default to 10

        try{
            $workbook = $loader->load($requestData->excelFilePath);                 //load the object with data from the excel file
        } catch(Exception $e){
            return json_encode(  excelError::createError( $e->getMessage() )  );
        }
        
        $ryWorkbook = $workbook->toArray();
        
        //make the sheet data displayable for HTML
        foreach($ryWorkbook[jsonKeys::excelWorksheets] as $sheet){
            $htmlSafeSheetData = $this->makeWorksheetHTMLSafe( $sheet[jsonKeys::sheetData] );
            $sheet[jsonKeys::sheetData] = $htmlSafeSheetData;   
        }
        
        return json_encode( $ryWorkbook );                             //send back the resulting object as JSON
    }
}

/**
 * Enters excel data into a mySQL database
 */
class handleCommitExcelRequest extends handleRequestAbstract implements IHandleRequestStrategy{
    /**
     * Only reports error right now, not implemented
     * @param type $requestData
     * @return string
     */
    public function handleRequest($requestData) {
        $loader = new modelExcelTransaction();
        $workbook = $loader->load($requestData->excelFilePath);
        
        //make a new excelWorkbook
        //make a changes needed
        //commit the changes
        if($loader->commit($workbook)===0){                                     //just for the stub out, this will change
           return excelError::createError("Commit functionality not yet created"); 
        }
    }
}

/**
 * Returns all data from excel file (not a preview) to be displayed on the webpage. Not sure how often this will be used
 */
class handleGetAllExcelRequest extends handleRequestAbstract implements IHandleRequestStrategy{
    /**
     * Only reports error right now, not implemented
     * @param type $requestData
     * @return string
     */
    public function handleRequest($requestData) {
        return '{"responseStatus":"error"}';
    }
}

?>
