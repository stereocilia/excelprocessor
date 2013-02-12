<?php
require_once ROOT_PATH . '\inc\php\model\modelProcessExcel.php';

/**
 * Handles AJAX calls to process excel files
 * 
 * Create a new object and call handleRequest. Gets the request options for the Excel file and sets an object with those options. The object is passed to a loader where it is filled with data. The resulting data is returned as a JSON object. The JSON string is then returned.
 */
class controllerProcessExcel {
    
    private $requestData = NULL;
    
    /**
     * Process the GET request
     * @return string Return value depends on the action requested
     */
    public function handleRequest(){
        //TODO: This should be handled by the "landing page"
        $this->requestData = json_decode(  stripcslashes( $_GET['data'] )  );

        //TODO: there is a better way to handle the decision to choose between actions
        //TODO: these decision MAY be able to be called by the landing page, depending on how the problem above is solved
        switch ($this->requestData->action){
            case "preview" :                                                    //see a preview with no processing
                break;
            case "load" :
                return $this->handleLoad();
                break;
            case "commit" :
                break;
            default: break;
        }
    }
    
    /**
     * Call when the JSON object's property action is set to 'load'
     * @return string The string returned is a JSON object that represent the ExcelSheet object that has been loaded
     */
    private function handleLoad(){
        $loader = new modelProcessExcel();
        //if the preview option was set in the JSON object passed, limit the amount of
        //returned rows
        if( isset($this->requestData->showPreview) ){
            if ($this->requestData->showPreview == FALSE){
                $loader->isPreview = FALSE;
            }
        }
        //load the object with data from the excel file
        $objExcelProcessor = $loader->load($this->requestData->excelFilePath);
        
        if($objExcelProcessor) return $objExcelProcessor->toJSON();             //send back the resulting object as JSON
        else return NULL;
    }
}
?>
