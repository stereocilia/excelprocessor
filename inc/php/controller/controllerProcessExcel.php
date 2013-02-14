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
     * @param a JSON object that represents the data send with the original HTTP GET request
     * @return string Return value depends on the action requested
     */
    public function handleRequest($requestData){

        $this->requestData = $requestData;

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
        $loader->previewLength = 2;                                             //how many rows will be previewed. default to 10
        
        //this function must be responsible for finding the column index
        //why?
        //because data may have to be taken from the excel file twice.
        //the excel file object cannot call itself to be loaded, that doesn't make sense.
        //this is the next best place I can think of since it asks for object to be loaded
        $excelWorksheet = $loader->load($this->requestData->excelFilePath);  //load the object with data from the excel file
        //call the method for finding the column heading, get and index back
        $columnIndex = $excelWorksheet->findColumnIndex();
        if($columnIndex != 0){
            //now load the file again with the new index if its not 0
            $loader->columnIndex = $columnIndex;
            $excelWorksheet = $loader->load($this->requestData->excelFilePath); 
        }
        
        return $excelWorksheet->toJSON();             //send back the resulting object as JSON
    }
}
?>
