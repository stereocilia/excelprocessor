<?php
//require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/PHPExcel/Classes/PHPExcel.php';  //PHPExcel Libraryrequire_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/excelProcessor.php';
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/model/modelProcessExcel.php';

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
        $this->requestData = json_decode($_GET['data']);

        //TODO: maybe have a raw previw, where it doesn't try to figure anything out... just returns 50 rows or so and lets the user pick which one will be the row that contains the column header
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
        //send back the resulting object as JSON
        return $objExcelProcessor->toJSON();
    }
}
?>
