<?php
//require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/PHPExcel/Classes/PHPExcel.php';  //PHPExcel Libraryrequire_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/common.php';
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/model/modelProcessExcel.php';

/**
 * Handles AJAX calls to process excel files
 * 
 * Create a new object and call handleRequest. It will return a JSON string that has data and other properties of the excel file.
 */
class controllerProcessExcel {
    
    private $requestData = NULL;
    private $excelFileToRead = "";
    private $loader = NULL;
    private $excelFile = NULL;
    private $excelSheet = NULL;
    private $ryExcelSheet = array();
    private $action = "";
    private $primitiveTypes = NULL;
    
    /**
     * 
     * @return string Return value depends on the action requested
     */
    public function handleRequest(){
        $this->requestData = json_decode($_GET['data']);
        $this->action = $this->requestData->action;
        
        $this->excelFileToRead = $this->requestData->excelFilePath;//"";  //path to the file that will be read

        $this->loader = new modelProcessExcel();
        
        //TODO: maybe have a raw previw, where it doesn't try to figure anything out... just returns 50 rows or so and lets the user pick which one will be the row that contains the column header
        switch ($this->action){
            case "load" :
                return $this->handleLoad();
                break;
            case "commit" :
                break;
            default: break;
        }
    }
    
    /**
     * 
     * @return string The string returned is a JSON object that represent the ExcelSheet object that has been loaded
     */
    private function handleLoad(){
        //if the preview option was set in the JSON object passed, limit the amount of
        //returned rows
        if( isset($this->requestData->showPreview) ){
            if ($this->requestData->showPreview == TRUE){
                $this->loader->isPreview = TRUE;
            }
        }
        
        /* future options
        if( isset($this->requestData->columnHeadingIndex) ){
            //use the integer value to be the row that give the value names for the
            //table
        }
         */
        $this->excelFile = $this->loader->load($this->excelFileToRead);
        $this->excelSheet = $this->excelFile->getSheet(0);
        
        $this->ryExcelSheet = $this->excelSheet->toArray(null,true);
        
        $ryReturn = array();
        $ryReturn["dataTypes"] = $this->getColumnDataTypes();
        $ryReturn["excelData"] = $this->ryExcelSheet;
        return json_encode($ryReturn);
    }
    
    /**
     * Gets  data types of all columns for the current excelSheet of $this object
     */
    private function getColumnDataTypes(){
        $this->primitiveTypes = $this->getColumnPrimitiveDataTypes();           //get what the PHPExcel library says the data types are
        $cellTypes = array();
        $i = 0;
        $xlSheet = $this->ryExcelSheet[1];
        foreach($xlSheet as $cell){
            //get the type for each cell
            switch($this->primitiveTypes[$i]){
                case "s" :
                    $cellTypes[] = "string";
                    break;
                case "b" :
                    $cellTypes[] = "boolean";
                    break;
                case "n" :
                    $cellTypes[] = $this->handleNumericType($cell);
                    break;
                default :
                    $cellTypes[]= "";
                    break;
            }
            $i++;
        }
        return $cellTypes;
    }
    private function handleNumericType($number){
        if( is_time($number) )
            return "time";
        elseif( is_date($number) )
            return "date";
        else
            return "number";
    }
    /**
     * Gets primative data types of all columns for the current excelSheet of $this object
     * @return type
     */
    private function getColumnPrimitiveDataTypes(){
        $rowIterator = $this->excelSheet->getRowIterator();
        $rowIterator->next();   //find the second row
        $row = $rowIterator->current();
        $cellTypes = array();
        foreach($row->getCellIterator() as $cell){
            //get the type for each cell
            $cellTypes[]=$cell->getDataType();
        }
        return $cellTypes;
    }

}
?>
