<?php
require_once 'common.php';
/**
 * This does all the processing of an Excel file.
 *
 * @author Martin Magana
 */
class excelProcessor {  //excelFile
    
    public $excelFile = NULL;
    public $excelSheet = NULL;
    public $ryExcelSheet = NULL;
    
    /**
     * Tries to find the row containing the names for all the columns
     */
    private function findColumnHeading(){
        //find the first row that has all consecutive cells
        // "stub it out", what is the return value?
    }
    
    public function toJSON(){
        //right now only processes the first sheet
        $this->excelSheet = $this->excelFile->getSheet(0);
        
        $this->ryExcelSheet = $this->excelSheet->toArray();
        //$this->removeNullRows();
        
        $ryReturn = array();
        $ryReturn["dataTypes"] = $this->getColumnDataTypes();
        //this will eventually be an array of sheets
        $ryReturn["excelData"] = $this->ryExcelSheet;
        $ryReturn["responseStatus"] = "success";
        
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
    
    //TODO: error in formatting makes rows with all empty cells appear (null). these should be removed
    /**
     * Removes rows that have all cells set to null
     */
    private function removeNullRows(){
        if($this->ryExcelSheet){
            $ryExcelSheetTemp = array();
            $cellCount = count($this->ryExcelSheet[0]);
            foreach($this->ryExcelSheet as $row){
                if($row[0] == "null"){  //if the first cell is null, check each each cell
                    $isAllNull = TRUE;
                    for($i=1;$i<$cellCount;$i++){
                        if($row[$i] != "null"){
                            $isAllNull = FALSE;
                        }
                    }
                    if(!$isAllNull){ //this row shall be kept
                        $ryExcelSheetTemp[] = $row;
                    }
                }
            }
            $this->ryExcelSheet = $ryExcelSheetTemp;
        }
    }

    public function __get($name)
    {
        if (array_key_exists($name, $this->data)) {
            return $this->data[$name];
        } else
            return null;
    }
    
    //get all the sheets
    //preview all the sheets
}

?>
