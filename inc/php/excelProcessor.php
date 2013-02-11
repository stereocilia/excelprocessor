<?php
require_once 'common.php';
/**
 * This does all the processing of an Excel file.
 *
 * @author Martin Magana
 */
class excelProcessor {
    
    public $excelFile = NULL;
    public $excelSheet = NULL;
    public $ryExcelSheet = NULL;
    
    public function toJSON(){
        //right now only processes the first sheet
        $this->excelSheet = $this->excelFile->getSheet(0);
        
        $this->ryExcelSheet = $this->excelSheet->toArray(null,true);
        
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
    
    public function __get($name)
    {
        if (array_key_exists($name, $this->data)) {
            return $this->data[$name];
        } else
            return null;
    }
}

?>
