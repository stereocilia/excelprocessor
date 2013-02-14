<?php
require_once 'common.php';
/**
 * This does all the processing of an Excel file.
 *
 * @author Martin Magana
 */
class excelWorkbook {  //excelFile
    
    public $excelFile = NULL;
    public $excelSheet = NULL;
    public $ryExcelSheet = NULL;
    
    public function __construct($excelFile = NULL) {
        $this->excelFile = $excelFile;
        if($this->excelFile){
            //get as much information from the file as you can, here
            
            //for now, just processing the first sheet
            $this->excelSheet = $this->excelFile->getSheet(0);

            $this->ryExcelSheet = $this->excelSheet->toArray();
            $this->removeNullRows();
        }
    }
    
    /**
     * Tries to find the row containing the names for all the columns
     * 
     * @return Returns the index of the row that appears to conatin the column headings
     */
    public function findColumnIndex(){   //right now, just the first worksheet. it will have to eventually cycle throw all sheets
        
        //find the first row that has all consecutive cells
        return 0;
    }
    
    public function toJSON(){
        if($this->excelFile){
        
            //$columnHeadingIndex = $this->findColumnHeading($this->ryExcelSheet);    //find the column heading of a single excel sheet

            $ryReturn = array();
            $ryReturn["dataTypes"] = $this->getColumnDataTypes();
            $ryReturn["excelData"] = $this->ryExcelSheet;                       //this will eventually be an array of sheets
            $ryReturn["responseStatus"] = "success";

            return json_encode($ryReturn);
        } else {
            return '{"responseStatus":"error"}';    //if there is no excel file, and error must be reported
        }
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
        $isTime = is_time($number);
        $isDate = is_date( str_replace('-', '/', $number) );    //the '-' character seems to not register with strtotime, so replacing it with '/' character
        if( $isTime ){
            return "time";
        }
        if( $isDate ){
            return "date";
        } 
        if(!$isTime && !$isDate) {
            return "number";
        }
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
    
    /**
     * Removes rows that have all cells set to null in the ryExcelSheet private member
     */
    private function removeNullRows(){
        if($this->ryExcelSheet){
            $ryExcelSheetTemp = array();
            $cellCount = count($this->ryExcelSheet[0]); //trying to get the count of total cells in a single row
            foreach($this->ryExcelSheet as $row){       //going through each row
                if($row[0] == null){  //if the first cell is null, check each each cell
                    $isAllNull = TRUE;  //assume all of the cells are null unless one is found with data
                    for($i=1;$i<$cellCount;$i++){       //now find a cell that does not have null
                        if($row[$i] != null){
                            $isAllNull = FALSE;
                        }
                    }
                    if(!$isAllNull){ //this row shall be kept
                        $ryExcelSheetTemp[] = $row;
                    }
                } else {
                   $ryExcelSheetTemp[] = $row;  //this row should be kept because it the first cell was not null
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
