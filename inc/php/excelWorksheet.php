<?php
require_once 'common.php';
/**
 * An excel worksheet loaded into memory
 * 
 * This is a wrapper class for PHPExcel
 *
 * @author Martin Magana
 */
class excelWorkbook {
    
    public $excelWorkbook = NULL;
    public $excelWorksheet = NULL;
    public $ryExcelWorksheet = NULL;
    
    public function __construct($excelFile = NULL) {
        $this->excelWorkbook = $excelFile;
        if($this->excelWorkbook){
            //get as much information from the file as you can, here
            
            //for now, just processing the first sheet
            $this->excelWorksheet = $this->excelWorkbook->getSheet(0);

            $this->ryExcelWorksheet = $this->excelWorksheet->toArray();
            //$this->removeNullRows();
        }
    }
    
    /**
     * Tries to find the row containing the names for all the columns
     * 
     * @return Returns the index of the row that appears to conatin the column headings
     */
    public function findColumnIndex(){   //right now, just the first worksheet. it will have to eventually cycle throw all sheets
    //
        //find the first row that has all consecutive cells
        $ryDataFilledCellCounts = array();
        $columnIndex = NULL;
        
        foreach($this->ryExcelWorksheet as $row){   //go through each row
            $dataFilledCellCount = 0;
            $isConsecutive = TRUE;            
            foreach ($row as $cell){            //go through each cell
               if(  ( empty($cell) || $cell == "null" ) && $isConsecutive ){    //if the cell is considered empty AND the cells are still considered consecutive
                   $isConsecutive = FALSE;                               //then we are done doing a cell count
               } elseif($isConsecutive) {                           //if the cells are still consecutive
                   $dataFilledCellCount++;              //if the cell is not empty, count it
               }
            }
            $ryDataFilledCellCounts[] = $dataFilledCellCount;
        }
        
        //? Why did it only count four cells?
        
        //find which row had the high count of consecutive data filled cells
        //return the index of that row
        $highestCount = 0;
        foreach($ryDataFilledCellCounts as $count){
            if($count>$highestCount){
                $highestCount = $count;
            }
        }
        //now find the first occurance of the highest count
        $i = 1;
        foreach ($ryDataFilledCellCounts as $count){
            if($columnIndex === NULL){
                if($count == $highestCount){
                    $columnIndex = $i;
                }
            }
            $i++;
        }
        
        return $columnIndex;
    }
    
    public function toJSON(){
        if($this->excelWorkbook){
            
            $this->removeNullRows();
            
            $ryReturn = array();
            
            $ryReturn["dataTypes"] = $this->getColumnDataTypes();
            
            $ryReturn["excelData"] = $this->ryExcelWorksheet;                       //this will eventually be an array of sheets
            
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
        $xlSheet = $this->ryExcelWorksheet[1];
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
     * 
     * I may deprecate this function since the getDataType function of a cell isn't very useful and the iterator functions are having problems.
     * It's creating trouble and I don't real gain anything from it
     * 
     * @return type
     */
    private function getColumnPrimitiveDataTypes(){
        //TODO: For some reason this is retrieving empty cells, wtf?
        $rowIterator = $this->excelWorksheet->getRowIterator();
        $rowIterator->next();   //find the second row
        //$rowIterator->next();   //find the second row
        $row = $rowIterator->current();
        $cells = $row->getCellIterator();
        //$cells->setIterateOnlyExistingCells(false);
        $cellTypes = array();
        
        foreach($cells as $cell){
            //get the type for each cell
            $cellTypes[]=$cell->getDataType();
        }
        return $cellTypes;
    }
    
    /**
     * Removes rows that have all cells set to null in the ryExcelSheet private member
     */
    public function removeNullRows($startIndex = 0){
        if($this->ryExcelWorksheet){
            $ryExcelSheetTemp = array();
            for($i=$startIndex;$i<count($this->ryExcelWorksheet);$i++){
                if($this->ryExcelWorksheet[$i][0] == null){  //if the first cell is null, check each each cell
                    $isAllNull = TRUE;  //assume all of the cells are null unless one is found with data
                    
                    foreach($this->ryExcelWorksheet[$i] as $cell){       //now find a cell that does not have null
                        if($cell != null){
                            $isAllNull = FALSE;
                        }
                    }
                    if(!$isAllNull){ //this row shall be kept because a cell was found that was not null
                        $ryExcelSheetTemp[] = $this->ryExcelWorksheet[$i];
                    }
                } else {
                   $ryExcelSheetTemp[] = $this->ryExcelWorksheet[$i];  //this row should be kept because it the first cell was not null
                }
            }
            $this->ryExcelWorksheet = $ryExcelSheetTemp;
        }
    }
    /**
     * Finds the most occuring length of a row
     */
    private function mostCommonRowLength(){
        //get the cell count for each row
        //store in array
        //create a new array for the counts
        //loop through the original array
        //see if the original array has a key value of the value from the current loop
        //if not, create it as a key with the integer 1
        //if the original array does have the key value, increment it by 1
        //loop through the counts array
        //set the key of the first value to a var called highestCount
        //compare each value to highestCount
        //if the value is greater than the count of highestCounut, set it as the new key
        //return the count of the key with the greatest value, this is the most common row length
        return 0;   //do this for now until the function is implemented
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
