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
    /**
     * Array of PHPExcel_Worksheet
     * @var array 
     */
    private $_excelWorksheets = NULL;
    
    /**
     * Array of Excel Worksheets represented as arrays
     * @var type 
     */
    private $_ryExcelWorksheets = array();
    
    /**
     * PHPExcel file returned from the load function of a PHPExcel_Reader class
     * 
     * @var PHPExcel 
     */
    public $excelWorkbook = NULL;
    
    public $columnHeadingIndecies = array();
    
    public $columnHeadingIndeciesLength = array();
    
    //TODO: This must be made private and accessable through getter, it's read only. For now I'll leave it so it shows up on code hinting
    public $sheetCount = 0;
    
    public function __construct(PHPExcel $PHPExcelFile = NULL) {
        $this->excelWorkbook = $PHPExcelFile;
        if($this->excelWorkbook){
            
            //TODO: PRBO - This really has to get cleaned up. It's like this for now because I'm working on it
            
            $this->removeHiddenColumns();
            
            $this->sheetCount = $this->excelWorkbook->getSheetCount();

            $this->_excelWorksheets = $this->excelWorkbook->getAllSheets();    
            
            foreach($this->excelWorkbook->getWorksheetIterator() as $sheet){    //turn all sheets into arrays
                $this->_ryExcelWorksheets[] = $sheet->toArray();
            }
            
            $this->findColumnHeadingIndecies();
        }
    }
    public function findColumnHeadingIndecies(){
            for($i=0;$i<$this->sheetCount;$i++){
                $this->columnHeadingIndecies[] = $this->findColumnHeadingIndex($i);
                //TODO: They must only be the length of consecutive data filled cells.
                $columnHeadingRow = $this->_ryExcelWorksheets[$i][ ($this->columnHeadingIndecies[$i]-1) ];  //column heading row of the current sheet, as an array of cells
                $this->columnHeadingIndeciesLength[] = $this->consecutiveDataCellCount($columnHeadingRow);  //the count of consecutively filled cells of data
            }
    }
    /**
     * Finds the count of cells in a row that are consecutively filled with data (not empty or null)
     * @param array $row
     */
    private function consecutiveDataCellCount($row){
        $dataFilledCellCount = 0;
        $isConsecutive = TRUE;            
        foreach ($row as $cell){                                        //go through each cell
           if(  ( empty($cell) || $cell == "null" ) && $isConsecutive ){//if the cell is considered empty AND the cells are still considered consecutive
               $isConsecutive = FALSE;                                  //then we are done doing a cell count
           } elseif($isConsecutive) {                                   //if the cells are still consecutive
               $dataFilledCellCount++;                                  //if the cell is not empty, count it
           }
        }
        return $dataFilledCellCount;
    }
    /**
     * Tries to find the row containing the names for all the columns
     * 
     * @return Returns the index of the row that appears to conatin the column headings
     */
    public function findColumnHeadingIndex($sheetIndex = 0){                                          //right now, just the first worksheet. it will have to eventually cycle throw all sheets
        
        $columnHeadingIndex = NULL;
        //make sure the index is withing the range of the sheet count
        if($sheetIndex < $this->sheetCount){

            //TODO: PRBO - What is the best way to maintain the state of the column index between HTTP requests?
            //NOTES: Session var? This can be overrideen by the JSON object that is passed if the user changed is

            //find the first row that has all consecutive cells
            $ryDataFilledCellCounts = array();
            
            if($this->_ryExcelWorksheets[$sheetIndex]){
                foreach($this->_ryExcelWorksheets[$sheetIndex] as $row){                          //go through each row
                    $ryDataFilledCellCounts[] = $this->consecutiveDataCellCount($row);
                }
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
                    if($columnHeadingIndex === NULL){
                        if($count == $highestCount){
                            $columnHeadingIndex = $i;
                        }
                    }
                    $i++;
                }
            }
        }
        //TODO: PRBO - Create a more accurate way of verifying this is the column heading index
        //NOTES: At this point the column heading index should be compared with the most common row length.
        //If the count of consecutive data filled cells is 15, but the most common row length is 14,
        //you may have found a row with an extra cell. In this case, finding the first row of 14 
        //consecutive data filled cells is more likely the row you want.
        return $columnHeadingIndex;
    }
    
    public function toArray(){
        //TODO: PRBO - Maybe use a better name for this. Its not just an array because many actions have been performed on it.
    }
    
    /**
     * 
     * @return string JSON string that represent the excelWorkbook of this object
     */
    public function toJSON(){
        if($this->excelWorkbook){

            $ryJSONReturn = array();
            for($i=0;$i<count($this->_ryExcelWorksheets);$i++){                 //loop through each worksheet

               
               $this->_ryExcelWorksheets[$i] = $this->setColumnHeadingIndexOfArrayWorksheet($this->_ryExcelWorksheets[$i], $this->columnHeadingIndecies[$i]);
               $this->_ryExcelWorksheets[$i] = $this->removeNullRows($this->_ryExcelWorksheets[$i]);        //remove null rows
               $this->_ryExcelWorksheets[$i] = $this->removeColumnsBeyondBounds($this->_ryExcelWorksheets[$i], $this->columnHeadingIndeciesLength[$i]);   //removes data longer than the column heading length
               
               $ryJSONReturn["excelWorksheets"][$i]["columnTypes"] = $this->getColumnDataTypes($i);         //get data types
               $ryJSONReturn["excelWorksheets"][$i]["title"] = $this->_excelWorksheets[$i]->getTitle();     //get sheet titles
               $ryJSONReturn["excelWorksheets"][$i]["sheetData"] = $this->_ryExcelWorksheets[$i];           //put the worksheet in the array
            }
            
            $ryJSONReturn["responseStatus"] = "success";                                                    //say everything went well

            return json_encode($ryJSONReturn);
            
        } else {
            $ryJSONError = array();
            $ryJSONError["responseStatus"] = "error";
            $ryJSONError["errorMessage"] = "The file could not be found";
            return json_encode($ryJSONError);                                   //if there is no excel file, and error must be reported
            
        }
    }
    
    public function setColumnHeadingIndexOfArrayWorksheet($worksheet, $columnHeadingIndex){
        $ryReturn = array();
        for($i=($columnHeadingIndex-1);$i<count($worksheet);$i++){
            $ryReturn[] = $worksheet[$i];
        }
        return $ryReturn;
    }
    
   /**
     * Gets  data types of all columns for the current excelSheet of $this object
     */
    private function getColumnDataTypes($sheetIndex = 0){
        $primitiveTypes = $this->getColumnPrimitiveDataTypes($sheetIndex);      //get what the PHPExcel library says the data types are
        $cellTypes = array();
        $i = 0;
        $sheet = $this->_ryExcelWorksheets[$sheetIndex];                        //get the sheet asked for
        $row = $sheet[ $this->columnHeadingIndecies[$sheetIndex] ];           //use the row after the columnHeadingIndex for this sheet
        foreach($row as $cell){                                                 //get the type for each cell
            switch($primitiveTypes[$i]){                                        //using a switch statement here because all values are known
                case "s" :
                    if(is_time($cell)){
                       $cellTypes[] = "time";
                    } else {
                       $cellTypes[] = "string"; 
                    }
                    
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
    /**
     * Handles types that are date, 
     * @param type $number
     * @return string
     */
    private function handleNumericType($value){
        $isTime = is_time($value);
        $isDate = is_date( str_replace('-', '/', $value) );                    //the '-' character seems to not register with strtotime, so replacing it with '/' character
        if( $isTime ){
            return "time";
        }
        if( $isDate ){
            return "date";
        } 
        if(!$isTime && !$isDate) {
            if( is_numeric($value) ){
              return "number";  
            } else {
                return "string";
            }
            
        }
    }
    
    //NOTE: Should I really rely on the data type that Excel gives me? Wouldn't it be more accurate to find the data types on my own?
    /**
     * Gets primative data types of all columns for the current excelSheet of $this object
     * 
     * @return array The data types for each cell as given by PHPExcel_Cell->getDataType()
     */
    private function getColumnPrimitiveDataTypes($sheetIndex = 0){
        $rowIterator = $this->_excelWorksheets[$sheetIndex]->getRowIterator();
        //get the iterator one row after the column index
        for($i=1;$i<=$this->columnHeadingIndecies[$sheetIndex];$i++){
            $rowIterator->next();
        }
        //get the generic data types for each cell and store them in an array
        $cellTypes = array();
        foreach($rowIterator->current()->getCellIterator() as $cell){
            //get the type for each cell
            $cellTypes[]=$cell->getDataType();
        }
        return $cellTypes;
    }
    
    /**
     * Removes rows that have all cells set to null in the ryExcelSheet private member
     */
    public function removeNullRows($ryExcelWorksheet){
        $startIndex = 0;
        if($ryExcelWorksheet){
            $ryExcelSheetTemp = array();
            for($i=$startIndex;$i<count($ryExcelWorksheet);$i++){
                if($ryExcelWorksheet[$i][0] == null){  //if the first cell is null, check each each cell
                    $isAllNull = TRUE;  //assume all of the cells are null unless one is found with data
                    
                    foreach($ryExcelWorksheet[$i] as $cell){       //now find a cell that does not have null
                        if($cell != null){
                            $isAllNull = FALSE;
                        }
                    }
                    if(!$isAllNull){ //this row shall be kept because a cell was found that was not null
                        $ryExcelSheetTemp[] = $ryExcelWorksheet[$i];
                    }
                } else {
                   $ryExcelSheetTemp[] = $ryExcelWorksheet[$i];  //this row should be kept because it the first cell was not null
                }
            }
            return $ryExcelSheetTemp;
        }
    }
    
    private function removeHiddenColumns(){
        //TODO: PRBO - If the last column is removed, data in the last row reamins for some reason
        foreach($this->excelWorkbook->getAllSheets() as $sheet){
            foreach($sheet->getColumnDimensions() as $dimension){
                if( !$dimension->getVisible() ){
                    $sheet->removeColumn( $dimension->getColumnIndex() );
                }
            }
        }
    }
    
    /**
     * Rebuilds an array without data that is longer than the column heading row, then returns the new array
     * @param array $ryExcelWorksheet The worksheet (as an array) that will be rebuilt
     * @param integer $bounds The amount of cells that each new row will have
     * @return array
     */
    private function removeColumnsBeyondBounds($ryExcelWorksheet, $bounds){
        $ryReturn = array();
        foreach($ryExcelWorksheet as $row){
            $ryRow = array();
            for($i=0;$i<$bounds;$i++){
                $ryRow[] = $row[$i];
            }
            $ryReturn[] = $ryRow;
        }
        return $ryReturn;
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
