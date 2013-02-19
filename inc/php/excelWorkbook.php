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
    
    private $columnHeadingIndices = array();
    
    private $columnHeadingIndicesLength = array();
    
    private $lastDatasetRows = array();
    
    //TODO: This must be made private and accessable through getter, it's read only. For now I'll leave it so it shows up on code hinting
    private $sheetCount = 0;
    
    public function __construct(PHPExcel $PHPExcelFile = NULL) {
        $this->excelWorkbook = $PHPExcelFile;
        if($this->excelWorkbook){
            
            $this->removeHiddenColumns();
            
            $this->sheetCount = $this->excelWorkbook->getSheetCount();

            $this->_excelWorksheets = $this->excelWorkbook->getAllSheets();    
            
            foreach($this->excelWorkbook->getWorksheetIterator() as $sheet){    //turn all sheets into arrays
                $this->_ryExcelWorksheets[] = $sheet->toArray();
            }
            
            $this->findColumnHeadingIndices();
        }
    }
    public function findColumnHeadingIndices(){
            //must already have sheet count for this to work
            for($i=0;$i<$this->sheetCount;$i++){
                $this->columnHeadingIndices[] = $this->findColumnHeadingIndex($i);
                //TODO: They must only be the length of consecutive data filled cells.
                $columnHeadingRow = $this->_ryExcelWorksheets[$i][ ($this->columnHeadingIndices[$i]-1) ];  //column heading row of the current sheet, as an array of cells
                $this->columnHeadingIndicesLength[] = $this->consecutiveDataCellCount($columnHeadingRow);  //the count of consecutively filled cells of data
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
        //must already have an array of excel worksheets for this to work
        $columnHeadingIndex = NULL;
        //make sure the index is withing the range of the sheet count
        if($sheetIndex < $this->sheetCount){

            //TODO: PRBO - What is the best way to maintain the state of the column index between HTTP requests? Or should it just redectect every time?
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
    
    private function findLastDatasetRows(){
        //find the last dataset row of each sheet
        for($i=0;$i<$this->sheetCount;$i++){
            $this->lastDatasetRows[] = $this->findLastDatasetRow($i);
        }
       
    }
    
    private function findLastDatasetRow($sheetIndex = 0){
        //assumes that the first row is the column heading row. setColumnHeadinIndexOfArrayWorksheet() should be called first
        //find the first row after the column heading row where the first cell is empty. this will be the last row of the dataset
        $lastDatasetRow = NULL;
        if($sheetIndex<$this->sheetCount){                                      //given index must be in range
            $sheet = $this->_ryExcelWorksheets[$sheetIndex];
            foreach($sheet as $index => $row){
                if(  ( empty($row[0]) || $row[0] == "null" ) && $lastDatasetRow === NULL ){            //if the first cell is empty or null and last row has yet been found
                    $lastDatasetRow = ($index - 1);                                   //set the row before this one the index
                }
            }
        }
        if($lastDatasetRow === NULL){   //set to last row if null
            $lastDatasetRow = (  count( $this->_ryExcelWorksheets[$sheetIndex] ) - 1  );
        }
        return $lastDatasetRow;
    }
    
    public function toArray(){
        //NOTE: The array MUST be able to become a valid JSON object... this means no empty arrays!
        $ryReturn = array();
        if($this->excelWorkbook){
            for($i=0;$i<count($this->_ryExcelWorksheets);$i++){                 //loop through each worksheet
               //TODO: PRBO - Should the setColumnHeading... and setLastDataset... be called together in a function since the both reorganize the data?
               $this->_ryExcelWorksheets[$i] = $this->setColumnHeadingIndexOfArrayWorksheet($this->_ryExcelWorksheets[$i], $this->columnHeadingIndices[$i]);//make the first row the column heading row
               //TODO: Why doesn't this work if called in the constructor function?
               $this->findLastDatasetRows();    //does this need to be called here or somewhere else?
               $this->_ryExcelWorksheets[$i] = $this->setLastDatasetRowOfArrayWorksheet($this->_ryExcelWorksheets[$i], $this->lastDatasetRows[$i]);
               
               $this->_ryExcelWorksheets[$i] = $this->removeColumnsBeyondBounds($this->_ryExcelWorksheets[$i], $this->columnHeadingIndicesLength[$i]);   //removes data longer than the column heading length
               $this->_ryExcelWorksheets[$i] = $this->removeNullRows($this->_ryExcelWorksheets[$i]);        //remove null rows
               
               
               $ryReturn["excelWorksheets"][$i]["columnTypes"] = $this->getColumnDataTypes($i);         //get data types
               $ryReturn["excelWorksheets"][$i]["title"] = $this->_excelWorksheets[$i]->getTitle();     //get sheet titles
               $ryReturn["excelWorksheets"][$i]["sheetData"] = $this->_ryExcelWorksheets[$i];           //put the worksheet in the array
            }
            
            $ryReturn["responseStatus"] = "success";                                                    //say everything went well
            
        } else {
            $ryReturn["responseStatus"] = "error";
            $ryReturn["errorMessage"] = "The file could not be found";
        }
        return $ryReturn;
    }
    
    private function setColumnHeadingIndexOfArrayWorksheet($worksheet, $columnHeadingIndex){
        $ryReturn = array();
        for($i=($columnHeadingIndex-1);$i<count($worksheet);$i++){
            $ryReturn[] = $worksheet[$i];
        }
        return $ryReturn;
    }
    
    private function setLastDatasetRowOfArrayWorksheet($worksheet, $lastDatasetRow){
        //rebuild array backward
        $ryReturn = array();
        for($i=( $lastDatasetRow );$i>=0;$i--){
            $ryReturn[] = $worksheet[$i];
        }
        return array_reverse($ryReturn);
    }
    
   /**
     * Gets  data types of all columns for the current excelSheet of $this object
     */
    private function getColumnDataTypes($sheetIndex = 0){
        $primitiveTypes = $this->getColumnPrimitiveDataTypes($sheetIndex);      //get what the PHPExcel library says the data types are
        $cellTypes = array();
        $i = 0;
        $sheet = $this->_ryExcelWorksheets[$sheetIndex];                        //get the sheet asked for
        $row = $sheet[ $this->columnHeadingIndices[$sheetIndex] ];           //use the row after the columnHeadingIndex for this sheet
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
        for($i=1;$i<=$this->columnHeadingIndices[$sheetIndex];$i++){
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
        //TODO: PRBO - If the last column is removed, data in the last row remains for some reason
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
        //TODO: PRBO - mostCommonRowLength - Create the logic for this function
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
