<?php
require_once 'common.php';
require_once 'excelError.php';
/**
 * An excel worksheet loaded into memory
 * 
 * This is a wrapper class for PHPExcel
 *
 * @author Martin Magana
 */
class excelWorkbook {
    /**
     * PHPExcel file returned from the load function of a PHPExcel_Reader class
     * 
     * @var PHPExcel 
     */
    private $_excelWorkbook = NULL;
    /**
     * Array of Excel Worksheets represented as arrays
     * @var type 
     */
    private $_ryExcelWorksheets = array();    
    /**
     * Array of all indecies of column heading rows for every sheets in excelWorkbook member variable
     * @var array 
     */
    private $columnHeadingIndices = array();
    /**
     * The row length of each column heading index row
     * @var array 
     */
    private $columnHeadingIndicesLength = array();
    /**
     * The row index of each sheet in _excelWorkbook member variable that is considered to be the last row of the dataset
     * @var array 
     */
    private $lastDatasetRows = array();
    //TODO: PRBO - sheetCount should be available through getter function, effectively making it read only to the outside
    /**
     * Count of Excel Worksheets in the excelWorkbook member variable
     * @var int 
     */
    private $sheetCount = 0;
    
    public function __construct(PHPExcel $PHPExcelFile = NULL) {
        $this->_excelWorkbook = $PHPExcelFile;
        if($this->_excelWorkbook){
            
            $this->removeHiddenColumns();
            
            $this->sheetCount = $this->_excelWorkbook->getSheetCount();   
            
            foreach($this->_excelWorkbook->getWorksheetIterator() as $sheet){    //turn all sheets into arrays
                $this->_ryExcelWorksheets[] = $sheet->toArray();
            }
            
            $this->findColumnHeadingIndices();
        }
    }
    /**
     * Find the indices of all column heading rows for all sheets in _ryExcelWorksheets member variable
     */
    private function findColumnHeadingIndices(){
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
        //skips verifying if $row is an array for a little extra speed
        $count = 0;   
        foreach ($row as $cell){                                        //go through each cell
            if(  $cell === null  ){
               return $count;
           } else {
               $count++;                                  //if the cell is not empty, count it
           }
        }
        return $count;
    }
    
    //TODO: PRBO - findColumnHeadingIndex - Should I be passing the sheet index or the sheet itself?
    /**
     * Tries to find the row containing the names for all the columns
     * @param int $sheetIndex Index of the sheet in the excelWorkbook member variable
     * @return int Returns the index of the row that appears to conatin the column headings
     */
    public function findColumnHeadingIndex($sheetIndex = 0){
        $columnHeadingIndex = NULL;                                             //must already have an array of excel worksheets for this to work
        
        if($sheetIndex < $this->sheetCount){                                    //index must be within sheet count
            //TODO: PRBO - What is the best way to maintain the state of the column index between HTTP requests? Or should it just redectect every time?
            //NOTES: Session var? This can be overrideen by the JSON object that is passed if the user changed is

            $ryDataFilledCellCounts = array();

            foreach($this->_ryExcelWorksheets[$sheetIndex] as $row){                                            //go through each row
                $ryDataFilledCellCounts[] = $this->consecutiveDataCellCount($row);                              //get the count of consecutive data cells
            }
            $mostCommonRowLength = $this->mostCommonRowLength($ryDataFilledCellCounts);                         //find length of most common row
            
            $columnHeadingIndex = ( (int)array_search($mostCommonRowLength, $ryDataFilledCellCounts) ) + 1;     //find the first occurance of that row, this is the index         
        }   
        return $columnHeadingIndex;
    }
    /**
     * Find the index of the row that is considered to be the last of the dataset for all sheets
     */
    private function findLastDatasetRows(){
        //find the last dataset row of each sheet
        for($i=0;$i<$this->sheetCount;$i++){
            $this->lastDatasetRows[] = $this->findLastDatasetRow($i);
        }
       
    }
    /**
     * Find the index of the row that is considered to be the last of the dataset for the given sheet
     * @param int $sheetIndex Index of the sheet to use for member variable _excelWorkbook
     * @return int
     */
    private function findLastDatasetRow($sheetIndex = 0, $startRow = 0){
        //find the first row after the column heading row where the first cell is empty. this will be the last row of the dataset
        $lastRow = NULL;
        if($sheetIndex<$this->sheetCount){                                      //given index must be in range
            $sheet = $this->_ryExcelWorksheets[$sheetIndex];
            for($i=$startRow;$i<count($sheet) && $lastRow === NULL ;$i++){
                if(  ( empty($sheet[$i][0]) || $sheet[$i][0] == "null" ) && $lastRow === NULL ){            //if the first cell is empty or null and last row has yet been found
                    $lastRow = ($i - 1);                                   //set the row before this one the index
                }
            }
            if($lastRow === NULL){   //set to last row if null
                $lastRow = ( count($sheet) - 1 );
            }
        }
        return $lastRow;
    }
    /**
     * Creates an array that represents the Excel Worksheet. This array can be encoded to a JSON object
     * 
     * The first property is excelWorksheets, which is an array all of the worksheets in the workbook
     * JSONObject.excelWorksheets[0] would be the first worksheet.
     * Each worksheet has the properties columnTypes, title, and sheetData.
     * JSONObject.excelWorksheets[0].columnTypes    - Array containing the types of each column
     * JSONObject.excelWorksheets[0].title          - Title of this worksheet
     * JSONObject.excelWorksheets[0].sheetData      - The data of the actual worksheet (array of rows)
     * Also has a response status to tell success or failure of the HTTP response.
     * JSONObject.responseStatus can be "success" or "error"
     * If the response status is an error then it can contain a message
     * JSONObject.errorMessage                      - Message associated with this error.
     * 
     * @return array
     */
    public function toArray(){
        //NOTE: The array MUST be able to become a valid JSON object... this means no empty arrays!
        $ryReturn = array();
        if($this->_excelWorkbook){
            for($i=0;$i<count($this->_ryExcelWorksheets);$i++){                 //loop through each worksheet
                
                //TODO: PRBO - toArray - This stuff is called in the wrong place
               $this->_ryExcelWorksheets[$i] = array_slice($this->_ryExcelWorksheets[$i], $this->columnHeadingIndices[$i]-1);
               $this->findLastDatasetRows();
               $this->_ryExcelWorksheets[$i] = $this->setLastDatasetRowOfArrayWorksheet($this->_ryExcelWorksheets[$i], $this->lastDatasetRows[$i]);
               
               $this->_ryExcelWorksheets[$i] = $this->removeColumnsBeyondBounds($this->_ryExcelWorksheets[$i], $this->columnHeadingIndicesLength[$i]);   //removes data longer than the column heading length
               $this->_ryExcelWorksheets[$i] = $this->removeNullRows($this->_ryExcelWorksheets[$i]);        //remove null rows
               
               
               $ryReturn["excelWorksheets"][$i]["columnTypes"] = $this->getColumnDataTypes($i);         //get data types
               $ryReturn["excelWorksheets"][$i]["title"] = $this->_excelWorkbook->getSheet($i)->getTitle();     //get sheet titles
               $ryReturn["excelWorksheets"][$i]["sheetData"] = $this->_ryExcelWorksheets[$i];           //put the worksheet in the array
            }
            
            $ryReturn["responseStatus"] = "success";                                                    //say everything went well
            
        } else {                                                                                        //no workbook, no array
            $ryReturn = excelError::createError("The file could not be found.");
        }
        return $ryReturn;
    }
    //TODO: PRBO - setColumnHeadingIndex: Is there a PHP function that already does this? Seems like there would be.
    /**
     * Sets the first item on the given worksheet array to the given column heading index
     * @param array $worksheet The array to set the index of
     * @param int $columnHeadingIndex Index of the array element that will now be first
     * @return array
     */
    private function setColumnHeadingIndex($worksheet, $columnHeadingIndex){
        $ryReturn = array();
        for($i=($columnHeadingIndex-1);$i<count($worksheet);$i++){
            $ryReturn[] = $worksheet[$i];
        }
        return $ryReturn;
    }
    /**
     * Removes every element of the given array after the given index
     * @param array $worksheet
     * @param int $lastDatasetRow
     * @return array
     */
    private function setLastDatasetRowOfArrayWorksheet($worksheet, $lastDatasetRow){    //rebuild array backward
        $ryReturn = array();
        for( $i=($lastDatasetRow) ; $i>=0 ; $i-- ){
            $ryReturn[] = $worksheet[$i];
        }
        return array_reverse($ryReturn);
    }
    //TODO: PRBO - getColumnDataTypes - This should calculate a sample of several rows and the datatype that occurs most should be used. Example: what if the first entry is null?
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
     * @param ing $number
     * @return string A string that describes the type
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
     * @param int $sheetIndex Index of the sheet to process from the excelWorkbook member variable
     * @return array The data types for each cell as given by PHPExcel_Cell->getDataType()
     */
    private function getColumnPrimitiveDataTypes($sheetIndex = 0){
        $rowIterator = $this->_excelWorkbook->getSheet($sheetIndex)->getRowIterator();
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
     * Removes rows that have a data type of null for every cell
     * @param array $ryExcelWorksheet
     * @return array
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
    /**
     * Removes the columns of all sheets in excelWorkbook that are marked as hidden
     */
    private function removeHiddenColumns(){
        //TODO: PRBO - FIXME: If the last column is removed, data in the last row remains for some reason
        foreach($this->_excelWorkbook->getAllSheets() as $sheet){
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
     * @return array The array without columns exceeding the given length
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
     * Find the length of the most commonly occuring row
     * @param array $worksheet Excel worksheet as array
     * @return int The length of the most commonly occuring row
     */
    private function mostCommonRowLength($worksheet){
        if( is_array($worksheet) ){                                             //make sure this is an array
            $worksheet = array_count_values($worksheet);                        //get total count of each key in array
            arsort($worksheet, SORT_NUMERIC);                                   //reverse sort, keep keys
            while(  (int)key($worksheet) < 2  ){                                //if this item is less than two rows
                next($worksheet);                                               //get the next item, this is not the column heading index
            }
            return (int)key($worksheet);                                        //return, force to int
        }
        //TODO: PRBO - mostCommonRowLength - Unhandled error: If the passed value is not an array, the function does not return a value
    }
}

?>
