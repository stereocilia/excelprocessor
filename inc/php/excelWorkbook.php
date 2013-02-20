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
     * 
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
            $mostCommonRowLength = $this->mostCommonRowLength($ryDataFilledCellCounts);

            $columnHeadingIndex = $this->firstRowOf($ryDataFilledCellCounts, $mostCommonRowLength);             //find the first occurance of that row, this is the index
        }   
        return $columnHeadingIndex;
    }
    
    /**
     * Finds the index that a row or a specific length occurs
     * @param array $worksheet A worksheet as an array
     * @param int $length The length of the row to find. This is the count of cells.
     * @return int
     */
    private function firstRowOf($worksheet, $length){
        $firstOccuranceIndex = NULL;
        if( is_array($worksheet) && is_integer($length) ){
                $i = 1;
                foreach ($worksheet as $value){
                    if(  $firstOccuranceIndex === NULL && $value == $length  ){ //if first occurance not set and this length matches the given length
                        $firstOccuranceIndex = $i;
                        return $firstOccuranceIndex;
                    }
                    $i++;
                }
        }
        return $firstOccuranceIndex;
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
               $ryReturn["excelWorksheets"][$i]["title"] = $this->excelWorkbook->getSheet($i)->getTitle();     //get sheet titles
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
        $rowIterator = $this->excelWorkbook->getSheet($sheetIndex)->getRowIterator();
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
        //TODO: PRBO - FIXME: If the last column is removed, data in the last row remains for some reason
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
    private function mostCommonRowLength($worksheet){
        if( is_array($worksheet) ){
            $worksheet = array_count_values($worksheet);
            arsort($worksheet, SORT_NUMERIC);
            return key($worksheet);
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
