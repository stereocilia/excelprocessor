<?php
require_once 'common.php';
require_once 'excelError.php';
require_once 'jsonKeys.php';
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
    private $_columnHeadingIndices = array();
    /**
     * The row length of each column heading index row
     * @var array 
     */
    private $_columnHeadingIndicesLength = array();
    /**
     * The row index of each sheet in _excelWorkbook member variable that is considered to be the last row of the dataset
     * @var array 
     */
    private $_lastDatasetRows = array();
    /**
     * Count of Excel Worksheets in the excelWorkbook member variable
     * @var int 
     */
    private $_sheetCount = 0;
    /**
     * Letter based index of hidden columns for each Worksheet
     * @var array 
     */
    private $_hiddenColumnIndecies = array();
    /**
     * Indicates if the load function has been called
     * @var boolean 
     */
    private $_hasLoaded = false;
    
    public function __construct(PHPExcel $PHPExcelFile = NULL) {
        $this->_excelWorkbook = $PHPExcelFile;
    }
    
    public function load(PHPExcel $PHPExcelFile = NULL){
        if($PHPExcelFile){
            $this->_excelWorkbook = $PHPExcelFile;
        }
        if($this->_excelWorkbook){
            $this->_hasLoaded = TRUE;
            $this->excelWorkbookChanged();
        } else {
            excelError::throwError( new Exception("load() function called with no excelWorkbook object available.") );
        }
        
    }
    /**
     * Call after the workbook has been modified in someway and you want to rebuild the data of the class. Basically refreshed the entire object.
     */
    private function excelWorkbookChanged(){
                
            $this->checkLoadWithException(__FUNCTION__);                        //make sure something is loaded or everything will fail
            
            
            //Remove worksheets with less that two rows, they do not represent a dataset
            $sheetIndex = 0;
            foreach($this->_excelWorkbook->getAllSheets() as $sheet){
                if(  $sheet->getHighestRow() < 2  ){
                    $this->_excelWorkbook->removeSheetByIndex($sheetIndex);
                } else {
                    $sheetIndex++;
                }
            }
            
            $this->_sheetCount = $this->_excelWorkbook->getSheetCount();
            
            //make temporary array the represents all the sheets
            $ryExcelWorksheetAssoc = array();
            foreach($this->_excelWorkbook->getWorksheetIterator() as $sheet){    //turn all sheets into arrays
                $ryExcelWorksheetAssoc[] = $sheet->toArray(null, true, true, true);
            }
            
            $ryExcelWorksheetAssocNoCols = $this->removeHiddenColumns($ryExcelWorksheetAssoc);          //remove hidden columns
            
            $this->_ryExcelWorksheets = $this->makePlainArray($ryExcelWorksheetAssocNoCols);      //remove associate keys from all arrays, set array to member variable
            
            //SUGGEST: PRBO - excelWorkbookChanged - These function should be combined. Conceptually, they are finding the range of the dataset, and the both individually loop through the sheets
            $this->findColumnHeadingIndices();
            $this->findLastDatasetRows();
            
            //-----------
            //detecting merged cells and creating an error if found
            //$mergedCells = $this->detectMergedCells();
            
            //TODO: PRBO - find the merged cells in the dataset with the new indices information
            
            for($i=0;$i < count($this->_ryExcelWorksheets);$i++){                 //loop through each worksheet
               $dataSetLength = $this->_lastDatasetRows[$i]+1 - $this->_columnHeadingIndices[$i];
               $this->_ryExcelWorksheets[$i] = array_slice($this->_ryExcelWorksheets[$i], $this->_columnHeadingIndices[$i]-1, $dataSetLength);
               
               $this->_ryExcelWorksheets[$i] = $this->removeColumnsBeyondBounds($this->_ryExcelWorksheets[$i], $this->_columnHeadingIndicesLength[$i]);   //removes data longer than the column heading length
               $this->_ryExcelWorksheets[$i] = $this->removeNullRows($this->_ryExcelWorksheets[$i]);        //remove null rows
            }
            


//            $excelError = NULL;
//            foreach ($mergedCells as $key => $isMerged){
//                if($isMerged){
//                    if($excelError === NULL){
//                        $excelError = new excelError("Merged cells detected in sheet(s) ");
//                        $excelError->setType(excelError::VALTYPEMERGEDCELLS);
//                    }
//                    $excelError->addToMessage($key+1 . " ,");
//                }
//            }
//            if($excelError !== NULL){
//                //remove the trailing comma
//                $newMessage = $excelError->getMessage();
//                $excelError->setMessage( substr_replace($newMessage, '. ', strlen($newMessage)-2) );
//                $excelError->throwSelfAsJSON();
//            }
    }
    /**
     * Check if the object is properly loaded, if not throws an exception with appropriate message
     * @param string $functionName The name of the function calling this function
     * @throws Exception
     */
    private function checkLoadWithException($functionName = ""){
        $suffix = "";
        if($functionName){
            $suffix = " from call of " . $functionName;
        }
        if(!$this->_excelWorkbook){
            throw new Exception(excelError::MSGFILENOTFOUND . $suffix );
        } elseif(!$this->_hasLoaded) {
            throw new Exception(excelError::MSGLOADNOTCALLED . $suffix );
        }

    }
    private function detectMergedCells(){
        //TODO: PRBO - FIXME: detectMergedCells checks the ENTIRE file, but should only be detecting with the dataset
        //TODO: PRBO - FIXME: detectMergedCells needs to check the entire dataset, not just the preview
        $ryMerged = array();
        foreach($this->_excelWorkbook->getAllSheets() as $sheetkey => $sheet){
            $mergedCells = $sheet->getMergeCells();
            if($mergedCells){
                foreach($mergedCells as $mergedCell){
                    $parsedCells = $this->parseMergedCellCoordinates($mergedCell);
                    $col1int = $this->col2int($parsedCells["startColumn"]);
                    $col2int = $this->col2int($parsedCells["stopColumn"]);
                    //if the stop column is larger than the length of the rows
                    //or if the 
                    $ryMerged[] = $parsedCells; //break for debugger
                }
            } 
        }
        return $ryMerged;
    }
    private function parseMergedCellCoordinates($cellCoordinates){
       //split by colin
       $ryReturn = array();
       $ryStartStop = explode(":", $cellCoordinates);
       $i=0;
       $j=0;
       while(!is_numeric($ryStartStop[0][$i])){
           $i++;
       }
       while(!is_numeric($ryStartStop[0][$j])){
           $j++;
       }
       $ryReturn["startColumn"] = substr($ryStartStop[0], 0, $i);
       $ryReturn["startRow"] = substr($ryStartStop[0], $i);
       $ryReturn["stopColumn"] = substr($ryStartStop[1], 0, $j);
       $ryReturn["stopRow"] = substr($ryStartStop[1], $j);
       return $ryReturn;
    }
    /**
     * Takes a column and changes it to an int
     * @param type $column
     */
    private function col2int($column){
        //for each character, convert to an int and add it together
        $ryInt = array();
        for($i=0;$i<strlen($column);$i++){
            $ryInt[] = (  ord($column[$i]) - 64  );
        }
        $ryInt = array_reverse($ryInt);
        
        $sum = 0;
        for($i=0;$i<strlen($column);$i++){
          if($i === 0){
              $sum += $ryInt[$i];
          } else {
              $sum += (  $ryInt[$i] * pow(26, $i)  );
          }
        }
        return $sum;
    }
    /**
     * Recurse through array elements and make all arrays integer index based
     * @param array $ry
     * @return array
     */
    private function makePlainArray(array $ry){
        foreach($ry as &$element){
            if(is_array($element)){
                $element = $this->makePlainArray($element);
            }
        }
        $ry = array_values($ry);
        return $ry;
    }
    /**
     * Find the indices of all column heading rows for all sheets in _ryExcelWorksheets member variable
     */
    private function findColumnHeadingIndices(){
            //must already have sheet count for this to work
            for($i=0;$i<$this->_sheetCount;$i++){
                $this->_columnHeadingIndices[] = $this->findColumnHeadingIndex($i);
                $columnHeadingRow = $this->_ryExcelWorksheets[$i][ ($this->_columnHeadingIndices[$i]-1) ];  //column heading row of the current sheet, as an array of cells
                $this->_columnHeadingIndicesLength[] = $this->consecutiveDataCellCount($columnHeadingRow);  //the count of consecutively filled cells of data
            }
    }
    /**
     * Finds the count of cells in a row that are consecutively filled with data (not empty or null)
     * @param array $row
     */
    private function consecutiveDataCellCount(array $row){
        if($row){
            $count = 0;   
            foreach ($row as $cell){                                        //go through each cell
                if(  $cell === null  ){
                   return $count;
               } else {
                   $count++;                                  //if the cell is not empty, count it
               }
            }
            return $count;
        } else {
            excelError::throwError( new Exception("Array length must be greater than zero") );
        }
    }
    /**
     * Tries to find the row containing the names for all the columns
     * @param int $sheetIndex Index of the sheet in the excelWorkbook member variable
     * @return int Returns the index of the row that appears to conatin the column headings
     */
    public function findColumnHeadingIndex($sheetIndex = 0){
        $columnHeadingIndex = NULL;                                             //must already have an array of excel worksheets for this to work
        
        if($sheetIndex < $this->_sheetCount){                                    //index must be within sheet count
            if(count($this->_ryExcelWorksheets[$sheetIndex])>=2){
                $ryDataFilledCellCounts = array();

                foreach($this->_ryExcelWorksheets[$sheetIndex] as $row){                                            //go through each row
                    $ryDataFilledCellCounts[] = $this->consecutiveDataCellCount($row);                              //get the count of consecutive data cells
                }
                //TODO: PRBO - error here when loading ODS file
                $mostCommonRowLength = $this->mostCommonRowLength($ryDataFilledCellCounts);                         //find length of most common row

                $columnHeadingIndex = ( (int)array_search($mostCommonRowLength, $ryDataFilledCellCounts) ) + 1;     //find the first occurance of that row, this is the index         

                return $columnHeadingIndex;
            } else {
                excelError::throwError( 
                        new Exception("Can not process a sheet with less than two rows. sheetIndex is $sheetIndex with a count of " . 
                                        count($this->_ryExcelWorksheets[$sheetIndex])
                        ) 
                );
            }
        }  else {
            excelError::throwError(new Exception("Sheet index out of bounds. sheetIndex is $sheetIndex : sheetCount is " . $this->_sheetCount));
        } 
        
    }
    /**
     * Find the index of the row that is considered to be the last of the dataset for all sheets
     */
    private function findLastDatasetRows(){
        //find the last dataset row of each sheet
        for($i=0;$i<$this->_sheetCount;$i++){
            $this->_lastDatasetRows[] = $this->findLastDatasetRow($i, $this->_columnHeadingIndices[$i]-1);
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
        if($sheetIndex < $this->_sheetCount){                                      //given index must be in range
            $sheet = $this->_ryExcelWorksheets[$sheetIndex];
            for($i=$startRow;$i<count($sheet) && $lastRow === NULL ;$i++){
                if(  ( empty($sheet[$i][0]) || $sheet[$i][0] == "null" )  ){            //if the first cell is empty or null and last row has yet been found
                    $lastRow = ($i);                                   //set the row before this one the index
                }
            }
            if($lastRow === NULL){   //set to last row if null
                $lastRow = ( count($sheet) );
            }
            return $lastRow;
        } else {
            excelError::throwError(new Exception("Sheet index out of bounds. sheetIndex is $sheetIndex : sheetCount is " . $this->_sheetCount));
        }
        
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
     * @param boolean $htmlSafe True if all characters should be optimized to display as HTML
     * @return array
     */
    public function toArray(){
        //NOTE: The array MUST be able to become a valid JSON object... this means no empty arrays!
        $ryReturn = array();
        if($this->_hasLoaded){
            for($i=0;$i<count($this->_ryExcelWorksheets);$i++){                 //loop through each worksheet
               $ryReturn[jsonKeys::excelWorksheets][$i][jsonKeys::columnTypes] = $this->getColumnDataTypes($i);         //get data types
               $ryReturn[jsonKeys::excelWorksheets][$i][jsonKeys::title] = $this->_excelWorkbook->getSheet($i)->getTitle();     //get sheet titles
               $ryReturn[jsonKeys::excelWorksheets][$i][jsonKeys::sheetData] = $this->_ryExcelWorksheets[$i];           //put the worksheet in the array
            }
            
            $ryReturn[jsonKeys::responseStatus] = jsonKeys::responseSuccess;                                                    //say everything went well
            
        } else {                                                                                        //no workbook, no array
            if(!$this->_excelWorkbook){
                $ryReturn = excelError::createJSONError("The file could not be found.");   
            } else {
                $ryReturn = excelError::createJSONError("The load() function was not called on excelWorkbook object before trying to use it.");   
            }
            
        }
        return $ryReturn;
    }
    //SUGGEST: PRBO - getColumnDataTypes - This should calculate a sample of several rows and the datatype that occurs most should be used. Example: what if the first entry is null?
    /**
     * Gets  data types of all columns for the current excelSheet of $this object
     */
    private function getColumnDataTypes($sheetIndex = 0){
        $primitiveTypes = $this->getColumnPrimitiveDataTypes($sheetIndex);      //get what the PHPExcel library says the data types are
        $cellTypes = array();
        $i = 0;
        $sheet = $this->_ryExcelWorksheets[$sheetIndex];                        //get the sheet asked for
        $row = $sheet[ $this->_columnHeadingIndices[$sheetIndex] ];           //use the row after the columnHeadingIndex for this sheet
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
        for($i=1;$i<=$this->_columnHeadingIndices[$sheetIndex];$i++){
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
     * @param array $sheet
     * @return array
     */
    public function removeNullRows(array $sheet){
        $startIndex = 0;
        if($sheet){
            $ryExcelSheetTemp = array();
            for($i=$startIndex;$i<count($sheet);$i++){
                if($sheet[$i][0] == null){  //if the first cell is null, check each each cell
                    $isAllNull = TRUE;  //assume all of the cells are null unless one is found with data
                    
                    foreach($sheet[$i] as $cell){       //now find a cell that does not have null
                        if($cell != null){
                            $isAllNull = FALSE;
                        }
                    }
                    if(!$isAllNull){ //this row shall be kept because a cell was found that was not null
                        $ryExcelSheetTemp[] = $sheet[$i];
                    }
                } else {
                   $ryExcelSheetTemp[] = $sheet[$i];  //this row should be kept because it the first cell was not null
                }
            }
            return $ryExcelSheetTemp;
        }
    }
    /**
     * Find the hidden columns in each sheet
     */
    private function findHiddenColumns(){
        foreach($this->_excelWorkbook->getAllSheets() as $sheet){
            $ryHiddenColumnIndecies = array();
            foreach($sheet->getColumnDimensions() as $dimension){
                if( !$dimension->getVisible() ){
                    $ryHiddenColumnIndecies[] = $dimension->getColumnIndex();
                }
            }
            $this->_hiddenColumnIndecies[] = $ryHiddenColumnIndecies;
        }
    }
    /**
     * Removes the columns of all sheets in excelWorkbook that are marked as hidden
     */
    private function removeHiddenColumns(array $sheets){
        $this->findHiddenColumns();
        for($i=0;$i<count($sheets);$i++){
            for($j=0;$j<count($this->_hiddenColumnIndecies[$i]);$j++){
                for($k=0;$k<count($sheets[$i]);$k++){
                    unset($sheets[$i][$k][$this->_hiddenColumnIndecies[$i][$j]]);
                }
            }
        }
        return $sheets;
    }
    /**
     * Rebuilds an array without data that is longer than the column heading row, then returns the new array
     * @param array $sheet The worksheet (as an array) that will be rebuilt
     * @param integer $bounds The amount of cells that each new row will have
     * @return array The array without columns exceeding the given length
     */
    private function removeColumnsBeyondBounds(array $sheet, $bounds){
        if($sheet){
            $ryReturn = array();
            foreach($sheet as $row){
                $ryRow = array();
                for($i=0;$i<$bounds;$i++){
                    $ryRow[] = $row[$i];
                }
                $ryReturn[] = $ryRow;
            }
            return $ryReturn;
        } else {
            excelError::throwError(new Exception("Array length must be greater than zero"));
        }
    }
    /**
     * Find the length of the most commonly occuring row
     * @param array $worksheet Excel worksheet as array
     * @return int The length of the most commonly occuring row
     */
    private function mostCommonRowLength(array $sheet){
        if($sheet){
            $sheet = array_count_values($sheet);                        //get total count of each key in array
            arsort($sheet, SORT_NUMERIC);                                   //reverse sort, keep keys
            while(  (int)key($sheet) < 2  ){                                //if this item is less than two rows
                next($sheet);                                               //get the next item, this is not the column heading index
            }
            return (int)key($sheet);                                        //return, force to int
        } else {
            excelError::throwError(new Exception("Array length must be greater than zero"));
        }
    }
}

?>
