<?php
require_once ROOT_PATH . '\inc\php\PHPExcel\Classes\PHPExcel.php';  //PHPExcel Libraryrequire_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file
require_once ROOT_PATH . '\inc\php\excelWorkbook.php';
require_once ROOT_PATH . '\inc\php\previewFilter.php';           //allows preview data sample of Excel file

/**
 * Handles loading an Excel file
 * 
 * Create a new modelProcessExcel object. Set options that are set as public.
 * Call load to have the resulting Excel file returned.
 */
class modelExcelTransaction {
    /**
     * The first row returned from the preview
     * @var integer 
     */
    public $columnHeadingIndex = 1;
    
    /**
     * How many rows will be displayed with the preview
     * @var integer 
     */
    public $previewLength = 10;
    
     /**
     * If set to true BEFORE calling the load function, load will return a preview of the data in the Excel file loaded.
     * @var boolen  
     */
    public $isPreview = TRUE;
    
    /**
     * If set, the data types will define how they are entered into a database when commit is executed. If not set, the default values will be used.
     * @var array 
     */
    public $dataType = NULL;
    
    /**
     * Loads an Excel file and return an excelWorkbook object
     * @param string $filePath The path, including file name, to the Excel file that will be loaded
     * @return excelWorkbook Represents the loaded Excel file
     */
    public function load($excelFileToRead = ""){
        if(file_exists($excelFileToRead)){
            $reader = PHPExcel_IOFactory::createReaderForFile($excelFileToRead);//This finds the appropriate reader object to read the file
            if($this->isPreview){                                               //read as preview
               $reader->setReadFilter( new previewFilter($this->columnHeadingIndex, $this->columnHeadingIndex + $this->previewLength) );
            }
            $loadedPHPExcelFile = $reader->load($excelFileToRead);              //load the file into memory. this takes the longest to process
            $workbook = new excelWorkbook($loadedPHPExcelFile);         //wrap the PHPExcel file in our excelWorkbook class
            $workbook->columnHeadingIndex = $this->columnHeadingIndex;  //set the column heading row (index)
            return $workbook;
        } else {
            return new excelWorkbook();    //returning an empty excelWorksheet will produce an error when it is converted to JSON. Since the file could not be loaded, this object represents that.
        }
    }
    /**
     * This will commit the file to a database
     * @return int Represents a status code
     */
    public function commit(excelWorkbook $objWorkBook){
        //commit the given workbook
        //TODO: PRBO - stub out
        return 0; //return fail
    }
}
?>
