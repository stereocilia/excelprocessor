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
     * The first row of the preview
     * @var int 
     */
    public $previewStart = 1;
    /**
     * How many rows will be displayed with the preview
     * @var int
     */
    public $previewLength = 10;

     /**
     * If set to true BEFORE calling the load function, load will return a preview of the data in the Excel file loaded.
     * @var boolean  
     */
    public $isPreview = TRUE;
    
    /**
     * Loads an Excel file and return an excelWorkbook object
     * @param string $filePath The path, including file name, to the Excel file that will be loaded
     * @return excelWorkbook Represents the loaded Excel file
     */
    public function load($excelFileToRead = ""){
        if(file_exists($excelFileToRead)){
            $reader = PHPExcel_IOFactory::createReaderForFile($excelFileToRead);//This finds the appropriate reader object to read the file
            if($this->isPreview){                                               //read as preview
               $reader->setReadFilter( new previewFilter( $this->previewStart, $this->previewStart + $this->previewLength) );
            }
            $loadedPHPExcelFile = $reader->load($excelFileToRead);              //load the file into memory. this takes the longest to process
            $workbook = new excelWorkbook($loadedPHPExcelFile);         //wrap the PHPExcel file in our excelWorkbook class
            $workbook->excelFilePath = $excelFileToRead;
            $workbook->load();
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
        //see if an existing table already exists for the information
        //if it does, append new information to the end
        //if it does not, create a new table for it
        //insert each row of data into the table
        //report success to the user
        return 0; //return fail
    }
}
?>
