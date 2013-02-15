<?php
require_once ROOT_PATH . '\inc\php\PHPExcel\Classes\PHPExcel.php';  //PHPExcel Libraryrequire_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file
require_once ROOT_PATH . '\inc\php\excelWorksheet.php';
require_once ROOT_PATH . '\inc\php\previewSheet.php';           //allows preview data sample of Excel file

/**
 * Handles loading an Excel file
 * 
 * Create a new modelProcessExcel object. Set options that are set as public.
 * Call load to have the resulting Excel file returned.
 */
class modelProcessExcel {
    public $columnIndex = 1;                                                    //this will be the first row returned from the preview
    public $previewLength = 10;                                                 //how many rows will be displayed with the preview
     /**
     *
     * @var boolen $isPreview If set to true BEFORE calling the load function, load will return a preview of the data in the Excel file loaded. 
     */
    public $isPreview = TRUE;                                                  //True to show limited amount of rows, as set in /inc/php/previewSheet.php
    /**
     *
     * @var array If set, the data types will define how they are entered into a database when commit is executed. If not set, the default values will be used.
     */
    public $dataType = NULL;
    /**
     * Loads an Excel file into an object in memory
     * @param string $filePath The path, including file name, to the Excel file that will be loaded
     * @return Object An object from the PHPExcel library that represents the loaded Excel file. 
     */
    public function load($excelFileToRead = ""){
        if(file_exists($excelFileToRead)){
            //This finds the appropriate reader object to read the file
            $reader = PHPExcel_IOFactory::createReaderForFile($excelFileToRead);
            if($this->isPreview){                                               //read as preview
               $reader->setReadFilter( new previewSheet($this->columnIndex, $this->columnIndex + $this->previewLength) );
            }
            $loadedWorksheet = $reader->load($excelFileToRead);                 //load the file into memory. this takes the longest to process.
            $excelWorksheet = new excelWorkbook($loadedWorksheet);             //wrap the PHPExcel file in our excelWorksheet class
            $excelWorksheet->columnIndex = $this->columnIndex;
            return $excelWorksheet;
        } else {
            return new excelWorkbook();    //returning an empty excelWorksheet will produce an error when it is converted to JSON. Since the file could not be loaded, this object represents that.
        }
    }
    
    /**
     * This will commit the file to a database
     */
    public function commit(){
        //TODO: stub out
    }
    
    /**
     * Gets private properties
     * @param type $name Name of the property
     * @return null Returns the request property
     */
    public function __get($name)
    {
        if (array_key_exists($name, $this->data)) {
            return $this->data[$name];
        } else
            return null;
    }
}
?>
