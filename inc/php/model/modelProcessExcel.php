<?php
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/PHPExcel/Classes/PHPExcel.php';  //PHPExcel Libraryrequire_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file
require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/previewSheet.php';           //allows preview data sample of Excel file

/**
 * Handles loading an Excel file
 * 
 * Create a new modelProcessExcel object. Set options that are set as public.
 * Call load to have the resulting Excel file returned.
 */
class modelProcessExcel {
    private $reader = NULL;                                                      //Excel File reader object
    private $excelFile = NULL;                                                   //Object that represent the loaded Excel File
    
    /**
     *
     * @var boolen $isPreview If set to true BEFORE calling the load function, load will return a preview of the data in the Excel file loaded. 
     */
    public $isPreview = FALSE;                                                  //True to show limited amount of rows, as set in /inc/php/previewSheet.php
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
    public function load($filePath = ""){
        //TODO: Verify that file exists

        //This finds the appropriate reader object to read the file
        $this->reader = PHPExcel_IOFactory::createReaderForFile($filePath);
        //Allow only data to be read with no style information
        //$this->reader->setReadDataOnly(true);
        //read as preview
        if($this->isPreview){
           $this->reader->setReadFilter( new previewSheet() ); 
        }
        
        //load the file into memory. this takes the longest to process.
        $this->excelFile = $this->reader->load($filePath);
        
        //TODO: visually format the values according to their data type
        return $this->excelFile;
    }
    
    /**
     * This will commit the file to a database
     */
    public function commit(){
        
    }
    
    public function __get($name)
    {
        if (array_key_exists($name, $this->data)) {
            return $this->data[$name];
        } else
            return null;
    }
}
?>
