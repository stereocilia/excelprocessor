<?php

//TODO: PRBO - Make a class that inherits exception and have a function that converts it to JSON
/**
 * Creates error objects that can be converted to JSON and passed to the handling webpage
 *
 * @author Martin Magana <magana.web@gmail.com>
 */
class excelError {
    const KEYERROR = "responseStatus";
    const KEYTYPE = "errorType";
    const KEYERRORMSG = "errorMessage";
    
    const VALERROR = "error";
    const VALTYPEGENERAL = "general";
    const VALTYPEMERGEDCELLS = "mergedcells";
    
    const MSGDEFAULTERROR = "An undefined error has occured.";
    const MSGFILENOTFOUND = "The file could not be found.";
    const MSGLOADNOTCALLED = "The load() function was not called on excelWorkbook object before trying to use it.";
    
    const SELFTHROWN = "<SELFTHROWN>";
    
    private $ryError = array();
    
    private $errorTypes = array();
    
    public function __construct($msg = "") {
        $this->errorTypes[self::VALTYPEGENERAL] = self::KEYTYPE; 
        $this->errorTypes[self::VALTYPEMERGEDCELLS] = self::KEYTYPE; 
        
        $this->ryError[self::KEYERROR] = self::VALERROR;
        $this->ryError[self::KEYTYPE] = self::VALTYPEGENERAL;
        if( !empty($msg) ){
            $this->ryError[self::KEYERRORMSG] = $msg;
        } else {
            $this->ryError[self::KEYERRORMSG] = self::MSGDEFAULTERROR;
        }
    }
    
    public function loadException(Exception $e){
        $this->ryError[self::KEYERRORMSG] = $e->getMessage();
        $this->ryError[self::KEYERRORMSG] .= $e->getFile() . "<br>";
        $this->ryError[self::KEYERRORMSG] .= $e->getLine() . "<br>";
        $this->ryError[self::KEYERRORMSG] .= $e->getTraceAsString() . "<br>";
    }
    
    public function addToMessage($msg){
        $this->ryError[self::KEYERRORMSG] .= $msg;
    }
    
    public function setMessage($msg){
        $this->ryError[self::KEYERRORMSG] = $msg;
    }
    
    public function getMessage(){
        return $this->ryError[self::KEYERRORMSG];
    }
    
    public function setType($type = ""){
        if(array_key_exists($type, $this->errorTypes)){
            $this->ryError[self::KEYTYPE] = $type;
        } else {
            throw new Exception("Exception: supplied error type $type does not exist");
        }
    }
    
    public function toArray(){
        return $this->ryError;
    }
    
    public function throwSelf(){
        throw new Exception($this->ryError[self::KEYERRORMSG]);
    }
    
    public function throwSelfAsJSON(){
        throw new Exception("<SELFTHROWN>" . json_encode($this->ryError));
    }
    
    static public function createJSONError($msg=""){
        $ryError[self::KEYERROR] = self::VALERROR;
        $this->ryError[self::KEYTYPE] = self::VALTYPEGENERAL;
        if( !empty($msg) ){
            $ryError[self::KEYERRORMSG] = $msg;
        } else {
            $ryError[self::KEYERRORMSG] = self::MSGDEFAULTERROR;
        }
        return $ryError;
    }
}

?>
