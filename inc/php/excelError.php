<?php

//TODO: PRBO - Make a class that inherits exception and have a function that converts it to JSON
/**
 * Creates error objects that can be converted to JSON and passed to the handling webpage
 * 
 * Instantiating an object allows you to manipulate the error in the object before it is thrown. Using throwSelfAsJSON throws a JSON with extra properties that can be handled be the UI, such as the type of error. The static function throwError allows you to throw a JSON error in a single step; esp. useful when passing an Exception as the argument.
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
        $this->ryError[self::KEYERRORMSG] = $e->getMessage() . "<br>";
        $this->ryError[self::KEYERRORMSG] .= $e->getFile() . "<br>";
        $this->ryError[self::KEYERRORMSG] .= $e->getLine() . "<br>";
        $this->ryError[self::KEYERRORMSG] .= $e->getTraceAsString() . "<br>";
    }
    
    public function loadExceptionAndThrowAsJSON(Exception $e){
        $this->loadException($e);
        $this->throwSelfAsJSON();
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
        throw self::makeJSONException($this->ryError);
    }
    
    static public function createJSONError($msg=""){
        $ryError[self::KEYERROR] = self::VALERROR;
        $ryError[self::KEYTYPE] = self::VALTYPEGENERAL;
        if( !empty($msg) ){
            $ryError[self::KEYERRORMSG] = $msg;
        } else {
            $ryError[self::KEYERRORMSG] = self::MSGDEFAULTERROR;
        }
        return $ryError;
    }
    /**
     * Used to throw a generic error
     * 
     * This has the option of being passed an Exception which has specific information about the error that occured. This information is placed into an object that can be read as JSON and then thrown.
     * 
     * @param string|Exception $error Pass a string for a generic error with message or pass and Exception for a generic error with the Exception message, file, line, and trace attached.
     * @throws Exception
     */
    static public function throwError($error = NULL){
        $ryError[self::KEYERROR] = self::VALERROR;
        $ryError[self::KEYTYPE] = self::VALTYPEGENERAL;        
        if($error){                                                             //if not null
            if(is_string($error)){                                              //if a string
                if( !empty($error) ){                                             //if string not empty
                   $ryError[self::KEYERRORMSG] = $error;                        //set error message to passed string
               } else {
                   $ryError[self::KEYERRORMSG] = self::MSGDEFAULTERROR;         //use default error if empty string
               }               
            } elseif(get_class($error) === "Exception") {                       //if an Exception
                $ryError[self::KEYERRORMSG] = $error->getMessage() . "<br>";    //add message
                $ryError[self::KEYERRORMSG] .= "<b>File:</b> " . $error->getFile() . "<br>";      //add file
                $ryError[self::KEYERRORMSG] .= "<b>Line:</b> " . $error->getLine() . "<br>";      //add line
                $ryError[self::KEYERRORMSG] .= "<b>Trace:</b><br>" . nl2br( $error->getTraceAsString() ) . "<br>";    //add trace
            } else {                                                            //not a string or exception
                $ryError[self::KEYERRORMSG] = self::MSGDEFAULTERROR;            //use default error if empty string
            }
        }
        throw self::makeJSONException($ryError);
    }
    
    static public function catchExceptionToJSON(Exception $e){
        if(strpos($e->getMessage(), excelError::SELFTHROWN)!== FALSE){
            return substr($e->getMessage(), strlen(excelError::SELFTHROWN));
        }
        return json_encode(  excelError::createJSONError( $e->getMessage() )  );
    }
    
    static private function makeJSONException(array $ryError){
        return new Exception(self::SELFTHROWN . json_encode($ryError));
    }
}

?>
