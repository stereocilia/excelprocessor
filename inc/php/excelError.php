<?php
/**
 * Creates error objects that can be converted to JSON and passed to the handling webpage
 *
 * @author mmagana
 */
class excelError {
    const ERRORKEY = "responseStatus";
    const ERRORVALUE = "error";
    const ERRORMSGKEY = "errorMessage";
    const ERRORMSGVALUEDEFAULT = "An undefined error has occured";
    private $ryError = array();
    
    public function __construct($msg = "") {
        $this->ryError[self::ERRORKEY] = self::ERRORVALUE;
        if( !empty($msg) ){
            $this->ryError[self::ERRORMSGKEY] = $msg;
        } else {
            $this->ryError[self::ERRORMSGKEY] = self::ERRORMSGVALUEDEFAULT;
        }
    }
    
    public function toArray(){
        return $this->ryError;
    }
    
    static public function createError($msg=""){
        $ryError[self::ERRORKEY] = self::ERRORVALUE;
        if( !empty($msg) ){
            $ryError[self::ERRORMSGKEY] = $msg;
        } else {
            $ryError[self::ERRORMSGKEY] = self::ERRORMSGVALUEDEFAULT;
        }
        return $ryError;
    }
}

?>
