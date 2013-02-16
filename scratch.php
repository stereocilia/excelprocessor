<html>
    <head><title>Hello world</title></head>
    <body>
<?php
//define(  ROOT_PATH, realpath( dirname(__FILE__) )  );

require_once 'inc\php\common.php';
    //$timeFixed = str_replace('-', '/', "05-29-08");
    //echo $timeFixed;
    //echo "<br>";
    //echo strtotime($timeFixed); //change dashes to slashes?
    echo is_date("X");
    echo "<br>";
    
    
        $timePatternB = "/^[0-9]{1,2}[AM]|[PM]$/";
        if ( ! preg_match($timePatternB, "1PM") )
        {
            echo "no pattern match";
        }

        echo "pattern match";
    
    //if( is_date( "05-29-08" ) ){
    //    echo "this is a date";
    //} else {
    //    echo "this is not a date";
    //}
?>
    </body>
</html>
