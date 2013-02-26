<html>
    <head>
        <title>Hello world</title>
    </head>
    <body>
        <script type="text/javascript">
            var mystring = '{"name":"Sherie", "lastname": "mickeymaus"}';
            var myobject = JSON.parse(mystring);
            document.write(myobject.name);
        </script>
<?php
//$test = "test";
//echo substr_replace($test, 's', strlen($test)-1);
//echo ord('C')-64;
$column = 'AA';
$ryInt = array();
$int = 0;
$i = 0;
for($i=0;$i<strlen($column);$i++){
    $ryInt[] = (  ord($column[$i]) - 64  );
    //$int = (  ord($column[$i]) - 64  ) + $i * 26;
}
var_dump($ryInt);
//for($i=0;$i<strlen($column);$i++){
//    $ryInt[$i] = (   (  25 * ( strlen($column) - 1 - $i )  ) * $ryInt[$i]   ) + $ryInt[$i];
//}
//var_dump($ryInt);
//echo array_sum($ryInt);
$ryInt = array_reverse($ryInt);
$sum = 0;
for($i=0;$i<strlen($column);$i++){
  if($i === 0){
      $sum += $ryInt[$i];
  } else {
      $sum += (  $ryInt[$i] * pow(26, $i)  );
  }
}
echo $sum;
//define(  ROOT_PATH, realpath( dirname(__FILE__) )  );

//require_once 'inc\php\common.php';
    //$timeFixed = str_replace('-', '/', "05-29-08");
    //echo $timeFixed;
    //echo "<br>";
    //echo strtotime($timeFixed); //change dashes to slashes?
    //echo is_date("X");
    //echo "<br>";
    
    
        //$timePatternB = "/^[0-9]{1,2}[AM]|[PM]$/";
        //if ( ! preg_match($timePatternB, "1PM") )
        //{
        //    echo "no pattern match";
        //}

        //echo "pattern match";
    
    //if( is_date( "05-29-08" ) ){
    //    echo "this is a date";
    //} else {
    //    echo "this is not a date";
    //}
?>
    </body>
</html>
