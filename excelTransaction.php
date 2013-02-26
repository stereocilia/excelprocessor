<?php
/* All AJAX calls are made to this page, but handled by the controller. This is the equivelant of a bootstrap page. */

ob_start();                                                                     //allows catching of Xdebug errors

define(  "ROOT_PATH", realpath( dirname(__FILE__) )  );//this is used during development so I can move files around without worrying too much about the path

require_once 'inc\php\controller\controllerExcelTransaction.php';               //controller class

$requestData = json_decode(  stripcslashes( $_GET['data'] )  );                 //get the request data as a JSON object

$pageController = new controllerExcelTransaction();                             //create a controller

$output = $pageController->handleRequest($requestData);                         //give the request data to the controller

if(ob_get_length()){                                                            //if compiler tried to send output already, there was an error
    $output = json_encode(  excelError::createJSONError( ob_get_contents() )  );    //create error message out of the contents
}

ob_end_clean();                                                                 //stop caching and clear buffer

header('HTTP/1.1 200 OK');
header("Cache-Control: no-cache, must-revalidate");
header("Expires: Sat, 26 Jul 1997 05:00:00 GMT");
header('Content-Type: application/json; charset=utf-8');
header('Content-Length: ' . strlen($output));
echo $output;
?>
