<?php
/* All AJAX calls are made to this page, but handled by the controller. This is basically an init page. */

define(  ROOT_PATH, realpath( dirname(__FILE__) )  );

require_once 'inc\php\controller\controllerProcessExcel.php';

$pageController = new controllerProcessExcel();


echo ( $response = $pageController->handleRequest() ) ? $response : '{"responseStatus":"error"}';

//echo $pageController->handleRequest()
//model should return JSON error, not NULL. the model is responsible for reporting error, not this page

//NOTE: the page headers can be controlled and set as a json object
?>
