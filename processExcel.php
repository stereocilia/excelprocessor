<?php
/* All AJAX calls are made to this page, but handled by the controller. This is basically an init page. */

define(  ROOT_PATH, realpath( dirname(__FILE__) )  );

require_once 'inc\php\controller\controllerProcessExcel.php';

$pageController = new controllerProcessExcel();

echo ( $response = $pageController->handleRequest() ) ? $response : '{"responseStatus":"error"}';

//NOTE: the page headers can be controlled and set as a json object
?>
