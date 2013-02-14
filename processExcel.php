<?php
/* All AJAX calls are made to this page, but handled by the controller. This is the equivelant of a bootstrap page. */

define(  ROOT_PATH, realpath( dirname(__FILE__) )  );       //this is used during development so I can move files around without worrying too much about the path

require_once 'inc\php\controller\controllerProcessExcel.php';

$pageController = new controllerProcessExcel();

//make an object with all the information exposed through headers, cookies, session, etc, and pass it to the controller

$requestData = json_decode(  stripcslashes( $_GET['data'] )  );                 //get the request data as a JSON object

//if the requestData has a column index set, set this value to the session variable
//see if there is a session variable that has a column index specified for this file.
//if it does, then set the JSON object's property to this value
//STOP! Think about this. Specifically the states, will this really work?

echo $pageController->handleRequest($requestData)

//NOTE: the page headers can be controlled and set as a json object
?>
