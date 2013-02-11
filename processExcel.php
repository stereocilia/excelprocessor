<?php

require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/controller/controllerProcessExcel.php';

$pageController = new controllerProcessExcel();

$output = $pageController->handleRequest(); //separated into two steps for debugging
echo $output;

?>
