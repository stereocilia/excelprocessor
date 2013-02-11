<html>
    <head><title>Hello world</title></head>
    <body>
    <p>This is a paragraph</p>
<?php
define(  ROOT_PATH, realpath( dirname(__FILE__) )  );

require_once 'inc\php\controller\controllerProcessExcel.php';
    phpinfo();
    
    echo $_SERVER["DOCUMENT_ROOT"] . '/inc/php/controller/controllerProcessExcel.php';
?>
    </body>
</html>
