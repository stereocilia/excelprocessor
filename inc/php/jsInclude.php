<?php define( IS_DEV, TRUE ); //true means this is in development?>
<?php
    $scriptSuffix = ""; //this is added to the name of the javascript file
    if(IS_DEV){
        $scriptSuffix = "?" . time();   //this will not allow the script to be cached because it is different eveytime
    }
?>
<script type="text/javascript" src="inc/js/jquery-1.9.1.js"></script>
<script type="text/javascript" src="inc/js/excelRequest.js<?php echo $scriptSuffix; ?>"></script>
<script type="text/javascript" src="inc/js/processExcel.js<?php echo $scriptSuffix; ?>"></script>