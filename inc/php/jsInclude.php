<?php define( "IS_DEV", TRUE ); //true means this is in development ?>

<?php
    $scriptSuffix = ""; //added to the name of the javascript file
    if(IS_DEV){
        $scriptSuffix = "?" . time();   //will not allow the script to be cached because it is different eveytime
    }
?>

<?php if(IS_DEV) : ?>
    <link rel="stylesheet" href="inc/js/jquery-ui-1.10.1/themes/base/jquery-ui.css" />
    <script type="text/javascript" src="inc/js/jquery-1.9.1.js"></script>
    <script type="text/javascript" src="inc/js/jquery-ui-1.10.1/ui/jquery-ui.js"></script>

<?php else : ?>
    <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.1/themes/base/jquery-ui.css" />
    <script type="text/javascript" src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
    <script src="http://code.jquery.com/ui/1.10.1/jquery-ui.js"></script>
<?php endif; ?>
    
<script type="text/javascript" src="inc/js/excelError.js<?php echo $scriptSuffix; ?>"></script>
<script type="text/javascript" src="inc/js/excelRequest.js<?php echo $scriptSuffix; ?>"></script>
<script type="text/javascript" src="inc/js/excelFile.js<?php echo $scriptSuffix; ?>"></script>
