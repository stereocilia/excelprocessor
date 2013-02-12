<?php

require_once 'PHPExcel\Classes\PHPExcel.php';
    /**
     * Controls how many items are returns in a preview
     * 
     * Change the class properties $startRow and $stopRow to change the range
     */
    class previewSheet implements PHPExcel_Reader_IReadFilter
    {
            private $startRow = 1;
            private $stopRow = 10;

            public function readCell($column, $row, $worksheetName = '') {
                    // Read title row and rows 20 - 30
                    if ($row<20) {
                            return true;
                    } else return false;
            }
    }
?>
