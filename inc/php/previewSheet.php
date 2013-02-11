<?php

require_once $_SERVER["DOCUMENT_ROOT"] . '/inc/php/PHPExcel/Classes/PHPExcel.php';
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
                    if ($row >= $this->startRow && $row <= $this->stopRow) {
                            return true;
                    }
                    return false;
            }
    }
?>
