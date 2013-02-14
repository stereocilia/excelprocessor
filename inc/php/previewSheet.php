<?php

require_once 'PHPExcel\Classes\PHPExcel.php';
    /**
     * Controls how many items are returns in a preview
     * 
     * Change the class properties $startRow and $stopRow to change the range
     */
    class previewSheet implements PHPExcel_Reader_IReadFilter
    {
            private $startRow = 0;
            private $stopRow = 0;
            
            public function __construct($startRow = 1, $stopRow = 10) {
                $this->startRow = $startRow;
                $this->stopRow = $stopRow;
            }

            public function readCell($column, $row, $worksheetName = '') {
                    if ($row >= $this->startRow && $row <= $this->stopRow) {
                        return true;
                    } else { 
                        return false;
                    }
            }
    }
?>
