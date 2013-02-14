<?php
    /**
     * Determines if a string evaluates to a date
     * @param string $str The string to evaluate as a date
     * @return boolean TRUE if the evaluated string is a date
     */
    function is_date( $str )    //TODO: this isn't working on Idaho_MarshBird_2008_edit.xlsx
    { 
      $stamp = strtotime( $str ); 

      if (!is_numeric($stamp)) 
      { 
         return FALSE; 
      } 
      $month = date( 'm', $stamp ); 
      $day   = date( 'd', $stamp ); 
      $year  = date( 'Y', $stamp ); 

      if (checkdate($month, $day, $year)) 
      { 
         return TRUE; 
      } 

      return FALSE; 
    }
    
    function is_time($time)     //TODO: This won't find a time like 1PM
    {
        // accepts HHHH:MM:SS, e.g. 23:59:30 or 12:30 or 120:17
        $timePatternA = "/^(\-)?[0-9]{1,4}:[0-9]{1,2}(:[0-9]{1,2})?$/";
        $timePatternB = "/^[0-9]{1,2}[AM]|[PM]";
        if ( ! preg_match($timePatternA, $time) )//&& ! preg_match($timePatternB, $time) )
        {
            return false;
        }

        return true;
    }
?>
