<?php
    
    if(isset($_POST['upload'])) {
        if(isset($_FILES['excel']['name']) && $_FILES['excel']['name'] != "") {
           $allowedExtensions = array("xls","xlsx");
           $ext = pathinfo($_FILES['excel']['name'], PATHINFO_EXTENSION);
           
           if(in_array($ext, $allowedExtensions)) {
                   // Uploaded file
                  $file = "uploads/".$_FILES['excel']['name'];
                  $isUploaded = copy($_FILES['excel']['tmp_name'], $file);
                  // check uploaded file
                  if($isUploaded) {
                       // Include PHPExcel files and database configuration file
                       include(__DIR__ .'/phpexcel/Classes/PHPExcel/IOFactory.php');
                       try {
                           // load uploaded file
                           $objPHPExcel = PHPExcel_IOFactory::load($file);
                       } catch (Exception $e) {
                            die('Error loading file "' . pathinfo($file, PATHINFO_BASENAME). '": ' . $e->getMessage());
                       }
                       
                       // Specify the excel sheet index
                       $sheet = $objPHPExcel->getSheet(0);
                       $total_rows = $sheet->getHighestRow();
                       $highestColumn      = $sheet->getHighestColumn();	
                       $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);		
                       
                       //	loop over the rows
                       for ($row = 1; $row <= $total_rows; ++ $row) {
                           for ($col = 0; $col < $highestColumnIndex; ++ $col) {
                               $cell = $sheet->getCellByColumnAndRow($col, $row);
                               $val = $cell->getValue();
                               $records[$row][$col] = $val;
                           }
                       }
                       foreach($records as $row){
                           	print_r($row);
                       }
                       echo "<br/>Data inserted in Database";
                   
                       unlink($file);
                   } else {
                       echo '<span class="msg">File not uploaded!</span>';
                   }
           } else {
               echo '<span class="msg">Please upload excel sheet.</span>';
           }
       } else {
           echo '<span class="msg">Please upload excel file.</span>';
       }
   }



?>