<?php
	/*User defined functions of PHPExcel*/
	
	//Set Sheet Title
	function SetSheetTitle($title){
		global $objPHPExcel;
		$objPHPExcel->getActiveSheet()->setTitle($title);
	}
	//Merge Two Cells
	function CellMerge($cells){
		global $objPHPExcel;
		//$objPHPExcel->mergeCells($cells);
		$objPHPExcel->getActiveSheet()->mergeCells($cells);
	}
	//Set Cell Value
	function SetCellValue($index,$cell,$text){
		global $objPHPExcel;
		$objPHPExcel->setActiveSheetIndex($index)->setCellValue($cell, $text);
	}
	//Set Font Name,size,bold,italic,unerline,color
	function SetCellFont($cell,$font_name,$font_size,$style_bold,$style_italic,$style_underline,$color_code){
		global $objPHPExcel;
	
		$styleArray = array(
             'name'      => $font_name,
			 'size'      => $font_size,
             'bold'      => $style_bold,
             'italic'    => $style_italic,
             'underline' => $style_underline,
             'color'     => array(
                 'rgb' => $color_code
             )
         );
		
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->applyFromArray($styleArray);
	}
	//Set Alignment of the Text
	function SetAlignment($cell,$hor_obj,$ver_obj){
		global $objPHPExcel;
		
		if($hor_obj == 'horizontal_left'){
			$hor_value = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;	
		}else if($hor_obj == 'horizontal_right'){
			$hor_value = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;	
		}else if($hor_obj == 'horizontal_center'){
			$hor_value = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;	
		}else if($hor_obj == 'horizontal_justify'){
			$hor_value = PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY;	
		}else if($hor_obj == ''){
			$hor_value = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;	
		}
		
		if($ver_obj == 'vertical_top'){
			$ver_value = PHPExcel_Style_Alignment::VERTICAL_TOP;	
		}else if($ver_obj == 'vertical_bottom'){
			$ver_value = PHPExcel_Style_Alignment::VERTICAL_BOTTOM;	
		}else if($ver_obj == 'vertical_center'){
			$ver_value = PHPExcel_Style_Alignment::VERTICAL_CENTER;	
		}else if($ver_obj == 'vertical_justify'){
			$ver_value = PHPExcel_Style_Alignment::VERTICAL_JUSTIFY;	
		}else if($ver_obj == ''){
			$ver_value = PHPExcel_Style_Alignment::VERTICAL_TOP;	
		}
		
		$alignArray = array(
        	'alignment' => array(
            	'horizontal' => $hor_value,
				'vertical' => $ver_value,
        	)
    	);	
		$objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray($alignArray);
	}
	//Set Row Size
	function SetRowSize($row_value,$row_height){	
	
	    global $objPHPExcel;
		
		if($row_height != ''){
			$objPHPExcel->getActiveSheet()->getRowDimension($row_value)->setRowHeight($row_height);
		}
	}
	//Set Column Size
	function SetColumnSize($col_value,$col_width){	
	
	    global $objPHPExcel;
		$cell_explode = explode(",",$col_value);
		if(count($cell_explode) > 1){
			foreach(range($cell_explode[0],$cell_explode[1]) as $columnID) {
				$objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setWidth($col_width+0.71);
			}
		}else{
			$objPHPExcel->getActiveSheet()->getColumnDimension($col_value)->setWidth($col_width+0.71);
		}
	}
	//Set Cell Wrap Text
	function SetWrapText($cell){
		global $objPHPExcel;
		
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setWrapText(true);
	}
	//Set Cell Background Color
	function SetBackgroundColor($cells,$color){
		global $objPHPExcel;
	
		$objPHPExcel->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
			'type' => PHPExcel_Style_Fill::FILL_SOLID,
			'startcolor' => array(
				 'rgb' => $color
			)
		));
	}
	//Set Cell Border
	function SetCellborder($start_cell,$end_cell,$type){
		global $objPHPExcel;
		
		if($type == ''){
			$value = PHPExcel_Style_Border::BORDER_THIN;
		}else if($type == 'border_medium'){
			$value = PHPExcel_Style_Border::BORDER_MEDIUM;
		}else if($type == 'border_thick'){
			$value = PHPExcel_Style_Border::BORDER_THICK;
		}
		$styleArray = array(
		  'borders' => array(
			'allborders' => array(
			  'style' => $value
			)
		  )
		);
	    if($end_cell != ""){
			$objPHPExcel->getActiveSheet()->getStyle($start_cell.':'.$end_cell)->applyFromArray($styleArray);
		}else{
			$objPHPExcel->getActiveSheet()->getStyle($start_cell.':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($styleArray);
		}
	}
	
	//Set Outside Border
	
	function SetOutsideborder($start_cell,$end_cell,$type){
		global $objPHPExcel;
		
		if($type == ''){
			$value = PHPExcel_Style_Border::BORDER_THIN;
		}else if($type == 'border_medium'){
			$value = PHPExcel_Style_Border::BORDER_MEDIUM;
		}else if($type == 'border_thick'){
			$value = PHPExcel_Style_Border::BORDER_THICK;
		}else if($type == 'border_none'){
			$value = PHPExcel_Style_Border::BORDER_NONE;
		}
		//if($border_type == 'outline'){
			$styleArray = array(
			  'borders' => array(
				'outline' => array(
				  'style' => $value
				)
			  )
			);
		//}
	    if($end_cell != ""){
			$objPHPExcel->getActiveSheet()->getStyle($start_cell.':'.$end_cell)->applyFromArray($styleArray);
		}else{
			$objPHPExcel->getActiveSheet()->getStyle($start_cell.':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($styleArray);
		}
	}
	
	//After Set Sum or Other Functions, use this function
	function CalculatedCellValue($cell){
		global $objPHPExcel;
		
		$objPHPExcel->getActiveSheet()->getCell($cell)->getCalculatedValue();	
	}
	//Set same value
	/*function SameCellValue($cell){
		global $objPHPExcel;
		
		$objPHPExcel->getActiveSheet()->getCell($cell)->getValue();	
	}*/
	
	//Set Background Color
	function SetBgColor($clr){
		global $objPHPExcel;
	
		$objPHPExcel->getActiveSheet()->getStyle("A1".':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($clr);	
	}
	
	
	
	//Set Freezepane
	
	function SetFreezePane($cell){
		global $objPHPExcel;
		
		$objPHPExcel->getActiveSheet()->freezePane($cell);	
	}
	
	function SetHideRow($row){
		
		
		global $objPHPExcel;
		if(is_numeric($row)){
			$objPHPExcel->getActiveSheet()->getRowDimension($row)->setVisible(false);
		}else{
			$objPHPExcel->getActiveSheet()->getColumnDimension($row)->setVisible(false);
		}
		
	}
	function SetZoomLevel($size){
		global $objPHPExcel;
		
	$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale($size);	
	}
	
	function LoadDropList($cell,$items){
		global $objPHPExcel;
		
		$objValidation = $objPHPExcel->getActiveSheet()->getCell($cell)->getDataValidation();
		$objValidation->setType(PHPExcel_Cell_DataValidation::TYPE_LIST);
		$objValidation->setErrorStyle(PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
		$objValidation->setAllowBlank(true);
		$objValidation->setShowDropDown(true);
		$objValidation->setErrorTitle('Input error');
		$objValidation->setError('Value is not in list');
		$objValidation->setFormula1('"'.$items.'"');
	}
	
	function InsertLogo($cell,$path){
		global $objPHPExcel;
		$objDrawing = new PHPExcel_Worksheet_Drawing();
		$objDrawing->setPath($path);
		$objDrawing->setCoordinates($cell);
		$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
	}
	

?>