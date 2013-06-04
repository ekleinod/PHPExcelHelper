<?php

		// start table
		$this->PHPExcel->createWorksheet();
		$this->PHPExcel->setDefaultFont('Calibri', 11);

		// header
		$header = array();
		$header[] = array('label' => 'Header 1');
		$header[] = array('label' => 'Header 2', 'font' => 'Sans', 'size' => '8', 'bold' => true, 'italic' => true, 'color' => '0000FF', 'width' => 'auto', 'filter' => true, 'wrap' => true);
		$header[] = array('label' => 'Header 3', 'size' => '15', 'width' => 75, 'filter' => false, 'wrap' => false);
		$header[] = array('label' => 'Header 4', 'bold' => false, 'italic' => false);
		$header[] = array('label' => 'Header 5');

		$this->PHPExcel->addTableHeader($header, array('bold' => true, 'size' => 10, 'filter' => true, 'color' => '008800'), 2);

		// normal rows


		// footer
		$this->PHPExcel->addTableFooter();

		// output
		$this->PHPExcel->output('Full.xlsx');

?>

