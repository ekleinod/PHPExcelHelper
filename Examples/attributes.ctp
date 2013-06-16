<?php

		// start table
		$this->PHPExcel->createWorksheet();
		$this->PHPExcel->setDefaultFont('Calibri', 11);

		// header
		$header = array();
			$header[] = array('text' => 'Attribute', 'width' => 20, 'column' => array('font-weight' => 'bold'));
			$header[] = array('text' => 'text');
			$header[] = array('text' => 'font-name');
			$header[] = array('text' => 'font-size');
			$header[] = array('text' => 'font-weight');
			$header[] = array('text' => 'font-style');
			$header[] = array('text' => 'color');
			$header[] = array('text' => 'bg-color');
			$header[] = array('text' => 'wrap');
			$header[] = array('text' => 'width');
			$header[] = array('text' => 'column');
		$this->PHPExcel->addTableHeader($header, array('font-weight' => 'bold', 'font-size' => 10, 'width' => 'auto'));

		// normal rows
		$this->PHPExcel->addTableTexts('Values', '<text in cell>', '<name>', '<size in pt>',
																	 '"normal" or "bold" or "bolder" or "lighter"',
																	 '"normal" or "italic" or "oblique"',
																	 '<rgb>', '<rgb>', '"true" or "false"',
																	 '"auto" or <size in pt>', '<all attributes>');

		$data = array();
		$this->PHPExcel->addTableRow($data);

		$data = array();
			$data[] = array('text' => 'Remarks');
			$data[] = array();
			$data[] = array();
			$data[] = array();
			$data[] = array();
			$data[] = array('text' => 'format like "0080FF"', 'font-style' => 'italic');
			$data[] = array('text' => 'format like "0080FF"', 'font-style' => 'italic');
			$data[] = array();
			$data[] = array();
			$data[] = array('text' => 'header cells only', 'font-style' => 'italic');
			$data[] = array('text' => 'header cells only', 'font-style' => 'italic');
		$this->PHPExcel->addTableRow($data);

		$data = array();
		$this->PHPExcel->addTableRow($data);

		$data = array();
			$data[] = array('text' => 'Cell definitions override row definitions override column definitions.', 'font-weight' => 'normal');
		$this->PHPExcel->addTableRow($data);

		// output
		$this->PHPExcel->output('Attributes.xlsx');

?>

