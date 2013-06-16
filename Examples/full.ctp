<?php

		// start table
		$this->PHPExcel->createWorksheet();
		$this->PHPExcel->setDefaultFont('Calibri', 11);

		// header
		$header = array();
		$header[] = array('text' => 'Header 1', 'column' => array('font-weight' => 'bold'));
		$header[] = array('text' => 'Header 2', 'font-name' => 'Serif', 'font-size' => '8', 'font-weight' => 'bold', 'font-style' => 'italic', 'color' => '0000FF', 'bg-color' => '00DD00', 'width' => 'auto', 'wrap' => true);
		$header[] = array('text' => 'Header 3', 'font-size' => '15', 'width' => 75, 'wrap' => false);
		$header[] = array('text' => 'Header 4', 'font-weight' => 'bold', 'font-style' => 'normal');
		$header[] = array('text' => 'Header 5', 'wrap' => false);
		$header[] = array('text' => 'Header 6', 'column' => array('wrap' => true));
		$this->PHPExcel->addTableHeader($header, array('font-weight' => 'bold', 'font-size' => 10, 'color' => '008800', 'wrap' => true), 2, true);

		// normal rows
		$this->PHPExcel->addTableTexts(array('I am text in the first column.',
																	 'I am text in the second column.',
																	 'I am text in the third column.',
																	 'I am text in the fourth column.',
																	 'I am text in the fifth column.',
																	 'I am text in the sixth column.'));

		$data = array();
			$data[] = array('text' => 'I am text in the first column.');
			$data[] = array('text' => 'I am text in the second column.');
			$data[] = array('text' => 'I am text in the third column.');
			$data[] = array('text' => 'I am text in the fourth column.');
			$data[] = array('text' => 'I am text in the fifth column.');
			$data[] = array('text' => 'I am text in the sixth column.');
		$this->PHPExcel->addTableRow($data, array('bg-color' => 'CCCCCC'));

		$data = array();
			$data[] = array('text' => 'I am text in the first column.', 'font-weight' => 'normal');
			$data[] = array('text' => 'I am text in the second column.');
			$data[] = array();
			$data[] = array('text' => 'I am text in the fourth column.');
			$data[] = array('text' => 'I am text in the fifth column.');
			$data[] = array('text' => 'I am text in the sixth column.');
		$this->PHPExcel->addTableRow($data);

		// output
		$this->PHPExcel->output('Full.xlsx');

?>

