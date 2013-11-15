# PHPExcelHelper

PHPExcelHelper encapsulates work wih the PHPExcel classes in CakePHP.

PHPExcel is a great library for the creation of spreadsheet file formats, like Excel (BIFF) .xls, Excel 2007 (OfficeOpenXML) .xlsx, CSV, Libre/OpenOffice Calc .ods, Gnumeric, PDF, HTML, ...
For more information see the PHPExcel project homepage: http://phpexcel.codeplex.com/

This class is a fork of PhpExcelHelper: <https://github.com/segy/PhpExcel>

This fork concentrates on a more complex but user friendly interface to data adding.

## CakePHP

This class is a helper class for CakePHP 2.x.

In order to use the functions, follow these steps:

1. copy the *PHPExcel* classes to your CakePHP's vendor directory:

	`app/Vendor/PHPExcel`

	Now you should have a structure like

	`app/Vendor/PHPExcel/Classes/PHPExcel/...`
2. copy the *PHPExcelHelper* class to

	`app/View/Helper/`
3. include the generation code (see example)

## Basic Code

You have to use the following basic methods:

		// start table
		$this->PHPExcel->createWorksheet(['<name>'], [<size>], [<valignment>], [<halignment>], [<offset>]);
		$this->PHPExcel->setDefaultFont(['<name>'], [<size>]);
		$this->PHPExcel->setDefaultAlignment([<valignment>], [<halignment>]);

		// header
		$header = array();
			$header[] = array(<header definition>, <col definition>);
			$header[] = array(<header definition>);
			$header[] = array(...);
		$this->PHPExcel->addTableHeader($header, [array(<row definition>)], [<filter>]);

		// data rows
		$this->PHPExcel->addTableTexts(<texts>);
		$this->PHPExcel->addTableRow(<row definition>);
		$this->PHPExcel->addTableTexts(...);
		$this->PHPExcel->addTableRow(...);

		// output
		$this->PHPExcel->output([<filename>]);

## Example

This example gives you an overview of the possible attributes.
It is not complete, for a more complete example see folder "Examples".

In your Controller add the line:

	public $helpers = array('PHPExcel');

In your View add the lines:

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
			$data[] = array('text' => 'Cell definitions override row definitions override column definitions.',
				'font-weight' => 'normal');
		$this->PHPExcel->addTableRow($data);

		// output
		$this->PHPExcel->output('Attributes.xlsx');

## Feedback

Please don't hesitate to contact me, leave issues, fork this project etc. if you want to.

