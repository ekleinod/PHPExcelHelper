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

## Example

In your Controller add the line:

	public $helpers = array('PHPExcel');

In your View add the lines:

	// create worksheet with default font
	$this->PhpExcel->createWorksheet('Calibri', 12);

	// define table cells
	$table = array(
		array('label' => __('User'), 'width' => 'auto', 'filter' => true),
		array('label' => __('Type'), 'width' => 'auto', 'filter' => true),
		array('label' => __('Date'), 'width' => 'auto'),
		array('label' => __('Description'), 'width' => 50, 'wrap' => true),
		array('label' => __('Modified'), 'width' => 'auto')
	);

	// heading
	$this->PhpExcel->addTableHeader($table, array('name' => 'Cambria', 'bold' => true));

	// data
	foreach ($data as $d) {
		$this->PhpExcel->addTableRow(array(
			$d['User']['name'],
			$d['Type']['name'],
			$d['User']['date'],
			$d['User']['description'],
			$d['User']['modified']
		));
	}

	$this->PhpExcel->addTableFooter();
	$this->PhpExcel->output();

## Feedback

Please don't hesitate to contact me, leave issues, fork this project etc. if you want to.

