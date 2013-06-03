<?php

App::uses('AppHelper', 'View/Helper');

/**
 * PHPExcelHelper encapsulates work wih the PHPExcel classes.
 *
 * PHPExcel: <http://phpexcel.codeplex.com/>
 *
 * This class is a fork of PhpExcelHelper: <https://github.com/segy/PhpExcel>
 *
 * This fork concentrates on a more complex but user friendly interface to data adding.
 *
 * @author Ekkart Kleinod https://github.com/ekleinod/
 * <https://github.com/ekleinod/PHPExcelHelper>
 *
 * @version 1.0.1
 */
class PHPExcelHelper extends AppHelper {

	/**
	 * Instance of PHPExcel class
	 * @var object
	 */
	public $xls = null;

	/**
	 * Pointer to actual row
	 * @var int
	 */
	private $row = 1;

	/**
	 * Internal table params
	 * @var array
	 */
	private $tableParams = null;

	/**
	 * Constructor
	 */
	public function __construct(View $view, $settings = array()) {
		parent::__construct($view, $settings);
	}

	/**
	 * Create new worksheet.
	 *
	 * @param theFontName default font name (optional)
	 * @param theFontSize default font size (optional)
	 */
	public function createWorksheet($theFontName = null, $theFontSize = null) {
		$this->loadEssentials();
		$this->xls = new PHPExcel();
		$this->setDefaultFont($theFontName, $theFontSize);
	}

	/**
	 * Create new worksheet from existing file
	 */
	public function loadWorksheet($path) {
		$this->loadEssentials();
		$this->xls = PHPExcel_IOFactory::load($path);
	}

	/**
	 * Load vendor classes
	 */
	private function loadEssentials() {
		App::import('Vendor', 'PHPExcel/Classes/PHPExcel');
		if (!class_exists('PHPExcel')) {
			throw new CakeException('Vendor class PHPExcel not found!');
		}
	}

	/**
	 * Set row pointer
	 */
	public function setRow($to) {
		$this->row = (int) $to;
	}

	/**
	 * Set default font.
	 *
	 * @param theFontName default font name (optional)
	 * @param theFontSize default font size (optional)
	 */
	public function setDefaultFont($theFontName = null, $theFontSize = null) {
		if ($theFontName != null) {
			$this->xls->getDefaultStyle()->getFont()->setName($theFontName);
		}
		if ($theFontSize != null) {
			$this->xls->getDefaultStyle()->getFont()->setSize($theFontSize);
		}
	}

	/**
	 * Adds a table header with the given formatting.
	 *
	 * Formatting can be given for each cell. If a formatting is given
	 * as second parameter, it is used for each cell. Individual cell formats
	 * override second-parameter-formats.
	 *
	 * Possible format keys:
	 *
	 * - *label* entry text
	 * - *font* font name
	 * - *size* font size
	 * - *bold* bold text - "true" or "false" (default)
	 * - *italic* italic text - "true" or "false" (default)
	 * - *width* column width - "auto" or units
	 * - *filter* set filter to column? - "true" or "false" (default)
	 * - *wrap*	wrap text in column? - "true" or "false" (default)
	 *
	 * @param theEntries data holding entries
	 * @param theOffset column offset
	 * @param theGlobalParams global parameters
	 */
	public function addTableHeader($theEntries, $theGlobalParams = array(), $theOffset = 0) {

		// set internal params that need to be processed after data are inserted
		$this->tableParams = array(
			'header_row' => $this->row,
			'offset' => is_numeric($theOffset) ? (int) $theOffset : PHPExcel_Cell::columnIndexFromString($theOffset),
			'row_count' => 0,
			'auto_width' => array(),
			'filter' => array(),
			'wrap' => array()
		);

		// insert entries
		$this->addTableRow($theEntries, $theGlobalParams, true);

	}

	/**
	 * Adds a table row with the given formatting.
	 *
	 * Formatting can be given for each cell. If a formatting is given
	 * as second parameter, it is used for each cell. Individual cell formats
	 * override second-parameter-formats.
	 * Some parameters can only be used for headers.
	 *
	 * Possible format keys:
	 *
	 * - *label* entry text
	 * - *font* font name
	 * - *size* font size
	 * - *bold* bold text - "true" or "false" (default)
	 * - *italic* italic text - "true" or "false" (default)
	 * - *color* text color
	 * - *width* column width - "auto" or units (headers only)
	 * - *filter* set filter to column? - "true" or "false" (default) (headers only)
	 * - *wrap*	wrap text in column? - "true" or "false" (default) (headers only)
	 * - *offset* column offset - numeric or text (headers only)
	 *
	 * @param theEntries data holding entries
	 * @param theGlobalParams global parameters
	 * @param isHeader is table row header row?
	 */
	public function addTableRow($theEntries, $theGlobalParams = array(), $isHeader = false) {

		// use global params
		foreach ($theGlobalParams as $paramKey => $paramValue) {
			foreach ($theEntries as &$entry) {
				if (!array_key_exists($paramKey, $entry)) {
					$entry[$paramKey] = $paramValue;
				}
			}
		}

		// get current column
		$currentColumn = $this->tableParams['offset'];

		// print values
		foreach ($theEntries as $entry) {

			foreach ($entry as $entryKey => $entryValue) {

				switch ($entryKey) {

					case 'label':
						$this->xls->getActiveSheet()->setCellValueByColumnAndRow($currentColumn, $this->row, $entryValue);
						break;

					case 'font':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setName($entryValue);
						break;

					case 'size':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setSize($entryValue);
						break;

					case 'bold':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setBold($entryValue);
						break;

					case 'italic':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setItalic($entryValue);
						break;

					case 'color':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->getColor()->applyFromArray(array("rgb" => $entryValue));
						break;

					case 'width':
						if ($isHeader) {
							if ($entryValue == 'auto') {
								$this->tableParams['auto_width'][] = $currentColumn;
							} else {
								$this->xls->getActiveSheet()->getColumnDimensionByColumn($currentColumn)->setWidth((float) $entryValue);
							}
						}
						break;

					case 'filter':
						if ($isHeader && $entryValue) {
							$this->tableParams['filter'][] = $currentColumn;
						}
						break;

					case 'wrap':
						if ($isHeader && $entryValue) {
							$this->tableParams['wrap'][] = $currentColumn;
						}
						break;

				}
			}

			$currentColumn++;
		}

		$this->row++;
		$this->tableParams['row_count']++;

	}

	/**
	 * End table
	 * sets params and styles that required data to be inserted
	 */
	public function addTableFooter() {
		// auto width
		foreach ($this->tableParams['auto_width'] as $col)
			$this->xls->getActiveSheet()->getColumnDimensionByColumn($col)->setAutoSize(true);
		// filter (has to be set for whole range)
		if (count($this->tableParams['filter']))
			$this->xls->getActiveSheet()->setAutoFilter(PHPExcel_Cell::stringFromColumnIndex($this->tableParams['filter'][0]).($this->tableParams['header_row']).':'.PHPExcel_Cell::stringFromColumnIndex($this->tableParams['filter'][count($this->tableParams['filter']) - 1]).($this->tableParams['header_row'] + $this->tableParams['row_count']));
		// wrap
		foreach ($this->tableParams['wrap'] as $col)
			$this->xls->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col).($this->tableParams['header_row'] + 1).':'.PHPExcel_Cell::stringFromColumnIndex($col).($this->tableParams['header_row'] + $this->tableParams['row_count']))->getAlignment()->setWrapText(true);
	}

	/**
	 * Write array of data to actual row starting from column defined by offset
	 * Offset can be textual or numeric representation
	 */
	public function addData($theEntries, $offset = 0) {
		// solve textual representation
		if (!is_numeric($offset))
			$offset = PHPExcel_Cell::columnIndexFromString($offset);

		foreach ($theEntries as $d) {
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->row, $d);
		}
		$this->row++;
	}

	/**
	 * Output file to browser
	 */
	public function output($filename = 'export.xlsx') {
		// set layout
		$this->_View->layout = '';
		// headers
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="'.$filename.'"');
		header('Cache-Control: max-age=0');
		// writer
		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel2007');
		$objWriter->save('php://output');
		// clear memory
		$this->xls->disconnectWorksheets();
	}

}

/* EOF */
