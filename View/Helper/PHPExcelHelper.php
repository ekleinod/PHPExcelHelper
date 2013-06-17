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
		$this->setRow(1);
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
	 * Set row pointer.
	 *
	 * @param theRow row to set to
	 */
	public function setRow($theRow) {
		$this->row = (int) $theRow;
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
	 * Formatting can be given for each cell. If a row formatting is given
	 * it is used for each cell in this row.
	 *
	 * Each header definition may contain a column array of parameters for the entire column.
	 *
	 * Row formats override column formats.
	 * Individual cell formats override column and row formats.
	 *
	 * Possible format keys:
	 *
	 * - *text* entry text
	 * - *font-name* font name
	 * - *font-size* font size
	 * - *font-weight* font weight - "normal", "bold" or "bolder" or "lighter"
	 * - *font-style* font style - "normal", "italic" or "oblique"
	 * - *color* text color
	 * - *bg-color* background color
	 * - *wrap* wrap text - "true" or "false" (default)
	 * - *width* column width - "auto" or units
	 * - *column* formatting parameters (array) for entire column
	 *
	 * @param theEntries data holding entries
	 * @param theRowParams parameters for entire row
	 * @param theColOffset column offset
	 * @param theFilter switch on filter?
	 */
	public function addTableHeader($theEntries, $theRowParams = array(), $theColOffset = 0, $theFilter = false) {

		// set internal params that need to be processed after data are inserted
		$this->tableParams = array(
			'header_row' => $this->row,
			'col_offset' => is_numeric($theColOffset) ? (int) $theColOffset : PHPExcel_Cell::columnIndexFromString($theColOffset),
			'col_count' => 0,
			'row_count' => 0,
			'filter' => ($theFilter == true),
			'col_params' => array(),
			'auto_width' => array()
		);

		// store col params, set or store width
		$currentColumn = $this->tableParams['col_offset'];
		foreach ($theEntries as $entry) {

			// column parameters
			if (array_key_exists('column', $entry)) {
				$this->tableParams['col_params'][$currentColumn] = $entry['column'];
			}

			// column width
			$this->tableParams['auto_width'][$currentColumn] = false;

			if (array_key_exists('width', $theRowParams)) {
				if ($theRowParams['width'] == 'auto') {
					$this->tableParams['auto_width'][$currentColumn] = true;
				} else {
					$this->tableParams['auto_width'][$currentColumn] = false;
					$this->xls->getActiveSheet()->getColumnDimensionByColumn($currentColumn)->setWidth((float) $theRowParams['width']);
				}
			}

			if (array_key_exists('width', $entry)) {
				if ($entry['width'] == 'auto') {
					$this->tableParams['auto_width'][$currentColumn] = true;
				} else {
					$this->tableParams['auto_width'][$currentColumn] = false;
					$this->xls->getActiveSheet()->getColumnDimensionByColumn($currentColumn)->setWidth((float) $entry['width']);
				}
			}

			$currentColumn++;
		}

		// insert entries
		$this->addTableRow($theEntries, $theRowParams);

	}

	/**
	 * Adds a table row with the given formatting.
	 *
	 * Formatting can be given for each cell. If a row formatting is given
	 * it is used for each cell in this row.
	 *
	 * Row formats override column formats.
	 * Individual cell formats override column and row formats.
	 *
	 * Possible format keys:
	 *
	 * - *text* entry text
	 * - *font-name* font name
	 * - *font-size* font size
	 * - *font-weight* font weight - "normal", "bold" or "bolder" or "lighter"
	 * - *font-style* font style - "normal", "italic" or "oblique"
	 * - *color* text color
	 * - *bg-color* background color
	 * - *wrap* wrap text - "true" or "false" (default)
	 *
	 * @param theEntries data holding entries
	 * @param theRowParams parameters for entire row
	 */
	public function addTableRow($theEntries, $theRowParams = array()) {

		// use row params
		foreach ($theRowParams as $paramKey => $paramValue) {
			foreach ($theEntries as &$entryRow) {
				if (!array_key_exists($paramKey, $entryRow)) {
					$entryRow[$paramKey] = $paramValue;
				}
			}
		}

		// use column params
		$currentColumn = $this->tableParams['col_offset'];
		foreach ($theEntries as &$entryColumn) {
			if (array_key_exists($currentColumn, $this->tableParams['col_params'])) {
				foreach ($this->tableParams['col_params'][$currentColumn] as $paramKey => $paramValue) {
					if (!array_key_exists($paramKey, $entryColumn)) {
						$entryColumn[$paramKey] = $paramValue;
					}
				}
			}
			$currentColumn++;
		}

		// get current column
		$currentColumn = $this->tableParams['col_offset'];

		// print values
		foreach ($theEntries as $entry) {

			foreach ($entry as $entryKey => $entryValue) {

				switch ($entryKey) {

					case 'text':
						$this->xls->getActiveSheet()->setCellValueByColumnAndRow($currentColumn, $this->row, $entryValue);
						break;

					case 'font-name':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setName($entryValue);
						break;

					case 'font-size':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setSize($entryValue);
						break;

					case 'font-weight':
						switch ($entryValue) {
							case 'normal':
							case 'lighter':
								$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setBold(false);
								break;
							case 'bold':
							case 'bolder':
								$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setBold(true);
								break;
						}
						break;

					case 'font-style':
						switch ($entryValue) {
							case 'normal':
								$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setItalic(false);
								break;
							case 'italic':
							case 'oblique':
								$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->getFont()->setItalic(true);
								break;
						}
						break;

					case 'color':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->applyFromArray(array('font' => array('color' => array('rgb' => $entryValue))));
						break;

					case 'bg-color':
						$this->xls->getActiveSheet()->getStyleByColumnAndRow($currentColumn, $this->row)->applyFromArray(array('fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => $entryValue))));
						break;

					case 'wrap':
						if ($entryValue == true) {
							$this->xls->getActiveSheet()->getStyle(sprintf('%1$s%2$d:%1$s%2$d', PHPExcel_Cell::stringFromColumnIndex($currentColumn), $this->row))->getAlignment()->setWrapText(true);
						}
						break;

				}
			}

			$currentColumn++;
		}

		$this->row++;
		$this->tableParams['row_count']++;
		if (count($theEntries) > $this->tableParams['col_count']) {
			$this->tableParams['col_count'] = count($theEntries);
		}
	}

	/**
	 * Adds a table row with the gicen data as texts.
	 *
	 * No formatting can be given to a cell, only the text is inserted.
	 *
	 * @param theTexts texts for cells
	 */
	public function addTableTexts() {
		$data = array();
		foreach (func_get_args() as $text) {
			$data[] = array('text' => $text);
		}
		$this->addTableRow($data);
	}

	/**
	 * End table: sets params and styles that required data to be inserted.
	 */
	private function addTableFooter() {

		// auto width (for each column)
		foreach ($this->tableParams['auto_width'] as $col => $value) {
			$this->xls->getActiveSheet()->getColumnDimensionByColumn($col)->setAutoSize($value);
		}

		// filter (all columns)
		if ($this->tableParams['filter']) {
			$this->xls->getActiveSheet()->setAutoFilter(sprintf('%s%d:%s%d',
					PHPExcel_Cell::stringFromColumnIndex($this->tableParams['col_offset']),
					$this->tableParams['header_row'],
					PHPExcel_Cell::stringFromColumnIndex($this->tableParams['col_offset'] + $this->tableParams['col_count'] - 1),
					$this->tableParams['row_count']));
		}

	}

	/**
	 * Output file to browser
	 */
	public function output($filename = 'export.xlsx') {

		// set table footer
		$this->addTableFooter();

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
