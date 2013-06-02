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
	public function setDefaultFont(($theFontName = null, $theFontSize = null)) {
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
	 * as second parameter, it is used for each cell. Indidividual cell formats
	 * override second-parameter-formats.
	 *
	 * Possible format keys:
	 *
	 * - *label* entry text
	 * - *width* column width - "auto" or units
	 * - *filter* set filter to column? - "true" or "false" (default)
	 * - *wrap*	wrap text in column? - "true" or "false" (default)
	 * - *offset* column offset - numeric or text
	 * - *font* font name
	 * - *size* font size
	 * - *bold* bold text - "true" or "false" (default)
	 * - *italic* italic text - "true" or "false" (default)
	 *
	 * @param theEntries data holding entries
	 * @param theGlobalParams global parameters
	 */
	public function addTableHeader($data, $params = array()) {
		// offset
		$offset = 0;
		if (array_key_exists('offset', $params))
			$offset = is_numeric($params['offset']) ? (int)$params['offset'] : PHPExcel_Cell::columnIndexFromString($params['offset']);
		// font name
		if (array_key_exists('font', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setName($params['font']);
		// font size
		if (array_key_exists('size', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setSize($params['size']);
		// bold
		if (array_key_exists('bold', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setBold($params['bold']);
		// italic
		if (array_key_exists('italic', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setItalic($params['italic']);

		// set internal params that need to be processed after data are inserted
		$this->tableParams = array(
			'header_row' => $this->row,
			'offset' => $offset,
			'row_count' => 0,
			'auto_width' => array(),
			'filter' => array(),
			'wrap' => array()
		);

		foreach ($data as $d) {
			// set label
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset, $this->row, $d['label']);
			// set width
			if (array_key_exists('width', $d)) {
				if ($d['width'] == 'auto')
					$this->tableParams['auto_width'][] = $offset;
				else
					$this->xls->getActiveSheet()->getColumnDimensionByColumn($offset)->setWidth((float)$d['width']);
			}
			// filter
			if (array_key_exists('filter', $d) && $d['filter'])
				$this->tableParams['filter'][] = $offset;
			// wrap
			if (array_key_exists('wrap', $d) && $d['wrap'])
				$this->tableParams['wrap'][] = $offset;

			$offset++;
		}
		$this->row++;
	}

	/**
	 * Write array of data to actual row
	 */
	public function addTableRow($data) {
		$offset = $this->tableParams['offset'];

		foreach ($data as $d) {
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->row, $d);
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
	public function addData($data, $offset = 0) {
		// solve textual representation
		if (!is_numeric($offset))
			$offset = PHPExcel_Cell::columnIndexFromString($offset);

		foreach ($data as $d) {
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
