<?php

/*
 * This file is part of the Stinger Excel Creator package.
 *
 * (c) Oliver Kotte <oliver.kotte@stinger-soft.net>
 * (c) Florian Meyer <florian.meyer@stinger-soft.net>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
namespace StingerSoft\ExcelCreator;

use Doctrine\Common\Collections\ArrayCollection;
use Symfony\Component\PropertyAccess\PropertyAccess;
use Symfony\Component\PropertyAccess\PropertyAccessor;
use Symfony\Component\Translation\TranslatorInterface;
use StingerSoft\PhpCommons\String\Utils;

/**
 * Abstraction layer to represent a single worksheet inside an excel file
 */
class ConfiguredSheet {
	
	use Helper;

	/**
	 * The configured bindings
	 *
	 * @var ColumnBinding[]|ArrayCollection
	 */
	protected $bindings = null;

	/**
	 * The default background color for the table headers
	 *
	 * @var string
	 */
	protected $defaultHeaderBackgroundColor = 'B8CCE4';

	/**
	 * The default font color for the table headers
	 *
	 * @var string
	 */
	protected $defaultHeaderFontColor = 'FFFFFF';

	/**
	 * The default font color for the data cells
	 *
	 * @var string|null
	 */
	protected $defaultDataFontColor = null;

	/**
	 * The default backgrund color for the data cells
	 *
	 * @var string|null
	 */
	protected $defaultDataBackgroundColor = null;

	/**
	 * Property accessor to handle property path bindings
	 *
	 * @var PropertyAccessor
	 */
	protected $accessor;

	/**
	 * The data bound to this sheet
	 *
	 * @var array
	 */
	protected $data;

	/**
	 * The parent excel file
	 *
	 * @var ConfiguredExcel
	 */
	protected $excel;

	/**
	 * The PHP worksheet attached to this object
	 *
	 * @var \PHPExcel_Worksheet
	 */
	protected $sheet = null;

	/**
	 * Creates some extra data for each data item object
	 *
	 * @var callable
	 */
	protected $extraData = null;

	/**
	 * Default constructor
	 *
	 * @param ConfiguredExcel $excel
	 *        	The parent excel file
	 * @param \PHPExcel_Worksheet $sheet
	 *        	The underlying PHP Excel worksheet
	 * @param TranslatorInterface $translator
	 *        	Translator to support translation of bindings
	 */
	public function __construct(ConfiguredExcel $excel, \PHPExcel_Worksheet $sheet, TranslatorInterface $translator = null) {
		$this->bindings = new ArrayCollection();
		$this->accessor = PropertyAccess::createPropertyAccessorBuilder()->disableExceptionOnInvalidIndex()->getPropertyAccessor();
		$this->excel = $excel;
		$this->sheet = $sheet;
		$this->translator = $translator;
	}

	/**
	 * Adds a column binding to this sheet
	 *
	 * @param ColumnBinding $binding        	
	 */
	public function addColumnBinding(ColumnBinding $binding) {
		$this->bindings->add($binding);
	}

	/**
	 * Sets an array of data to bind against this sheet
	 *
	 * @param array $data        	
	 */
	public function setData(array $data) {
		$this->data = $data;
	}

	/**
	 * Renders the given data on the sheet
	 *
	 * @param int $startColumn
	 *        	The column to start rendering
	 * @param int $headerRow
	 *        	The row to start rendering
	 */
	public function applyData($startColumn = 0, $headerRow = 1) {
		$this->renderHeaderRow($startColumn, $headerRow);
		$this->renderDataRows($startColumn, $headerRow);
		$this->applyTableStyling($startColumn, $headerRow);
	}

	/**
	 * Renders the table header
	 *
	 * @param int $startColumn
	 *        	The column to start rendering
	 * @param int $headerRow
	 *        	The row to start rendering
	 */
	protected function renderHeaderRow($startColumn = 0, $headerRow = 1) {
		$this->sheet->getStyle(\PHPExcel_Cell::stringFromColumnIndex($startColumn) . $headerRow . ':' . \PHPExcel_Cell::stringFromColumnIndex($startColumn + $this->bindings->count() - 1) . $headerRow)->applyFromArray($this->getDefaultHeaderStyling());
		foreach($this->bindings as $binding) {
			$this->sheet->setCellValueByColumnAndRow($startColumn, $headerRow, $this->decodeHtmlEntity($this->translate($binding->getLabel(), $binding->getLabelTranslationDomain())));
			if($binding->getOutline()) {
				$this->sheet->getColumnDimensionByColumn($startColumn)->setOutlineLevel($binding->getOutline());
			}
			$startColumn++;
		}
	}

	/**
	 * Renders the data rows
	 *
	 * @param int $startColumn
	 *        	The column to start rendering
	 * @param int $headerRow
	 *        	The row to start rendering
	 */
	protected function renderDataRows($startColumn = 0, $headerRow = 1) {
		$row = $headerRow + 1;
		foreach($this->data as $item) {
			$extraData = array();
			if($this->extraData && is_callable($this->extraData)) {
				$extraData = call_user_func_array($this->extraData, array(
					$item 
				));
			}
			$column = $startColumn;
			foreach($this->bindings as $binding) {
				$cell = $this->sheet->getCellByColumnAndRow($column, $row);
				$this->renderDataCell($cell, $item, $binding, $extraData);
				$column++;
			}
			$row++;
		}
	}

	/**
	 *
	 * Renders a single data cell
	 *
	 * @param \PHPExcel_Cell $cell        	
	 * @param object|array $item        	
	 * @param ColumnBinding $binding        	
	 */
	protected function renderDataCell(\PHPExcel_Cell &$cell, $item, ColumnBinding $binding, array $extraData) {
		$value = $this->getPropertyFromObject($item, $binding, $binding->getBinding(), '', $extraData);
		$styling = $this->getPropertyFromObject($item, $binding, $binding->getDataStyling(), null, $extraData);
		
		if(!$styling) {
			$fontColor = $this->getPropertyFromObject($item, $binding, $binding->getDataFontColor(), $this->defaultDataFontColor, $extraData);
			$bgColor = $this->getPropertyFromObject($item, $binding, $binding->getDataBackgroundColor(), $this->defaultDataBackgroundColor, $extraData);
			if($fontColor || $bgColor) {
				$cell->getStyle()->applyFromArray($this->getDefaultDataStyling($fontColor, $bgColor));
			}
		} else {
			$cell->getStyle()->applyFromArray($styling);
		}
		
		if($binding->getFormatter() && is_callable($binding->getFormatter())) {
			$value = call_user_func($binding->getFormatter(), $value);
		}
		
		if($binding->getDecodeHtml()) {
			$value = $this->decodeHtmlEntity($value);
		}
		$cell->setValue($value);
	}

	/**
	 *
	 * @param array|object $item
	 *        	The object to fetch the property from
	 * @param ColumnBinding $binding
	 *        	The binding configuration for this column
	 * @param string|callable $path
	 *        	The property or callable to fetch the desired property
	 * @param string $default
	 *        	The default value if no property was fetched
	 * @return mixed The value of the requested property
	 */
	protected function getPropertyFromObject($item, ColumnBinding $binding, $path, $default = '', array $extraData) {
		if(!$path === null)
			return $default;
		if(is_string($path)) {
			try {
				$obj = $item;
				if(Utils::startsWith($path, '$')) {
					return $path;
				}
				if(Utils::startsWith($path, '!')) {
					$pathSegs = explode('.', $path);
					$extraDataId = substr($pathSegs[0], 1);
					if(!isset($extraData[$extraDataId]))
						return $default;
					$obj = $extraData[$extraDataId];
					$path = implode('.', array_slice($pathSegs, 1));
				}
				return $this->accessor->getValue($obj, $path);
			} catch(\Symfony\Component\PropertyAccess\Exception\UnexpectedTypeException $ute) {
				return $default;
			}
		} else if(is_callable($path)) {
			return call_user_func_array($path, array(
				$binding,
				$item,
				$extraData 
			));
		} else if(is_array($path)) {
			return $path;
		}
		return $default;
	}

	/**
	 * Applies the default table styling to the sheet, i.e.
	 * filters, etc.
	 *
	 * @param int $startColumn
	 *        	The column to start rendering
	 * @param int $headerRow
	 *        	The row to start rendering
	 */
	protected function applyTableStyling($startColumn = 0, $headerRow = 1) {
		// Header filterable
		$lastColumn = $this->sheet->getHighestDataColumn();
		$lastRow = $this->sheet->getHighestDataRow();
		$this->sheet->setAutoFilter(\PHPExcel_Cell::stringFromColumnIndex($startColumn) . $headerRow . ':' . $lastColumn . $lastRow);
		
		//
		$this->sheet->setShowSummaryBelow(false);
		$this->sheet->setShowSummaryRight(false);
		$this->sheet->freezePaneByColumnAndRow($startColumn, $headerRow + 1);
		
		foreach($this->bindings as $columnIndex => $binding) {
			if($binding->getWrapText()) {
				$this->sheet->getStyleByColumnAndRow($columnIndex, $headerRow + 1, $columnIndex, $lastRow)->getAlignment()->setWrapText(true);
			}
			if($binding->getColumnWidth()) {
				if($binding->getColumnWidth() == 'auto') {
					$this->sheet->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
				} else {
					$this->sheet->getColumnDimensionByColumn($columnIndex)->setAutoSize(false)->setWidth($binding->getColumnWidth());
				}
			}
		}
	}

	/**
	 * Returns the default styling for data cells
	 *
	 * @param string $fontColor        	
	 * @param string $bgColor        	
	 * @return string[]
	 */
	protected function getDefaultDataStyling($fontColor, $bgColor) {
		$styling = array();
		if($fontColor) {
			$styling = array(
				'font' => array(
					'color' => array(
						'rgb' => $fontColor 
					) 
				) 
			);
		}
		if($bgColor) {
			$styling['fill'] = array(
				'type' => \PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array(
					'rgb' => $bgColor 
				) 
			);
		}
		return $styling;
	}

	/**
	 * Returns the default styling for headers cells
	 *
	 * @return string[]
	 */
	protected function getDefaultHeaderStyling() {
		return array(
			'font' => array(
				'size' => 10,
				'bold' => true 
			),
			'fill' => array(
				'type' => \PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array(
					'rgb' => $this->defaultHeaderBackgroundColor 
				) 
			),
			'alignment' => array(
				'wrap' => true,
				'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				'vertical' => \PHPExcel_Style_Alignment::VERTICAL_CENTER 
			) 
		);
	}

	/**
	 *
	 * @param callable $extraData        	
	 */
	public function setExtraData($extraData) {
		$this->extraData = $extraData;
		return $this;
	}

	/**
	 * Returns the underlying PHP Excel Sheet object
	 *
	 * @return \PHPExcel_Worksheet
	 */
	public function getSheet() {
		return $this->sheet;
	}
}