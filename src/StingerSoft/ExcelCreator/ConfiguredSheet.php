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
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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
	 * The default font family for all cells
	 *
	 * @var string
	 */
	protected $defaultFontFamily = 'Arial';

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
	protected $defaultHeaderFontColor = '000000';

	/**
	 * The default font size for the table headers
	 *
	 * @var integer
	 */
	protected $defaultHeaderFontSize = 8;

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
	 * The default font size for the data cells
	 *
	 * @var integer
	 */
	protected $defaultDataFontSize = 8;

	/**
	 * Property accessor to handle property path bindings
	 *
	 * @var PropertyAccessor
	 */
	protected $accessor;

	/**
	 * The data bound to this sheet
	 *
	 * @var array|\Traversable
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
	 * @var Worksheet
	 */
	protected $sheet = null;

	/**
	 * Creates some extra data for each data item object
	 *
	 * @var callable
	 */
	protected $extraData = null;

	protected $groupByBinding = null;

	/**
	 * Default constructor
	 *
	 * @param ConfiguredExcel $excel
	 *        	The parent excel file
	 * @param Worksheet $sheet
	 *        	The underlying PHP Excel worksheet
	 * @param TranslatorInterface $translator
	 *        	Translator to support translation of bindings
	 */
	public function __construct(ConfiguredExcel $excel, Worksheet $sheet, TranslatorInterface $translator = null) {
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
	 * Return the index or key for the given column binding.
	 *
	 * @param ColumnBinding $binding the column binding to get the index for
	 * @return bool|int|mixed|string the key for needle if it is found in the array, false otherwise.
	 *                               If needle is found in haystack more than once, the first matching key is returned. To return
	 *                               the keys for all matching values,  use array_keys with the optional search_value parameter instead.
	 */
	public function getIndexForBinding(ColumnBinding $binding) {
		return $this->bindings->indexOf($binding);
	}

	/**
	 * Sets an array of data to bind against this sheet
	 *
	 * @param array|\Traversable $data
	 */
	public function setData($data) {
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
	public function applyData($startColumn = 1, $headerRow = 1) {
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
	protected function renderHeaderRow($startColumn = 1, $headerRow = 1) {
		$this->sheet->getStyle(Coordinate::stringFromColumnIndex($startColumn) . $headerRow . ':' . Coordinate::stringFromColumnIndex($startColumn + $this->bindings->count() - 1) . $headerRow)->applyFromArray($this->getDefaultHeaderStyling());
		foreach($this->bindings as $binding) {
			$cell = $this->sheet->getCellByColumnAndRow($startColumn, $headerRow);
			$cell->setValue($this->decodeHtmlEntity($this->translate($binding->getLabel(), $binding->getLabelTranslationDomain())));
			if($binding->getOutline()) {
				$this->sheet->getColumnDimensionByColumn($startColumn)->setOutlineLevel($binding->getOutline());
			}
			if($binding->getHeaderFontColor() || $binding->getHeaderBackgroundColor()) {
				$cell->getStyle()->applyFromArray($this->getDefaultHeaderStyling($binding->getHeaderFontColor(), $binding->getHeaderBackgroundColor()));
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
	protected function renderDataRows($startColumn = 1, $headerRow = 1) {
		$row = $headerRow + 1;
		$lastGroupingValue = null;
		foreach($this->data as $item) {
			$this->sheet->getStyleByColumnAndRow($startColumn, $row, $startColumn + count($this->bindings) - 1, $row)->applyFromArray($this->getDefaultDataStyling($this->defaultDataFontColor, $this->defaultDataBackgroundColor));
			$extraData = array();
			if($this->extraData && is_callable($this->extraData)) {
				$extraData = call_user_func_array($this->extraData, array(
					$item 
				));
			}
			$column = $startColumn;
			foreach($this->bindings as $binding) {
				$cell = $this->sheet->getCellByColumnAndRow($column, $row);
				$value = $this->renderDataCell($cell, $item, $binding, $extraData);
				if($this->groupByBinding == $binding) {
					if($value == $lastGroupingValue) {
						$this->sheet->getRowDimension($cell->getRow())->setOutlineLevel(1);
					} else {
						$lastGroupingValue = $value;
					}
				}
				
				$column++;
			}
			$row++;
		}
	}

	/**
	 *
	 * Renders a single data cell
	 *
	 * @param Cell $cell
	 * @param object|array $item        	
	 * @param ColumnBinding $binding        	
	 */
	protected function renderDataCell(Cell $cell, $item, ColumnBinding $binding, array $extraData) {
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
		$url = $binding->getLinkUrl();
		if($url !== null) {
			if(is_callable($url)) {
				$url = call_user_func_array($url, array($binding, $item, $extraData));
			}
		}
		$cell->setValue($value);
		if($url != null) {
			$cell->getHyperlink()->setUrl($url);
		}
		return $value;
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
					return substr($path, 1);
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
		$this->sheet->setAutoFilter(Coordinate::stringFromColumnIndex($startColumn) . $headerRow . ':' . $lastColumn . $lastRow);
		
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
		$styling = array(
			'font' => array(
				'name' => $this->defaultFontFamily,
				'size' => $this->defaultDataFontSize 
			) 
		);
		if($fontColor) {
			$styling['font']['color'] = array(
				'rgb' => $fontColor 
			);
		}
		if($bgColor) {
			$styling['fill'] = array(
				'fillType' => Fill::FILL_SOLID,
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
	protected function getDefaultHeaderStyling($fontColor = null, $bgColor = null) {
		return array(
			'font' => array(
				'name' => $this->defaultFontFamily,
				'size' => $this->defaultHeaderFontSize,
				'color'=> array(
					'rgb' => $fontColor ?: $this->defaultHeaderFontColor,
				),
				'bold' => true 
			),
			'fill' => array(
				'fillType' => Fill::FILL_SOLID,
				'color' => array(
					'rgb' => $bgColor ?: $this->defaultHeaderBackgroundColor 
				) 
			),
			'alignment' => array(
				'wrap' => true,
				'horizontal' => Alignment::HORIZONTAL_CENTER,
				'vertical' => Alignment::VERTICAL_CENTER
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
	 * @return Worksheet
	 */
	public function getSheet() {
		return $this->sheet;
	}

	/**
	 * Gets the binding to group rows with the same value generate by the binding
	 *
	 * @return ColumnBinding|null The binding to group rows with the same value generate by the binding
	 */
	public function getGroupByBinding() {
		return $this->groupByBinding;
	}

	/**
	 * Sets the binding to group rows with the same value generate by the binding
	 *
	 * @param ColumnBinding|null $groupByBinding
	 *        	The binding to group rows with the same value generate by the binding
	 * @return \StingerSoft\ExcelCreator\ConfiguredSheet
	 */
	public function setGroupByBinding($groupByBinding) {
		$this->groupByBinding = $groupByBinding;
		return $this;
	}
}