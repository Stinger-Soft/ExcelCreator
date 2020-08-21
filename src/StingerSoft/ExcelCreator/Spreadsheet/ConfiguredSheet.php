<?php
declare(strict_types=1);

/*
 * This file is part of the Stinger Excel Creator package.
 *
 * (c) Oliver Kotte <oliver.kotte@stinger-soft.net>
 * (c) Florian Meyer <florian.meyer@stinger-soft.net>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace StingerSoft\ExcelCreator\Spreadsheet;

use DateTime;
use Doctrine\Common\Collections\ArrayCollection;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\Helper;
use StingerSoft\PhpCommons\String\Utils;
use Symfony\Component\PropertyAccess\Exception\UnexpectedTypeException;
use Symfony\Component\PropertyAccess\PropertyAccess;
use Symfony\Component\PropertyAccess\PropertyAccessor;
use Symfony\Contracts\Translation\TranslatorInterface;
use Traversable;

/**
 * Abstraction layer to represent a single worksheet inside an excel file
 */
class ConfiguredSheet implements ConfiguredSheetInterface {

	use Helper;

	/**
	 * The configured bindings
	 *
	 * @var ColumnBinding[]|ArrayCollection
	 */
	protected $bindings;

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
	protected $defaultDataFontColor;

	/**
	 * The default backgrund color for the data cells
	 *
	 * @var string|null
	 */
	protected $defaultDataBackgroundColor;

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
	 * @var array|Traversable
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
	protected $sheet;

	/**
	 * Creates some extra data for each data item object
	 *
	 * @var callable
	 */
	protected $extraData;

	protected $groupByBinding;

	/**
	 * Default constructor
	 *
	 * @param ConfiguredExcel $excel
	 *            The parent excel file
	 * @param Worksheet $sheet
	 *            The underlying PHP Excel worksheet
	 * @param TranslatorInterface $translator
	 *            Translator to support translation of bindings
	 */
	public function __construct(ConfiguredExcel $excel, Worksheet $sheet, TranslatorInterface $translator = null) {
		$this->bindings = new ArrayCollection();
		$this->accessor = PropertyAccess::createPropertyAccessorBuilder()->disableExceptionOnInvalidIndex()->getPropertyAccessor();
		$this->excel = $excel;
		$this->sheet = $sheet;
		$this->translator = $translator;
	}

	/**
	 * @inheritDoc
	 */
	public function addColumnBinding(ColumnBinding $binding): ConfiguredSheetInterface {
		$this->bindings->add($binding);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getIndexForBinding(ColumnBinding $binding) {
		return $this->bindings->indexOf($binding);
	}

	/**
	 * @inheritDoc
	 */
	public function setData($data): ConfiguredSheetInterface {
		$this->data = $data;
		return $this;
	}

	/**
	 * @inheritDoc
	 * @throws Exception
	 */
	public function applyData(int $startColumn = 1, int $headerRow = 1): ConfiguredSheetInterface {
		$this->renderHeaderRow($startColumn, $headerRow);
		$this->renderDataRows($startColumn, $headerRow);
		$this->applyTableStyling($startColumn, $headerRow);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function setExtraData($extraData): ConfiguredSheetInterface {
		$this->extraData = $extraData;
		return $this;
	}

	/**
	 * Returns the underlying PHP Excel Sheet object
	 *
	 * @return Worksheet
	 */
	public function getSheet(): Worksheet {
		return $this->sheet;
	}

	/**
	 * @inheritDoc
	 */
	public function getGroupByBinding(): ColumnBinding {
		return $this->groupByBinding;
	}

	/**
	 * @inheritDoc
	 */
	public function setGroupByBinding(?ColumnBinding $groupByBinding = null): ConfiguredSheetInterface {
		$this->groupByBinding = $groupByBinding;
		return $this;
	}

	/**
	 * Renders the table header
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 * @throws Exception
	 */
	protected function renderHeaderRow($startColumn = 1, $headerRow = 1): void {
		$this->sheet->getStyle(Coordinate::stringFromColumnIndex($startColumn) . $headerRow . ':' . Coordinate::stringFromColumnIndex($startColumn + $this->bindings->count() - 1) . $headerRow)->applyFromArray($this->getDefaultHeaderStyling());
		foreach($this->bindings as $binding) {
			$cell = $this->sheet->getCellByColumnAndRow($startColumn, $headerRow);
			if($cell !== null) {
				$cell->setValue($this->decodeHtmlEntity($this->translate($binding->getLabel(), $binding->getLabelTranslationDomain())));
				if($binding->getOutline()) {
					$this->sheet->getColumnDimensionByColumn($startColumn)->setOutlineLevel($binding->getOutline());
				}
				if($binding->getHeaderFontColor() || $binding->getHeaderBackgroundColor()) {
					$cell->getStyle()->applyFromArray($this->getDefaultHeaderStyling($binding->getHeaderFontColor(), $binding->getHeaderBackgroundColor()));
				}
			}
			$startColumn++;
		}
	}

	/**
	 * Renders the data rows
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 * @throws Exception
	 */
	protected function renderDataRows($startColumn = 1, $headerRow = 1): void {
		$row = $headerRow + 1;
		$lastGroupingValue = null;
		foreach($this->data as $item) {
			$this->sheet->getStyleByColumnAndRow($startColumn, $row, $startColumn + count($this->bindings) - 1, $row)->applyFromArray($this->getDefaultDataStyling($this->defaultDataFontColor, $this->defaultDataBackgroundColor));
			$extraData = array();
			if($this->extraData && is_callable($this->extraData)) {
				$extraData = call_user_func($this->extraData, $item);
			}
			$column = $startColumn;
			foreach($this->bindings as $binding) {
				$cell = $this->sheet->getCellByColumnAndRow($column, $row);
				if($cell !== null) {
					$value = $this->renderDataCell($cell, $item, $binding, $extraData);
					if($this->groupByBinding === $binding) {
						if($value === $lastGroupingValue) {
							$this->sheet->getRowDimension($cell->getRow())->setOutlineLevel(1);
						} else {
							$lastGroupingValue = $value;
						}
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
	 * @param array $extraData
	 * @return mixed
	 * @throws Exception
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
		if(($url !== null) && is_callable($url)) {
			$url = call_user_func($url, $binding, $item, $extraData);
		}
		if($value instanceof DateTime) {
			$value = Date::PHPToExcel($value);
		}
		if($binding->getForcedCellType() === null) {
			$cell->setValue($value);
		} else {
			$cell->setValueExplicit($value, $binding->getForcedCellType());
		}
		if($url !== null) {
			$cell->getHyperlink()->setUrl($url);
		}
		if($binding->getInternalCellModifier() !== null) {
			call_user_func_array($binding->getInternalCellModifier(), array($binding, &$cell, $item, $extraData));
		}
		return $value;
	}

	/**
	 *
	 * @param array|object $item
	 *            The object to fetch the property from
	 * @param ColumnBinding $binding
	 *            The binding configuration for this column
	 * @param string|callable $path
	 *            The property or callable to fetch the desired property
	 * @param string $default
	 *            The default value if no property was fetched
	 * @param array $extraData
	 * @return mixed The value of the requested property
	 */
	protected function getPropertyFromObject($item, ColumnBinding $binding, $path, $default = '', array $extraData = array()) {
		// Before $path was a !, nobody knows why...
		if($path === null) {
			return $default;
		}
		if(is_string($path)) {
			try {
				$obj = $item;
				if(Utils::startsWith($path, '$')) {
					return substr($path, 1);
				}
				if(Utils::startsWith($path, '!')) {
					$pathSegs = explode('.', $path);
					$extraDataId = substr($pathSegs[0], 1);
					if(!isset($extraData[$extraDataId])) {
						return $default;
					}
					$obj = $extraData[$extraDataId];
					$path = implode('.', array_slice($pathSegs, 1));
				}
				return $this->accessor->getValue($obj, $path);
			} catch(UnexpectedTypeException $ute) {
				return $default;
			}
		} else if(is_callable($path)) {
			return call_user_func($path, $binding, $item, $extraData);
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
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 * @throws Exception
	 */
	protected function applyTableStyling($startColumn = 0, $headerRow = 1): void {
		// Header filterable
		$lastColumn = $this->sheet->getHighestDataColumn();
		$lastRow = $this->sheet->getHighestDataRow();
		$this->sheet->setAutoFilter(Coordinate::stringFromColumnIndex($startColumn) . $headerRow . ':' . $lastColumn . $lastRow);

		//
		$this->sheet->setShowSummaryBelow(false);
		$this->sheet->setShowSummaryRight(false);
		$this->sheet->freezePaneByColumnAndRow($startColumn, $headerRow + 1);

		foreach($this->bindings as $columnIndex => $binding) {
			//PHPExcel -> PHPSpreadsheet = +1
			$columnIndex++;
			if($binding->getWrapText()) {
				$this->sheet->getStyleByColumnAndRow($columnIndex, $headerRow + 1, $columnIndex, $lastRow)->getAlignment()->setWrapText(true);
			}
			if($binding->getColumnWidth()) {
				if($binding->getColumnWidth() === 'auto') {
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
	protected function getDefaultDataStyling($fontColor, $bgColor): array {
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
				'color'    => array(
					'rgb' => $bgColor
				)
			);
		}
		return $styling;
	}

	/**
	 * Returns the default styling for headers cells
	 *
	 * @param string|null $fontColor
	 * @param string|null $bgColor
	 * @return string[]
	 */
	protected function getDefaultHeaderStyling(?string $fontColor = null, ?string $bgColor = null): array {
		return array(
			'font'      => array(
				'name'  => $this->defaultFontFamily,
				'size'  => $this->defaultHeaderFontSize,
				'color' => array(
					'rgb' => $fontColor ?: $this->defaultHeaderFontColor,
				),
				'bold'  => true
			),
			'fill'      => array(
				'fillType' => Fill::FILL_SOLID,
				'color'    => array(
					'rgb' => $bgColor ?: $this->defaultHeaderBackgroundColor
				)
			),
			'alignment' => array(
				'wrap'       => true,
				'horizontal' => Alignment::HORIZONTAL_CENTER,
				'vertical'   => Alignment::VERTICAL_CENTER
			)
		);
	}
}
