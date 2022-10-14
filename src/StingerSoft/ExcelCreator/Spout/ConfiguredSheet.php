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

namespace StingerSoft\ExcelCreator\Spout;

use Box\Spout\Common\Entity\Cell;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Common\Exception\InvalidArgumentException;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\SpoutException;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Entity\Sheet;
use Box\Spout\Writer\Exception\SheetNotFoundException;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
use DateTime;
use Doctrine\Common\Collections\ArrayCollection;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\DataType\DataTypes;
use StingerSoft\ExcelCreator\Helper;
use StingerSoft\PhpCommons\String\Utils;
use Symfony\Component\PropertyAccess\Exception\UnexpectedTypeException;
use Symfony\Component\PropertyAccess\PropertyAccess;
use Symfony\Component\PropertyAccess\PropertyAccessor;
use Symfony\Contracts\Translation\TranslatorInterface;
use Traversable;
use function is_array;
use function is_callable;

class ConfiguredSheet implements ConfiguredSheetInterface {

	use Helper;

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
	 * The configured bindings
	 *
	 * @var ColumnBinding[]|ArrayCollection
	 */
	protected $bindings;

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
	 * @var Sheet
	 */
	protected $sheet;

	/**
	 * Creates some extra data for each data item object
	 *
	 * @var callable
	 */
	protected $extraData;

	protected $currentRow = 1;

	/**
	 * Default constructor
	 *
	 * @param ConfiguredExcel $excel
	 *            The parent excel file
	 * @param Sheet $sheet
	 *            The underlying PHP Excel worksheet
	 * @param TranslatorInterface $translator
	 *            Translator to support translation of bindings
	 */
	public function __construct(ConfiguredExcel $excel, Sheet $sheet, TranslatorInterface $translator = null) {
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
	 */
	public function getSourceSheet() : Sheet {
		return $this->sheet;
	}

	/**
	 * @inheritDoc
	 * @throws IOException
	 * @throws InvalidArgumentException
	 * @throws SpoutException
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 */
	public function applyData(int $startColumn = 1, int $headerRow = 1): ConfiguredSheetInterface {
		$this->renderHeaderRow($startColumn, $headerRow);
		$this->renderDataRows($startColumn, $headerRow + 1);
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
	 * @inheritDoc
	 */
	public function getGroupByBinding(): ?ColumnBinding {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setGroupByBinding(?ColumnBinding $groupByBinding = null): ConfiguredSheetInterface {
		return $this;
	}

	/**
	 * Returns the default styling for headers cells
	 *
	 * @param string $fontColor
	 * @param string $bgColor
	 *
	 * @return Style
	 */
	protected function getDefaultHeaderStyling($fontColor = null, $bgColor = null): Style {
		return (new StyleBuilder())
			->setFontSize($this->defaultHeaderFontSize)
			->setFontColor($fontColor ?: $this->defaultHeaderFontColor)
			->setBackgroundColor($bgColor ?: $this->defaultHeaderBackgroundColor)
			->setFontName($this->defaultFontFamily)
			->setShouldWrapText()
			->setFontBold()
			->build();
	}

	/**
	 * Returns the default styling for data cells
	 *
	 * @param string $fontColor
	 * @param string $bgColor
	 * @return Style
	 */
	protected function getDefaultDataStyling($fontColor = null, $bgColor = null): Style {
		$builder = new StyleBuilder();
		$builder->setFontSize($this->defaultDataFontSize);
		if($fontColor || $this->defaultDataFontColor) {
			$builder->setFontColor($fontColor ?: $this->defaultDataFontColor);
		}
		if($bgColor || $this->defaultDataBackgroundColor) {
			$builder->setBackgroundColor($bgColor ?: $this->defaultDataBackgroundColor);
		}
		$builder->setFontName($this->defaultFontFamily);
		return $builder->build();
	}

	/**
	 * Renders the table header
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 *
	 * @throws IOException
	 * @throws InvalidArgumentException
	 * @throws SpoutException
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 */
	protected function renderHeaderRow($startColumn = 1, $headerRow = 1): void {

		for($i = $this->currentRow; $i < $headerRow; $i++) {
			$this->addRow([]);
		}
		$headerRowData = [];
		for($i = 1; $i < $startColumn; $i++) {
			$headerRowData[] = WriterEntityFactory::createCell('');
		}

		foreach($this->bindings as $binding) {
			$headerRowData[] = WriterEntityFactory::createCell($this->decodeHtmlEntity($this->translate($binding->getLabel(), $binding->getLabelTranslationDomain())));
		}
		$this->addRow($headerRowData, $this->getDefaultHeaderStyling());

	}

	/**
	 * @param Cell[] $data
	 * @param Style|null $styling
	 * @throws IOException
	 * @throws InvalidArgumentException
	 * @throws SpoutException
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 */
	protected function addRow($data, Style $styling = null): void {
		if($styling === null) {
			$this->excel->addRow($this->sheet, $data);
		} else {
			$this->excel->addRowWithStyling($this->sheet, $data, $styling);
		}
		$this->currentRow++;
	}

	/**
	 * Renders the data rows
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $startRow
	 *            The row to start rendering
	 * @throws IOException
	 * @throws InvalidArgumentException
	 * @throws SpoutException
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 */
	protected function renderDataRows($startColumn = 1, $startRow = 1): void {
		$style = $this->getDefaultDataStyling();
		for($i = $this->currentRow; $i < $startRow; $i++) {
			$this->excel->addRow($this->sheet, []);
		}

		foreach($this->data as $item) {
			$extraData = array();
			if($this->extraData && is_callable($this->extraData)) {
				$extraData = call_user_func_array($this->extraData, array(
					$item
				));
			}
			$rowData = [];
			for($i = 1; $i < $startColumn; $i++) {
				$rowData[] = WriterEntityFactory::createCell('');
			}
			foreach($this->bindings as $binding) {
				$value = $this->getPropertyFromObject($item, $binding, $binding->getBinding(), '', $extraData);
				if($binding->getFormatter() && is_callable($binding->getFormatter())) {
					$value = call_user_func($binding->getFormatter(), $value);
				}

				if($binding->getDecodeHtml()) {
					$value = $this->decodeHtmlEntity($value);
				}

				if($value instanceof DateTime) {
					$value= Date::PHPToExcel($value);
				}

				$cell = WriterEntityFactory::createCell($value, $style);

				if($binding->getForcedCellType() !== null) {
					$explicitType = DataTypes::getSpoutCellType($binding->getForcedCellType());
					$cell->setType($explicitType);
				}

				$styling = $this->getPropertyFromObject($item, $binding, $binding->getDataStyling(), null, $extraData);
				if(!$styling) {
					$fontColor = $this->getPropertyFromObject($item, $binding, $binding->getDataFontColor(), $this->defaultDataFontColor, $extraData);
					$bgColor = $this->getPropertyFromObject($item, $binding, $binding->getDataBackgroundColor(), $this->defaultDataBackgroundColor, $extraData);
					if($fontColor || $bgColor) {
						$cell->setStyle($this->getDefaultDataStyling($fontColor, $bgColor));
					}
				} else {
					if(!$styling instanceof Style) {
						throw new \InvalidArgumentException('For the spout implementation you have to pass a Style object');
					}
					$cell->setStyle($styling);
				}

				if($binding->getInternalCellModifier() !== null) {
					call_user_func_array($binding->getInternalCellModifier(), array(
						$binding, &$cell, $item, $extraData
					));
				}
				$rowData[] = $cell;
			}
			$this->addRow($rowData, $style);
		}
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
			} catch(UnexpectedTypeException $ute) {
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
}