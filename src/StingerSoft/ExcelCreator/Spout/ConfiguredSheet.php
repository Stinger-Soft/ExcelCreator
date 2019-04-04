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

namespace StingerSoft\ExcelCreator\Spout;

use Box\Spout\Writer\Common\Sheet;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\Style\StyleBuilder;
use Doctrine\Common\Collections\ArrayCollection;
use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\Helper;
use StingerSoft\PhpCommons\String\Utils;
use Symfony\Component\PropertyAccess\PropertyAccess;
use Symfony\Component\PropertyAccess\PropertyAccessor;
use Symfony\Component\Translation\TranslatorInterface;

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
	 * The configured bindings
	 *
	 * @var ColumnBinding[]|ArrayCollection
	 */
	protected $bindings = null;

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
	 * @var Sheet
	 */
	protected $sheet = null;

	/**
	 * Creates some extra data for each data item object
	 *
	 * @var callable
	 */
	protected $extraData = null;

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
	public function addColumnBinding(ColumnBinding $binding) {
		$this->bindings->add($binding);
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
	public function setData($data) {
		$this->data = $data;
	}

	/**
	 * @inheritDoc
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
	 * @throws \Box\Spout\Common\Exception\SpoutException
	 * @throws \Box\Spout\Writer\Exception\SheetNotFoundException
	 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
	 */
	public function applyData($startColumn = 1, $headerRow = 1) {
		$this->renderHeaderRow($startColumn, $headerRow);
		$this->renderDataRows($startColumn, $headerRow + 1);
	}

	/**
	 * @inheritDoc
	 */
	public function setExtraData($extraData) {
		$this->extraData = $extraData;
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getGroupByBinding() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setGroupByBinding($groupByBinding) {
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
	protected function getDefaultHeaderStyling($fontColor = null, $bgColor = null) {
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
	protected function getDefaultDataStyling($fontColor = null, $bgColor = null) {
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
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
	 * @throws \Box\Spout\Common\Exception\SpoutException
	 * @throws \Box\Spout\Writer\Exception\SheetNotFoundException
	 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
	 */
	protected function renderHeaderRow($startColumn = 1, $headerRow = 1) {

		for($i = $this->currentRow; $i < $headerRow; $i++) {
			$this->addRow([]);
		}
		$headerRowData = [];
		for($i = 1; $i < $startColumn; $i++) {
			$headerRowData[] = '';
		}

		foreach($this->bindings as $binding) {
			$headerRowData[] = $this->decodeHtmlEntity($this->translate($binding->getLabel(), $binding->getLabelTranslationDomain()));
		}
		$this->addRow($headerRowData, $this->getDefaultHeaderStyling());

	}

	/**
	 * @param $data
	 * @param Style|null $styling
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
	 * @throws \Box\Spout\Common\Exception\SpoutException
	 * @throws \Box\Spout\Writer\Exception\SheetNotFoundException
	 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
	 */
	protected function addRow($data, Style $styling = null) {
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
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
	 * @throws \Box\Spout\Common\Exception\SpoutException
	 * @throws \Box\Spout\Writer\Exception\SheetNotFoundException
	 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
	 */
	protected function renderDataRows($startColumn = 1, $startRow = 1) {
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
				$rowData[] = '';
			}
			foreach($this->bindings as $binding) {
				$value = $this->getPropertyFromObject($item, $binding, $binding->getBinding(), '', $extraData);
				if($binding->getFormatter() && is_callable($binding->getFormatter())) {
					$value = call_user_func($binding->getFormatter(), $value);
				}

				if($binding->getDecodeHtml()) {
					$value = $this->decodeHtmlEntity($value);
				}
				if($value instanceof \DateTime) {
					$value = $value->format('Y-m-d H:i:s');
				}
				$rowData[] = $value;
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
			} catch(\Symfony\Component\PropertyAccess\Exception\UnexpectedTypeException $ute) {
				return $default;
			}
		} else if(\is_callable($path)) {
			return call_user_func_array($path, array(
				$binding,
				$item,
				$extraData
			));
		} else if(\is_array($path)) {
			return $path;
		}
		return $default;
	}
}