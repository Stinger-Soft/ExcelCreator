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

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Symfony\Component\Translation\TranslatorInterface;

/**
 * Abstraction layer to represent a single worksheet inside an excel file
 */
class ConfiguredSheet implements ConfiguredSheetInterface {

	protected $configuredSheet = null;

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
		$this->configuredSheet = new \StingerSoft\ExcelCreator\Spreadsheet\ConfiguredSheet($excel, $sheet, $translator);
	}

	/**
	 * Adds a column binding to this sheet
	 *
	 * @param ColumnBinding $binding
	 */
	public function addColumnBinding(ColumnBinding $binding) {
		$this->configuredSheet->addColumnBinding($binding);
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
		return $this->configuredSheet->getIndexForBinding($binding);
	}

	/**
	 * Sets an array of data to bind against this sheet
	 *
	 * @param array|\Traversable $data
	 */
	public function setData($data) {
		$this->configuredSheet->setData($data);
	}

	/**
	 * Renders the given data on the sheet
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 */
	public function applyData($startColumn = 1, $headerRow = 1) {
		$this->configuredSheet->applyData($startColumn, $headerRow);
	}

	/**
	 *
	 * @param callable $extraData
	 */
	public function setExtraData($extraData) {
		$this->configuredSheet->setExtraData($extraData);
		return $this;
	}

	/**
	 * Returns the underlying PHP Excel Sheet object
	 *
	 * @return Worksheet
	 */
	public function getSheet() {
		return $this->configuredSheet->getSheet();
	}

	/**
	 * Gets the binding to group rows with the same value generate by the binding
	 *
	 * @return ColumnBinding|null The binding to group rows with the same value generate by the binding
	 */
	public function getGroupByBinding() {
		return $this->configuredSheet->getGroupByBinding();
	}

	/**
	 * Sets the binding to group rows with the same value generate by the binding
	 *
	 * @param ColumnBinding|null $groupByBinding
	 *            The binding to group rows with the same value generate by the binding
	 * @return \StingerSoft\ExcelCreator\ConfiguredSheet
	 */
	public function setGroupByBinding($groupByBinding) {
		$this->configuredSheet->setGroupByBinding($groupByBinding);
		return $this;
	}
}