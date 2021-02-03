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

namespace StingerSoft\ExcelCreator;

interface ConfiguredSheetInterface {
	/**
	 * Adds a column binding to this sheet
	 *
	 * @param ColumnBinding $binding
	 * @return ConfiguredSheetInterface
	 */
	public function addColumnBinding(ColumnBinding $binding): ConfiguredSheetInterface;

	/**
	 * Return the index or key for the given column binding.
	 *
	 * @param ColumnBinding $binding the column binding to get the index for
	 * @return bool|int|mixed|string the key for needle if it is found in the array, false otherwise.
	 *                               If needle is found in haystack more than once, the first matching key is returned. To return
	 *                               the keys for all matching values,  use array_keys with the optional search_value parameter instead.
	 */
	public function getIndexForBinding(ColumnBinding $binding);

	/**
	 * Sets an array of data to bind against this sheet
	 *
	 * @param array|\Traversable $data
	 * @return ConfiguredSheetInterface
	 */
	public function setData($data): ConfiguredSheetInterface;

	/**
	 * Renders the given data on the sheet
	 *
	 * @param int $startColumn
	 *            The column to start rendering
	 * @param int $headerRow
	 *            The row to start rendering
	 * @return ConfiguredSheetInterface
	 */
	public function applyData(int $startColumn = 1, int $headerRow = 1): ConfiguredSheetInterface;

	/**
	 *
	 * @param callable $extraData
	 * @return ConfiguredSheetInterface
	 */
	public function setExtraData($extraData): ConfiguredSheetInterface;

	/**
	 * Gets the binding to group rows with the same value generate by the binding
	 *
	 * @return ColumnBinding|null The binding to group rows with the same value generate by the binding
	 */
	public function getGroupByBinding(): ?ColumnBinding;

	/**
	 * Sets the binding to group rows with the same value generate by the binding
	 *
	 * @param ColumnBinding|null $groupByBinding
	 *            The binding to group rows with the same value generate by the binding
	 * @return self
	 */
	public function setGroupByBinding(?ColumnBinding $groupByBinding = null): ConfiguredSheetInterface;

	/**
	 * @return mixed
	 */
	public function getSourceSheet();
}