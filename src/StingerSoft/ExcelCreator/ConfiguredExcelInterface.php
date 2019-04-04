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

interface ConfiguredExcelInterface {

	/**
	 * Adds and returns a new sheet
	 *
	 * @param string $title
	 *            The title of the new sheet
	 * @return ConfiguredSheetInterface
	 */
	public function addSheet($title);

	/**
	 * Returns the worksheets of this excel file
	 *
	 * @return ConfiguredSheetInterface[]
	 */
	public function getSheets();

	/**
	 * Returns the title of this excel file
	 *
	 * @return string The title of this excel file
	 */
	public function getTitle();

	/**
	 * Sets the title of this excel file
	 *
	 * @param string $title
	 *            The title of this excel file
	 */
	public function setTitle($title);

	/**
	 * Get the author of this excel file
	 *
	 * @return string The author of this excel file
	 */
	public function getCreator();

	/**
	 * Sets the author of this excel file
	 *
	 * @param string $creator
	 *            The author of this excel file
	 */
	public function setCreator($creator);

	/**
	 *
	 * @return string The company this excel file is created by
	 */
	public function getCompany();

	/**
	 *
	 * @param string $company
	 *            The company this excel file is created by
	 */
	public function setCompany($company);

	/**
	 * @return void
	 */
	public function writeToFile($filename);


}