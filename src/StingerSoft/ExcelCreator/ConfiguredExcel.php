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

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Symfony\Component\Translation\TranslatorInterface;

/**
 * Abstraction class to represent a single excel file
 */
class ConfiguredExcel implements ConfiguredExcelInterface {

	protected $excel;

	/**
	 * Default contructor
	 */
	public function __construct(TranslatorInterface $translator = null) {
		$this->excel = new \StingerSoft\ExcelCreator\Spreadsheet\ConfiguredExcel($translator);
	}

	/**
	 * Adds and returns a new sheet
	 *
	 * @param string $title
	 *            The title of the new sheet
	 * @return ConfiguredSheetInterface
	 */
	public function addSheet($title) {
		return $this->excel->addSheet($title);
	}

	/**
	 * Returns the worksheets of this excel file
	 *
	 * @return ConfiguredSheetInterface[]
	 */
	public function getSheets() {
		return $this->excel->getSheets();
	}

	/**
	 * Returns the title of this excel file
	 *
	 * @return string The title of this excel file
	 */
	public function getTitle() {
		return $this->excel->getTitle();
	}

	/**
	 * Sets the title of this excel file
	 *
	 * @param string $title
	 *            The title of this excel file
	 */
	public function setTitle($title) {
		$this->excel->setTitle($title);
	}

	/**
	 * Get the author of this excel file
	 *
	 * @return string The author of this excel file
	 */
	public function getCreator() {
		return $this->excel->getCreator();
	}

	/**
	 * Sets the author of this excel file
	 *
	 * @param string $creator
	 *            The author of this excel file
	 */
	public function setCreator($creator) {
		$this->excel->setCreator($creator);
	}

	/**
	 *
	 * @return string The company this excel file is created by
	 */
	public function getCompany() {
		return $this->excel->getCompany();
	}

	/**
	 *
	 * @param string $company
	 *            The company this excel file is created by
	 */
	public function setCompany($company) {
		return $this->excel->setCompany($company);
	}

	/**
	 * Returns the underyling PHPExcel object
	 *
	 * @return Spreadsheet The underyling PHPExcel object
	 */
	public function getPhpExcel() {
		return $this->excel->getPhpExcel();
	}

	/**
	 * @param $filename
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeToFile($filename) {
		$this->excel->writeToFile($filename);
	}
}