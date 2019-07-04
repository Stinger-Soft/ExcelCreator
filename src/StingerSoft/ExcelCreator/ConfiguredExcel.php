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

use Doctrine\Common\Collections\Collection;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Symfony\Contracts\Translation\TranslatorInterface;

/**
 * Abstraction class to represent a single excel file
 */
class ConfiguredExcel implements ConfiguredExcelInterface {

	protected $excel;

	/**
	 * Default constructor
	 * @param TranslatorInterface|null $translator
	 */
	public function __construct(TranslatorInterface $translator = null) {
		$this->excel = new \StingerSoft\ExcelCreator\Spreadsheet\ConfiguredExcel(null, $translator);
	}

	/**
	 * @inheritDoc
	 * @throws Exception
	 */
	public function addSheet(string $title): ConfiguredSheetInterface {
		return $this->excel->addSheet($title);
	}

	/**
	 * @inheritDoc
	 */
	public function getSheets(): Collection {
		return $this->excel->getSheets();
	}

	/**
	 * @inheritDoc
	 */
	public function getTitle(): string {
		return $this->excel->getTitle();
	}

	/**
	 * @inheritDoc
	 */
	public function setTitle(?string $title = null): ConfiguredExcelInterface {
		$this->excel->setTitle($title);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCreator(): ?string {
		return $this->excel->getCreator();
	}

	/**
	 * @inheritDoc
	 */
	public function setCreator(?string $creator = null): ConfiguredExcelInterface {
		$this->excel->setCreator($creator);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCompany(): ?string {
		return $this->excel->getCompany();
	}

	/**
	 * @inheritDoc
	 */
	public function setCompany(?string $company = null): ConfiguredExcelInterface {
		return $this->excel->setCompany($company);
	}

	/**
	 * Returns the underyling PHPExcel object
	 *
	 * @return Spreadsheet The underyling PHPExcel object
	 */
	public function getPhpExcel(): Spreadsheet {
		return $this->excel->getPhpExcel();
	}

	/**
	 * @inheritDoc
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeToFile(string $filename): ConfiguredExcelInterface {
		$this->excel->writeToFile($filename);
		return $this;
	}
}