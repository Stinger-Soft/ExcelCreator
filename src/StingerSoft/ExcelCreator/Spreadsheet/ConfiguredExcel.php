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

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use StingerSoft\ExcelCreator\ConfiguredExcelInterface;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\Helper;
use Symfony\Contracts\Translation\TranslatorInterface;

/**
 * Abstraction class to represent a single excel file
 */
class ConfiguredExcel implements ConfiguredExcelInterface {

	use Helper;

	/**
	 * The sheets of this excel file
	 *
	 * @var ArrayCollection|ConfiguredSheetInterface[]
	 */
	protected $sheets;

	/**
	 * The underlying excel file of PHPExcel
	 *
	 * @var Spreadsheet
	 */
	protected $phpExcel;

	/**
	 * Default constructor
	 * @param TranslatorInterface|null $translator
	 */
	public function __construct(TranslatorInterface $translator = null) {
		$this->phpExcel = new Spreadsheet();
		$this->sheets = new ArrayCollection();
		$this->translator = $translator;
	}

	/**
	 * @inheritDoc
	 * @throws Exception
	 */
	public function addSheet(string $title): ConfiguredSheetInterface {
		$excelSheet = null;
		if($this->sheets->isEmpty()) {
			$excelSheet = $this->phpExcel->getActiveSheet();
		} else {
			$excelSheet = $this->phpExcel->createSheet($this->sheets->count());
		}
		$this->setSheetTitle($excelSheet, $title);
		$sheet = new ConfiguredSheet($this, $excelSheet, $this->translator);
		$this->sheets->add($sheet);
		return $sheet;
	}

	/**
	 * @inheritDoc
	 */
	public function getSheets(): Collection {
		return $this->sheets;
	}

	/**
	 * @inheritDoc
	 */
	public function getTitle(): ?string {
		return $this->phpExcel->getProperties()->getTitle();
	}

	/**
	 * @inheritDoc
	 */
	public function setTitle(?string $title = null): ConfiguredExcelInterface {
		$this->phpExcel->getProperties()->setTitle($title);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCreator(): ?string {
		return $this->phpExcel->getProperties()->getCreator();
	}

	/**
	 * @inheritDoc
	 */
	public function setCreator(?string $creator = null): ConfiguredExcelInterface {
		$this->phpExcel->getProperties()->setCreator($creator);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCompany(): ?string {
		return $this->phpExcel->getProperties()->getCompany();
	}

	/**
	 * @inheritDoc
	 */
	public function setCompany(?string $company = null): ConfiguredExcelInterface {
		$this->phpExcel->getProperties()->setCompany($company);
		return $this;
	}

	/**
	 * Returns the underlying PHPExcel object
	 *
	 * @return Spreadsheet The underlying PHPExcel object
	 */
	public function getPhpExcel(): Spreadsheet {
		return $this->phpExcel;
	}

	/**
	 * @inheritDoc
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeToFile(string $filename): ConfiguredExcelInterface {
		$objPHPExcelWriter = IOFactory::createWriter($this->getPhpExcel(), 'Xlsx');
		$objPHPExcelWriter->save($filename);
		return $this;
	}

}