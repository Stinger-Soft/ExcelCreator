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

use Box\Spout\Common\Type;
use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\Common\Sheet;
use Box\Spout\Writer\Exception\SheetNotFoundException;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Doctrine\Common\Collections\ArrayCollection;
use StingerSoft\ExcelCreator\ConfiguredExcelInterface;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\Helper;
use Symfony\Component\Translation\TranslatorInterface;

class ConfiguredExcel implements ConfiguredExcelInterface {

	use Helper;

	/**
	 * The sheets of this excel file
	 *
	 * @var ArrayCollection|ConfiguredSheetInterface[]
	 */
	protected $sheets;

	/**
	 * The underyling excel file of PHPExcel
	 *
	 * @var AbstractMultiSheetsWriter
	 */
	protected $phpExcel;

	/**
	 * @var string
	 */
	protected $tempFile;

	/**
	 * Default contructor
	 * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
	 * @throws \Box\Spout\Common\Exception\IOException
	 */
	public function __construct(TranslatorInterface $translator = null) {
		$this->tempFile = self::createTemporaryFile('xlsx');
		$this->phpExcel = WriterFactory::create(Type::XLSX);
		$this->phpExcel->openToFile($this->tempFile);
		$this->sheets = new ArrayCollection();
		$this->translator = $translator;

	}

	/**
	 * @inheritDoc
	 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
	 * @throws \Box\Spout\Writer\Exception\InvalidSheetNameException
	 */
	public function addSheet($title) {
		$excelSheet = null;
		if($this->sheets->isEmpty()) {
			$excelSheet = $this->phpExcel->getCurrentSheet();
		} else {
			$excelSheet = $this->phpExcel->addNewSheetAndMakeItCurrent();
		}
		$excelSheet->setName($this->cleanSheetTitle($title));
		$sheet = new ConfiguredSheet($this, $excelSheet, $this->translator);
		$this->sheets->add($sheet);
		return $sheet;
	}

	/**
	 * @param Sheet $sheet
	 * @param array $rowData
	 * @throws WriterNotOpenedException
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\SpoutException
	 */
	public function addRow(Sheet $sheet, array $rowData) {
		try {
			if($this->phpExcel->getCurrentSheet() !== $sheet) {
				$this->phpExcel->setCurrentSheet($sheet);
			}
		} catch(SheetNotFoundException $e) {
		} catch(WriterNotOpenedException $e) {
		}
		$this->phpExcel->addRow($rowData);
	}

	/**
	 * @param Sheet $sheet
	 * @param array $rowData
	 * @param Style $style
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 * @throws \Box\Spout\Common\Exception\IOException
	 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
	 */
	public function addRowWithStyling(Sheet $sheet, array $rowData, Style $style) {
		if($this->phpExcel->getCurrentSheet() !== $sheet) {
			$this->phpExcel->setCurrentSheet($sheet);
		}
		$this->phpExcel->addRowWithStyle($rowData, $style);
	}

	/**
	 * @inheritDoc
	 */
	public function getSheets() {
		return $this->sheets;
	}

	/**
	 * @inheritDoc
	 */
	public function getTitle() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setTitle($title) {
	}

	/**
	 * @inheritDoc
	 */
	public function getCreator() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setCreator($creator) {
	}

	/**
	 * @inheritDoc
	 */
	public function getCompany() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setCompany($company) {

	}

	/**
	 * @inheritDoc
	 */
	public function writeToFile($filename) {
		$this->phpExcel->close();
		copy($this->tempFile, $filename);
		@unlink($this->tempFile);
	}

}