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

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Writer\Common\Entity\Sheet;
use OpenSpout\Writer\Exception\InvalidSheetNameException;
use OpenSpout\Writer\Exception\SheetNotFoundException;
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\WriterInterface;
use OpenSpout\Writer\XLSX\Writer;
use OpenSpout\Writer\XLSX\Writer as XLSXWriter;
use StingerSoft\ExcelCreator\ConfiguredExcelInterface;
use StingerSoft\ExcelCreator\ConfiguredSheetInterface;
use StingerSoft\ExcelCreator\Helper;
use Symfony\Contracts\Translation\TranslatorInterface;

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
	 * @var WriterInterface
	 */
	protected $phpExcel;

	/**
	 * @var string
	 */
	protected $tempFile;

	/**
	 * Default constructor
	 *
	 * @param TranslatorInterface|null $translator
	 * @throws IOException
	 */
	public function __construct(TranslatorInterface $translator = null) {
		$this->tempFile = self::createTemporaryFile('xlsx');
		$this->phpExcel = new XLSXWriter();
		$this->phpExcel->openToFile($this->tempFile);
		$this->sheets = new ArrayCollection();
		$this->translator = $translator;

	}

	/**
	 * @inheritDoc
	 * @throws WriterNotOpenedException
	 * @throws InvalidSheetNameException
	 */
	public function addSheet(string $title): ConfiguredSheetInterface {
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

	public function setActiveSheet(ConfiguredSheetInterface $sheet): void {
		$sourceSheet = $sheet->getSourceSheet();
		if($sourceSheet instanceof \OpenSpout\Writer\Common\Entity\Sheet) {
			$this->phpExcel->setCurrentSheet($sourceSheet);
		}
	}

	/**
	 * @param Sheet $sheet
	 * @param array $rowData
	 * @throws WriterNotOpenedException
	 * @throws IOException
	 * @throws SpoutException
	 */
	public function addRow(Sheet $sheet, array $rowData): void {
		try {
			if($this->phpExcel->getCurrentSheet() !== $sheet) {
				$this->phpExcel->setCurrentSheet($sheet);
			}
		} catch(SheetNotFoundException $e) {
		} catch(WriterNotOpenedException $e) {
		}
		$this->phpExcel->addRow(new Row($rowData));
	}

	/**
	 * @param Sheet $sheet
	 * @param array $rowData
	 * @param Style $style
	 * @throws SheetNotFoundException
	 * @throws WriterNotOpenedException
	 * @throws IOException
	 * @throws InvalidArgumentException
	 */
	public function addRowWithStyling(Sheet $sheet, array $rowData, Style $style): void {
		if($this->phpExcel->getCurrentSheet() !== $sheet) {
			$this->phpExcel->setCurrentSheet($sheet);
		}
		$this->phpExcel->addRow(new Row($rowData, $style));
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
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setTitle(?string $title = null): ConfiguredExcelInterface {
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCreator(): ?string {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setCreator(?string $creator = null): ConfiguredExcelInterface {
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function getCompany(): ?string {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function setCompany(?string $company = null): ConfiguredExcelInterface {
		return $this;
	}

	/**
	 * Returns the underlying Writer object
	 *
	 * @return Writer The underlying Writer object
	 */
	public function getPhpExcel(): Writer {
		return $this->phpExcel;
	}

	/**
	 * @inheritDoc
	 */
	public function writeToFile(string $filename): ConfiguredExcelInterface {
		$this->phpExcel->close();
		copy($this->tempFile, $filename);
		@unlink($this->tempFile);
		return $this;
	}

}