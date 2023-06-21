e<?php
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

namespace StingerSoft\ExcelCreator\Utils;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

abstract class AbstractExcelImporter {

	abstract protected function getHeaderRow(): int;

	abstract protected function getFirstDataColumn(): int;

	abstract protected function getFirstDataRow(): int;

	abstract protected function getHeaders(): array;

	protected function onRow(Worksheet $sheet, int $row, array $headerMapping): void {
		throw new \LogicException('You must override the onRow() method in the concrete service class to use the iterateSheetRows function.');
	}

	protected function iterateSheetRows(Worksheet $sheet): void {
		$headerMapping = $this->getHeaderMapping($sheet);
		$this->checkMapping($headerMapping);
		$highestRow = $this->getHighestDataRow($sheet);
		for($row = $this->getFirstDataRow(); $row <= $highestRow; $row++) {
			$this->onRow($sheet, $row, $headerMapping);
		}
	}

	protected function getHighestDataRow(Worksheet $sheet): int {
		return $sheet->getHighestDataRow(Coordinate::stringFromColumnIndex($this->getFirstDataColumn()));
	}

	/**
	 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
	 */
	protected function openExcelFile(string $fileName): Spreadsheet {
		if(!file_exists($fileName)) {
			throw new \RuntimeException('Cannot find excel file ' . $fileName);
		}
		$reader = IOFactory::createReaderForFile($fileName);
		$reader->setReadDataOnly(true);
		return $reader->load($fileName);
	}

	protected function findColumnByName(Worksheet $sheet, string $value, int $headerRow = null): ?int {
		if($headerRow === null) {
			$headerRow = $this->getHeaderRow();
		}
		$highestColumn = Coordinate::columnIndexFromString($sheet->getHighestDataColumn($headerRow));
		for($column = $this->getFirstDataColumn(); $column <= $highestColumn; $column++) {
			$cell = $sheet->getCellByColumnAndRow($column, $headerRow);
			if($cell && $cell->getValue() === $value) {
				return $column;
			}
		}
		return null;
	}


	protected function getHeaderMapping(Worksheet $sheet): array {
		$result = array_combine($this->getHeaders(), $this->getHeaders());
		array_walk($result, function(&$item, $key) use ($sheet) {
			$item = $this->findColumnByName($sheet, $key);
		});
		return $result;
	}

	protected function checkMapping(array $mapping): void {
		if(in_array(null, $mapping, true)) {
			throw new \RuntimeException(
				'Not all necessary columns found in row ' . $this->getHeaderRow() . "!\r\n".
				"Required columns \r\n" .
				"================= \r\n" .
				implode(', ', $this->getHeaders())."\r\n\r\n".
				" Missing columns\r\n" .
				"================= \r\n" .
				implode(', ', array_keys(array_filter($mapping, function($value){
					return $value === null;
				})))
			);
		}
	}

	protected function getStringValue(Worksheet $sheet, int $column, int $row): ?string {
		$resStr = $sheet->getCellByColumnAndRow($column, $row)->getOldCalculatedValue();
		$resStr = $resStr ?? $sheet->getCellByColumnAndRow($column, $row)->getValue();
		if($resStr === null || $resStr === '-' || (is_string($resStr) && trim($resStr) === '')) {
			return null;
		}
		return (string)$resStr;
	}

	protected function getDateValue(Worksheet $sheet, int $column, int $row): ?\DateTimeInterface {
		$resStr = $sheet->getCellByColumnAndRow($column, $row)->getOldCalculatedValue();
		$resStr = $resStr ?? $sheet->getCellByColumnAndRow($column, $row)->getValue();
		if($resStr === null || $resStr === '-' || (is_string($resStr) && trim($resStr) === '')) {
			return null;
		}
		return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($resStr);
	}

	protected function getNumericValue(Worksheet $sheet, int $column, int $row): ?float {
		$resStr = $sheet->getCellByColumnAndRow($column, $row)->getValue();
		if(is_numeric($resStr)) {
			return (float)$resStr;
		}
		return null;
	}

	protected function getBoolValue(Worksheet $sheet, int $column, int $row): bool {
		$resStr = $sheet->getCellByColumnAndRow($column, $row)->getValue();
		if($resStr === null || $resStr === '-' || $resStr === '0' || $resStr === 0 ||  (is_string($resStr) && trim($resStr) === '')) {
			return false;
		}
		return true;
	}
}
