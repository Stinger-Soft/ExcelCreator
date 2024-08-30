<?php

namespace StingerSoft\ExcelCreator\Utils;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Reader\XLSX\Reader;

abstract class AbstractSpoutExcelImporter {
	abstract protected function getHeaderRow(): int;

	abstract protected function getFirstDataColumn(): int;

	abstract protected function getFirstDataRow(): int;

	abstract protected function getHeaders(): array;

	protected function onRow(SheetInterface $sheet, int $row, array $headerMapping): void {
		throw new \LogicException('You must override the onRow() method in the concrete service class to use the iterateSheetRows function.');
	}

	protected function iterateSheetRows(SheetInterface $sheet): void {
		$headerMapping = $this->getHeaderMapping($sheet);
		$this->checkMapping($headerMapping);
		$highestRow = $this->getHighestDataRow($sheet);
		foreach($sheet->getRowIterator() as $row) {
			$this->onRow($sheet, $row, $headerMapping);
		}
	}

	protected function getHighestDataRow(SheetInterface $sheet): int {
		$sheet->getRowIterator()->rewind();
		$rowId = 0;
		foreach ($sheet->getRowIterator() as $row) {
			$rowId++;
		}
		$sheet->getRowIterator()->rewind();
		return $rowId;
	}

	/**
	 */
	protected function openExcelFile(string $fileName): ReaderInterface {
		if(!file_exists($fileName)) {
			throw new \RuntimeException('Cannot find excel file ' . $fileName);
		}
		$reader = new Reader();
		$reader->open($fileName);
		return $reader;
	}

	protected function findColumnByName(SheetInterface $sheet, string $value, int $headerRow = null): ?int {
		if($headerRow === null) {
			$headerRow = $this->getHeaderRow();
		}
		$row = $this->getRowByNum($sheet, $headerRow);

		$highestColumn = $row->getNumCells();
		for($column = $this->getFirstDataColumn(); $column <= $highestColumn; $column++) {
			$cell = $row->getCellAtIndex($column);
			if($cell && $cell->getValue() === $value) {
				return $column;
			}
		}
		return null;
	}

	protected function getRowByNum(SheetInterface $sheet, int $row): ?Row {
		$sheet->getRowIterator()->rewind();
		for($rowId = 0; $rowId <= $row; $rowId++) {
			$sheet->getRowIterator()->next();
		}
		$result = $sheet->getRowIterator()->current();
		$sheet->getRowIterator()->rewind();
		return $result;
	}


	protected function getHeaderMapping(SheetInterface $sheet): array {
		$result = array_combine($this->getHeaders(), $this->getHeaders());
		array_walk($result, function(&$item, $key) use ($sheet) {
			$item = $this->findColumnByName($sheet, $key);
		});
		return $result;
	}

	protected function checkMapping(array $mapping): void {
		if(in_array(null, $mapping, true)) {
			throw new \RuntimeException('Not all necessary columns found in row ' . $this->getHeaderRow() . '! Required columns are: ' . implode(', ', $this->getHeaders()));
		}
	}

	protected function getStringValue(SheetInterface $sheet, int $column, Row $row): ?string {
		$cell = $row->getCellAtIndex($column);
		$resStr = $cell->getValue();
		if($resStr === null || $resStr === '-' || (is_string($resStr) && trim($resStr) === '')) {
			return null;
		}
		return (string)$resStr;
	}

	protected function getDateValue(SheetInterface $sheet, int $column, Row $row): ?\DateTimeInterface {
		$cell = $row->getCellAtIndex($column);
		$resStr = $cell->getValue();
		if($resStr === null || $resStr === '-' || (is_string($resStr) && trim($resStr) === '')) {
			return null;
		}
		return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($resStr);
	}

	protected function getNumericValue(SheetInterface $sheet, int $column, Row $row): ?float {
		$cell = $row->getCellAtIndex($column);
		$resStr = $cell->getValue();
		if(is_numeric($resStr)) {
			return (float)$resStr;
		}
		return null;
	}

	protected function getBoolValue(SheetInterface $sheet, int $column, Row $row): bool {
		$cell = $row->getCellAtIndex($column);
		$resStr = $cell->getValue();
		if($resStr === null || $resStr === '-' || $resStr === '0' || $resStr === 0 ||  (is_string($resStr) && trim($resStr) === '')) {
			return false;
		}
		return true;
	}
}