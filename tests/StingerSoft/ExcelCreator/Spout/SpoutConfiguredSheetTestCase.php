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

use PhpOffice\PhpSpreadsheet\IOFactory;
use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ConfiguredSheetTestCase;
use StingerSoft\ExcelCreator\ExcelFactory;

class SpoutConfiguredSheetTestCase extends ConfiguredSheetTestCase {

	public function getImplementation(): string {
		return ExcelFactory::TYPE_SPOUT;
	}

	/**
	 * OpenSpout is a streaming writer and cannot merge cells, so the group label is only written into the first
	 * column of each group; the remaining group columns and ungrouped columns are left blank in the group header row.
	 */
	public function testGroupHeadersDegradedWithoutMerges(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(5));

		$sheet->addColumnBinding((new ColumnBinding('Col A', '[0]'))->setGroupId('g1')->setGroupLabel('Group One'));
		$sheet->addColumnBinding((new ColumnBinding('Col B', '[1]'))->setGroupId('g1')->setGroupLabel('Group One'));
		$sheet->addColumnBinding(new ColumnBinding('Col C', '[2]'));
		$sheet->applyData();

		$worksheet = $this->writeAndReload($excel);

		// Group header row (row 1): label in the first column of the group only, rest blank (no merging possible).
		self::assertSame('Group One', $this->plainValue($worksheet, 'A1'));
		self::assertNull($worksheet->getCell('B1')->getValue());
		self::assertNull($worksheet->getCell('C1')->getValue());
		// Column header row (row 2).
		self::assertSame('Col A', $this->plainValue($worksheet, 'A2'));
		self::assertSame('Col B', $this->plainValue($worksheet, 'B2'));
		self::assertSame('Col C', $this->plainValue($worksheet, 'C2'));
		// OpenSpout does not support merged cells.
		self::assertCount(0, $worksheet->getMergeCells());
	}

	public function testNoGroupHeaderRowWithoutGroups(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(5));

		$sheet->addColumnBinding(new ColumnBinding('Col A', '[0]'));
		$sheet->addColumnBinding(new ColumnBinding('Col B', '[1]'));
		$sheet->applyData();

		$worksheet = $this->writeAndReload($excel);

		// Without any group the column header stays in row 1.
		self::assertSame('Col A', $this->plainValue($worksheet, 'A1'));
		self::assertSame('Col B', $this->plainValue($worksheet, 'B1'));
	}

	public function testStripedGroupHeadersUseAlternatingBackground(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(3));

		$sheet->addColumnBinding((new ColumnBinding('Col A', '[0]'))->setGroupId('g1')->setGroupLabel('Group One'));
		$sheet->addColumnBinding((new ColumnBinding('Col B', '[1]'))->setGroupId('g2')->setGroupLabel('Group Two'));
		$sheet->setStripedGroupHeaders(true);
		$sheet->applyData();

		$worksheet = $this->writeAndReload($excel);
		$first = $worksheet->getStyle('A1')->getFill()->getStartColor()->getRGB();
		$second = $worksheet->getStyle('B1')->getFill()->getStartColor()->getRGB();
		// The second group's header cell is striped, so the two adjacent group headers differ.
		self::assertNotSame($first, $second);
		self::assertSame('A8BCD4', $second);
	}

	/**
	 * Values written by OpenSpout are read back by PhpSpreadsheet as RichText objects; normalise them to plain text.
	 */
	private function plainValue(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet, string $coordinate): ?string {
		$value = $worksheet->getCell($coordinate)->getValue();
		if($value instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
			return $value->getPlainText();
		}
		return $value === null ? null : (string)$value;
	}

	private function writeAndReload($excel): \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet {
		$file = sys_get_temp_dir() . '/excel-creator-group-test-' . uniqid('', true) . '.xlsx';
		try {
			$excel->writeToFile($file);
			$worksheet = IOFactory::load($file)->getActiveSheet();
		} finally {
			@unlink($file);
		}
		return $worksheet;
	}

}