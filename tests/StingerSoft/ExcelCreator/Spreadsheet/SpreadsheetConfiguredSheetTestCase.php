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

use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ConfiguredSheetTestCase;
use StingerSoft\ExcelCreator\ExcelFactory;

class SpreadsheetConfiguredSheetTestCase extends ConfiguredSheetTestCase {

	public function getImplementation(): string {
		return ExcelFactory::TYPE_PHP_SPREADSHEET;
	}

	public function testGroupHeaders(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(5));

		$groupedA = (new ColumnBinding('Col A', '[0]'))->setGroupId('g1')->setGroupLabel('Group One');
		$groupedB = (new ColumnBinding('Col B', '[1]'))->setGroupId('g1')->setGroupLabel('Group One');
		$ungrouped = new ColumnBinding('Col C', '[2]');
		$sheet->addColumnBinding($groupedA);
		$sheet->addColumnBinding($groupedB);
		$sheet->addColumnBinding($ungrouped);

		$sheet->applyData();

		$worksheet = $sheet->getSourceSheet();
		// The group header is rendered in row 1, the column header in row 2, data from row 3 on.
		self::assertSame('Group One', $worksheet->getCell('A1')->getValue());
		self::assertSame('Col A', $worksheet->getCell('A2')->getValue());
		self::assertSame('Col B', $worksheet->getCell('B2')->getValue());
		self::assertSame('Col C', $worksheet->getCell('C2')->getValue());

		$merges = $worksheet->getMergeCells();
		// Grouped columns are merged horizontally under a single group header cell.
		self::assertArrayHasKey('A1:B1', $merges);
		// Ungrouped columns span both header rows.
		self::assertArrayHasKey('C1:C2', $merges);
	}

	public function testNoGroupHeaderRowWithoutGroups(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(5));

		$sheet->addColumnBinding(new ColumnBinding('Col A', '[0]'));
		$sheet->addColumnBinding(new ColumnBinding('Col B', '[1]'));

		$sheet->applyData();

		$worksheet = $sheet->getSourceSheet();
		// Without any group, the column header stays in row 1 and no cells are merged.
		self::assertSame('Col A', $worksheet->getCell('A1')->getValue());
		self::assertSame('Col B', $worksheet->getCell('B1')->getValue());
		self::assertCount(0, $worksheet->getMergeCells());
	}

	public function testStripedGroupHeadersUseAlternatingBackground(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(3));

		$sheet->addColumnBinding((new ColumnBinding('Col A', '[0]'))->setGroupId('g1')->setGroupLabel('Group One'));
		$sheet->addColumnBinding((new ColumnBinding('Col B', '[1]'))->setGroupId('g2')->setGroupLabel('Group Two'));
		$sheet->setStripedGroupHeaders(true);
		$sheet->applyData();

		$worksheet = $sheet->getSourceSheet();
		$first = $worksheet->getStyle('A1')->getFill()->getStartColor()->getRGB();
		$second = $worksheet->getStyle('B1')->getFill()->getStartColor()->getRGB();
		// Adjacent groups alternate: the first keeps the default header colour, the second is striped.
		self::assertSame('B8CCE4', $first);
		self::assertSame('A8BCD4', $second);
		self::assertNotSame($first, $second);
	}

	public function testGroupHeadersNotStripedByDefault(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(3));

		$sheet->addColumnBinding((new ColumnBinding('Col A', '[0]'))->setGroupId('g1')->setGroupLabel('Group One'));
		$sheet->addColumnBinding((new ColumnBinding('Col B', '[1]'))->setGroupId('g2')->setGroupLabel('Group Two'));
		$sheet->applyData();

		$worksheet = $sheet->getSourceSheet();
		$first = $worksheet->getStyle('A1')->getFill()->getStartColor()->getRGB();
		$second = $worksheet->getStyle('B1')->getFill()->getStartColor()->getRGB();
		// Striping is opt-in, so both group headers share the default background.
		self::assertSame('B8CCE4', $first);
		self::assertSame('B8CCE4', $second);
	}
}