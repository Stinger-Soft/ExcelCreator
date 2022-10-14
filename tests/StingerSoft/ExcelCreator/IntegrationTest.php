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

use PHPUnit\Framework\TestCase;

abstract class IntegrationTest extends TestCase {

	abstract public function getImplementation(): string;

	public function testSimpleCycle(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		$excel->setTitle('TestCase');
		$excel->setCreator('Me');
		$excel->setCompany('StingerSoft');

		$sheet1 = $excel->addSheet('Test Sheet1');

		for($i = 0; $i < 10; $i++) {
			$binding = new ColumnBinding();
			$binding->setLabel('Column ' . $i);
			$binding->setLabelTranslationDomain(false);
			$binding->setBinding('[' . $i . ']'); //
			$binding->setColumnWidth('auto');
			$binding->setWrapText(true);
			$binding->setOutline(1);
			$binding->setHeaderBackgroundColor('000000');
			$binding->setHeaderFontColor('FFFFFF');
			$binding->setDataFontColor('$FFFFFF');
			$binding->setDataBackgroundColor('$FAFAFA');
			$binding->setFormatter(static function ($value) {
				return strtoupper($value);
			});
			$sheet1->addColumnBinding($binding);
		}
		$sheet1->setData($this->getArrayData());
		$sheet1->applyData();
		self::assertTrue(true);
	}

	public function testBindingFunctionCycle(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		$excel->setTitle('TestCase');
		$excel->setCreator('Me');

		$sheet1 = $excel->addSheet('Test Sheet1');

		for($i = 0; $i < 10; $i++) {
			$binding = new ColumnBinding();
			$binding->setLabel('Column ' . $i);
			$binding->setLabelTranslationDomain(false);
			$binding->setColumnWidth('auto');
			$binding->setBinding(static function (ColumnBinding $bind, $item) {
				return 'Bound via callable';
			});
			$binding->setLinkUrl(static function (ColumnBinding $bind, $item) use ($i) {
				return 'http://www.google.com?q=' . $item[$i];
			});
			$sheet1->addColumnBinding($binding);
		}
		$sheet1->setData($this->getArrayData());
		$sheet1->applyData();
		self::assertTrue(true);
	}

	protected function getArrayData($count = 10, $columns = 10): array {
		$data = [];
		for($i = 0; $i < $count; $i++) {
			$item = [];
			for($j = 0; $j < $columns; $j++) {
				$item[$j] = 'Test ' . $i . ':' . $j;
			}
			$data[] = $item;
		}
		return $data;
	}
}