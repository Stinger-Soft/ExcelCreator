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
use Symfony\Contracts\Translation\TranslatorInterface;

abstract class ConfiguredSheetTest extends TestCase {

	abstract public function getImplementation(): string;

	public function testSetters(): void {
		$translator = $this->getMockBuilder(TranslatorInterface::class)->setMethods(['trans'])->getMockForAbstractClass();
		$translator->method('trans')->willReturn('translated');

		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());
		$sheet = $excel->addSheet('TestSheet');
		$sheet->setData($this->getArrayData(10));

		$simpleBinding = new ColumnBinding();
		$simpleBinding->setBinding('[1]');
		$simpleBinding->setLabel('simpleBinding');
		$simpleBinding->setLinkUrl('http://www.google.de');
		$sheet->addColumnBinding($simpleBinding);

		$index = $sheet->getIndexForBinding($simpleBinding);
		self::assertEquals(0, $index);

		$sheet->applyData();

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