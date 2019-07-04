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

use StingerSoft\ExcelCreator\ConfiguredExcelTest;
use StingerSoft\ExcelCreator\ExcelFactory;

class SpoutConfiguredExcelTest extends ConfiguredExcelTest {

	public function getImplementation(): string {
		return ExcelFactory::TYPE_SPOUT;
	}
}