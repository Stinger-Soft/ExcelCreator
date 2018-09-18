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
namespace StingerSoft\ExcelCreator\Formatter;

use PHPUnit\Framework\TestCase;

class YesNoFormatterTest extends TestCase {

	public function testCreateTranslationFormatter() {
		$pirateYesNoformatter = YesNoFormatter::createYesNoFormatter('Arrrr!', 'Avast!');
		
		$this->assertEquals('Arrrr!', $pirateYesNoformatter(true));
		$this->assertEquals('Avast!', $pirateYesNoformatter(false));
		$this->assertEquals('Avast!', $pirateYesNoformatter('nonono'));
	}
}