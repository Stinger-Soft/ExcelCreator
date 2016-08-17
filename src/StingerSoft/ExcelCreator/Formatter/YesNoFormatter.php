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

abstract class YesNoFormatter {
	
	public static function createYesNoFormatter($yesLabel, $noLabel) {
		return function($value) use ($yesLabel, $noLabel) {
			return filter_var($value, FILTER_VALIDATE_BOOLEAN) === true ? $yesLabel : $noLabel;
		};
	}
}