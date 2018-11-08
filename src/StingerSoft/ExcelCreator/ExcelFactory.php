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

namespace StingerSoft\ExcelCreator;

use Symfony\Component\Translation\TranslatorInterface;

class ExcelFactory {

	const TYPE_SPOUT = 'spout';

	const TYPE_PHP_SPREADSHEET = 'spout';

	private function __construct() {
	}

	public static function createConfiguredExcel($type = self::TYPE_PHP_SPREADSHEET, TranslatorInterface $translator = null) {
		$configuredExcel = null;
		switch($type) {
			case self::TYPE_SPOUT:
				$configuredExcel = new \StingerSoft\ExcelCreator\Spout\ConfiguredExcel($translator);
				break;
			default:
				$configuredExcel = new \StingerSoft\ExcelCreator\Spreadsheet\ConfiguredExcel($translator);
		}

		return $configuredExcel;
	}

}