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

namespace StingerSoft\ExcelCreator\CellModifier;

class NumberModifier extends AbstractModifier {

	/**
	 * @return callable
	 */
	public static function createIntFormatter(): callable {
		return self::createFormatModifier('#,##0');
	}

	/**
	 * @return callable
	 */
	public static function createFloatFormatter(): callable {
		return self::createFormatModifier('#,##0.00');
	}

}