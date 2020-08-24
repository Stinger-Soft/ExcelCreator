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

class DateTimeModifier extends AbstractModifier {

	/**
	 * @return callable
	 */
	public static function createDateModifier(): callable {
		return self::createFormatModifier('m/d/yyyy', 'mm-dd-yy');
	}

	/**
	 * @param bool $addSeconds Note: not working for SPOUT..
	 * @return callable
	 */
	public static function createDateTimeModifier(bool $addSeconds = false): callable {
		return self::createFormatModifier('m/d/yyyy h:mm' . ($addSeconds ? ':ss' : ''), 'm/d/yy h:mm');
	}

	/**
	 * @param bool $addSeconds
	 * @return callable
	 */
	public static function createTimeModifier(bool $addSeconds = false): callable {
		return self::createFormatModifier('h:mm' . ($addSeconds ? ':ss' : ''));
	}
}