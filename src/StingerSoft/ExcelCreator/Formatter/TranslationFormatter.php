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

use Symfony\Component\Translation\TranslatorInterface;

/**
 * Creates a deleage to translate binding results
 */
abstract class TranslationFormatter {

	public static function createTranslationFormatter(TranslatorInterface $translator, $domain) {
		return function ($value) use ($translator, $domain) {
			return $translator->trans($value, array(), $domain);
		};
	}
}