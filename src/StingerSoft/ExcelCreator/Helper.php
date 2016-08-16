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

/**
 * Some helper methods
 */
trait Helper {

	/**
	 *
	 * @var TranslatorInterface
	 */
	protected $translator;

	/**
	 *
	 * @return TranslatorInterface
	 */
	protected function getTranslator() {
		return $this->translator;
	}

	/**
	 * Decodes html entities in the given text and removes dashes
	 *
	 * @param string $text        	
	 */
	protected function decodeHtmlEntity($text) {
		$text = html_entity_decode($text);
		$text = str_replace('­', '', $text);
		return $text;
	}

	/**
	 * Translates the given key
	 *
	 * @param string $key        	
	 * @param string $domain        	
	 * @return string
	 */
	protected function translate($key, $domain) {
		if($domain === false || $this->getTranslator() === null)
			return $key;
		if($domain === null)
			$domain = 'messages';
		return $this->getTranslator()->trans($key, array(), $domain);
	}

	/**
	 * Sets the title (escaped and shortened) on the given sheet.
	 *
	 * @param \PHPExcel_Worksheet $sheet        	
	 * @param string $title        	
	 */
	protected function setSheetTitle(\PHPExcel_Worksheet $sheet, $title) {
		$sheet->setTitle(\substr(\str_replace($sheet::getInvalidCharacters(), '_', $title), 0, 31));
	}
}