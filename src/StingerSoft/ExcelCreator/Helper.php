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

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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
	 * @param Worksheet $sheet
	 * @param string $title        	
	 */
	protected function setSheetTitle(Worksheet $sheet, $title) {
		$sheet->setTitle($this->cleanSheetTitle($title));
	}

	/**
	 * @param $title
	 * @return bool|string
	 */
	protected function cleanSheetTitle($title) {
		return \substr(\str_replace(Worksheet::getInvalidCharacters(), '_', $title), 0, 31);
	}

	/**
	 *
	 * Creates a temporary file with the given content
	 *
	 * @param string $extension
	 * @param string $prefix
	 * @param mixed $content
	 * @return string the filename of the temporary file
	 */
	protected static function createTemporaryFile($extension = null, $prefix = 'stinger_', $content = null) {
		$filename = rtrim(sys_get_temp_dir(), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . uniqid($prefix, true);
		if(null !== $extension) {
			$filename .= '.' . $extension;
		}
		if(null !== $content) {
			file_put_contents($filename, $content);
		}
		return $filename;
	}

}