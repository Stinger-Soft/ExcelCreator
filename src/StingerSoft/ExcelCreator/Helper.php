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

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use function str_replace;
use function substr;
use Symfony\Contracts\Translation\TranslatorInterface;

/**
 * Some helper methods
 */
trait Helper {

	/**
	 *
	 * @var TranslatorInterface|null
	 */
	protected $translator;

	/**
	 *
	 * Creates a temporary file with the given content
	 *
	 * @param string $extension
	 * @param string $prefix
	 * @param mixed $content
	 * @return string the filename of the temporary file
	 */
	protected static function createTemporaryFile($extension = null, $prefix = 'stinger_', $content = null): string {
		$filename = rtrim(sys_get_temp_dir(), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . uniqid($prefix, true);
		if(null !== $extension) {
			$filename .= '.' . $extension;
		}
		if(null !== $content) {
			file_put_contents($filename, $content);
		}
		return $filename;
	}

	/**
	 *
	 * @return TranslatorInterface|null
	 */
	protected function getTranslator(): ?TranslatorInterface {
		return $this->translator;
	}

	/**
	 * Decodes html entities in the given text and removes dashes
	 *
	 * @param string $text
	 * @return string
	 */
	protected function decodeHtmlEntity(string $text): string {
		$text = html_entity_decode($text);
		$text = (string)str_replace('­', '', $text);
		return $text;
	}

	/**
	 * Translates the given key
	 *
	 * @param string $key
	 * @param string $domain
	 * @return string
	 */
	protected function translate(string $key, ?string $domain = null) : string {
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
	protected function setSheetTitle(Worksheet $sheet, string $title) : void {
		$sheet->setTitle($this->cleanSheetTitle($title));
	}

	/**
	 * @param $title
	 * @return bool|string
	 */
	protected function cleanSheetTitle(string $title) {
		return substr(str_replace(Worksheet::getInvalidCharacters(), '_', $title), 0, 31);
	}

}