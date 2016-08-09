<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

$ME_URL = $_ENV['APP_URL']."?path=";
/**
 * Common Html functions
 *
 * @SuppressWarnings(PHPMD.UnusedPrivateMethod) For readWPNode
 */
class WaffleHtml
{

    static $fStyle = array('name' => 'Arial', 'size' => 16, 'color' => '757575', 'bold' => false);
    static $h2Style = array('name' => 'Arial', 'size' => 22, 'color' => '757575', 'bold' => true);
    static $h3Style = array('name' => 'Arial', 'size' => 24, 'color' => '333333', 'bold' => false);
    static $aStyle = array('name' => 'Arial', 'size' => 16, 'color' => 'e83d84', 'bold' => false);

    static $pStyle = array( 'spaceAfter' => 222 );

    /**
     * Add HTML parts.
     *
     * Note: $stylesheet parameter is removed to avoid PHPMD error for unused parameter
     *
     * @param \PhpOffice\PhpWord\Element\AbstractContainer $element Where the parts need to be added
     * @param string $html The code to parse
     * @param bool $fullHTML If it's a full HTML, no need to add 'body' tag
     * @return void
     */
    public static function addHtml($element, $html, $fullHTML = false)
    {
      $doc = new DOMDocument();
      $doc->loadHTML('<?xml encoding="utf-8" ?>' . $html);

      foreach ($doc->documentElement->childNodes as $body) {
        foreach ($body->childNodes as $node) {
          self::handle($node, $element);
        }

      }
    }

    public static function handle($node, $element) {
      //echo $node->tagName;

      if ($node->nodeType === XML_TEXT_NODE) {
        if (strlen($node->textContent) > 3) {
          $element->addText($node->textContent, self::$fStyle, self::$pStyle);
        }
        return;
      }

      switch ($node->tagName) {
        case "h2":
          $element->addText($node->textContent, self::$h2Style, self::$pStyle);
          return;
        case "h3":
          $element->addText($node->textContent, self::$h3Style, self::$pStyle);
          return;
        case "p":
          //$element->addText($node->textContent, self::$fStyle, self::$pStyle);
          break;
        case "a":
          $element->addLink($ME_URL.$node->getAttribute('href'), $node->textContent, self::$aStyle, self::$pStyle);
          return;
        default:
          //$element->addText("Unhandled tag".$node->tagName, self::$fStyle, self::$pStyle);
      }

      if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
          self::handle($child, $element);
        }
      }
    }

}
