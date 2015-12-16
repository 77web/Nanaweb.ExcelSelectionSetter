<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Reader;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\XmlNamespace;
use Nanaweb\ExcelUtil\ZipArchive;

/**
 * Class SheetSelectionReader
 * ファイル内で選択されているシートを読み取る
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Reader
 */
class SheetSelectionReader implements ReaderInterface
{
    /**
     * @var BookUtil
     */
    private $bookUtil;

    public function __construct(BookUtil $bookUtil = null)
    {
        $this->bookUtil = $bookUtil ? $bookUtil : new BookUtil();
    }

    /**
     * 選択されているシート名を返す
     * @inheritdoc
     */
    public function read(ZipArchive $xlsx)
    {
        $worksheetXml = $xlsx->getFromName('xl/workbook.xml');
        if (!$worksheetXml) {
            throw new \RuntimeException('Could not find workbook.xml');
        }

        $sheets = $this->bookUtil->makeSheetMap($xlsx);

        $dom = new \DOMDocument();
        $dom->loadXML($worksheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        // 選択されているシートIDを調べる
        $activeSheetId = null;
        $views = $xpath->query('//s:workbook/s:bookViews/s:workbookView');
        if ($views->length == 1) {
            /** @var \DOMElement $workbookView */
            $workbookView = $views->item(0);
            $activeSheetId = $workbookView->getAttribute('activeTab');
        }

        return $activeSheetId && isset($sheets[$activeSheetId]) ? $sheets[$activeSheetId] : reset($sheets);
    }


}
