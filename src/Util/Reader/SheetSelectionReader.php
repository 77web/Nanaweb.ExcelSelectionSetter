<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Reader;

/**
 * Class SheetSelectionReader
 * ファイル内で選択されているシートを読み取る
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Reader
 */
class SheetSelectionReader implements ReaderInterface
{
    /**
     * 選択されているシート名を返す
     * @inheritdoc
     */
    public function read(\ZipArchive $xlsx)
    {
        $worksheetXml = $xlsx->getFromName('xl/workbook.xml');
        if (!$worksheetXml) {
            throw new \RuntimeException('Could not find workbook.xml');
        }

        $dom = new \DOMDocument();
        $dom->loadXML($worksheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // シートID => シート名の配列を作る
        $sheets = [];
        foreach ($xpath->query('//s:workbook/s:sheets/s:sheet') as $sheet) {
            /** @var \DOMElement $sheet */
            $sheets[$sheet->getAttribute('sheetId')] = $sheet->getAttribute('name');
        }

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
