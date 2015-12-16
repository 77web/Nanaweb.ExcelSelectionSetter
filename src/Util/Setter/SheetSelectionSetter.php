<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\XmlNamespace;
use Nanaweb\ExcelUtil\ZipArchive;

/**
 * Class SheetSelectionSetter
 * ブック内の特定のシートをアクティブにする
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Setter
 */
class SheetSelectionSetter implements SetterInterface
{
    /**
     * @var BookUtil
     */
    private $bookUtil;

    public function __construct(BookUtil $bookUtil = null)
    {
        $this->bookUtil = $bookUtil ? $bookUtil : new BookUtil();
    }

    public function set(ZipArchive $xlsx, $targetSheetName = null)
    {
        // workbook.xmlに書き込む
        $this->setInWorkbook($xlsx, $targetSheetName);

        // 対象のシートのxmlに書き込む
        $this->setInWorksheet($xlsx, $targetSheetName);
    }

    private function setInWorkbook(ZipArchive $xlsx, $targetSheetName)
    {
        $workbookXml = $xlsx->getFromName('xl/workbook.xml');
        if (!$workbookXml) {
            throw new \RuntimeException('Could not find workbook.xml');
        }

        // 対象のシートindexを探す
        $sheets = $this->bookUtil->makeSheetMap($xlsx);
        $i = 0;
        $targetSheetIndex = false;
        foreach ($sheets as $sheetName) {
            if ($sheetName == $targetSheetName) {
                $targetSheetIndex = $i;
                break;
            }
            $i++;
        }
        if ($targetSheetIndex === false) {
            throw new \RuntimeException(sprintf('Could not find sheet"%s" in workbook', $targetSheetName));
        }

        $dom = new \DOMDocument;
        $dom->loadXML($workbookXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $views = $xpath->query('//s:workbook/s:bookViews/s:workbookView');
        if ($views->length == 1) {
            /** @var \DOMElement $workbookView */
            $workbookView = $views->item(0);
            $workbookView->setAttribute('activeTab', $targetSheetIndex);
        } else {
            $workbookView = $dom->createElement('workbookView');
            $workbookView->setAttribute('activeTab', $targetSheetIndex);

            $bookView = $xpath->query('//s:workbook/s:bookViews')->item(0);
            $bookView->appendChild($workbookView);
        }

        $xlsx->addFromString('xl/workbook.xml', $dom->saveXML());
    }

    private function setInWorksheet(ZipArchive $xlsx, $targetSheetName)
    {
        $sheetFileMap = $this->bookUtil->makeSheetFileMap($xlsx);

        foreach ($sheetFileMap as $sheetName => $sheetFile) {
            $worksheetXml = $xlsx->getFromName($sheetFile);
            $tabSelected = $sheetName == $targetSheetName ? 1 : 0;

            $dom = new \DOMDocument;
            $dom->loadXML($worksheetXml);
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

            /** @var \DOMElement $sheetView */
            $sheetView = $xpath->query('//s:worksheet/s:sheetViews/s:sheetView')->item(0);
            $sheetView->setAttribute('tabSelected', $tabSelected);

            $xlsx->addFromString($sheetFile, $dom->saveXML());
        }
    }
}
