<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\XmlNamespace;

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

    public function set(\ZipArchive $xlsx, $targetSheetName = null)
    {
        $worksheetXml = $xlsx->getFromName('xl/workbook.xml');
        if (!$worksheetXml) {
            throw new \RuntimeException('Could not find workbook.xml');
        }

        // 対象のシートIDを探す
        $sheets = $this->bookUtil->makeSheetMap($xlsx);
        $targetSheetId = array_search($targetSheetName, $sheets);
        if ($targetSheetId === false) {
            throw new \RuntimeException(sprintf('Could not find sheet"%s" in workbook', $targetSheetName));
        }

        $dom = new \DOMDocument;
        $dom->loadXML($worksheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $views = $xpath->query('//s:workbook/s:bookViews/s:workbookView');
        if ($views->length == 1) {
            /** @var \DOMElement $workbookView */
            $workbookView = $views->item(0);
            $workbookView->setAttribute('activeTab', $targetSheetId);
        } else {
            $workbookView = $dom->createElement('workbookView');
            $workbookView->setAttribute('activeTab', $targetSheetId);

            $bookView = $xpath->query('//s:workbook/s:bookViews')->item(0);
            $bookView->appendChild($workbookView);
        }

        $xlsx->addFromString('xl/workbook.xml', $dom->saveXML());
    }

}
