<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Reader;

use Nanaweb\ExcelSelectionSetter\Exception\BrokenWorkbookException;
use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\XmlNamespace;
use Nanaweb\ExcelUtil\ZipArchive;

class CellSelectionReader implements ReaderInterface
{
    /**
     * @var BookUtil
     */
    private $bookUtil;

    public function __construct(BookUtil $bookUtil = null)
    {
        $this->bookUtil = $bookUtil ? $bookUtil : new BookUtil();
    }

    public function read(ZipArchive $xlsx, $targetSheetName = null)
    {
        // $targetSheetNameのシートxmlファイルを取得
        $sheetFileMap = $this->bookUtil->makeSheetFileMap($xlsx);
        if (!isset($sheetFileMap[$targetSheetName])) {
            throw BrokenWorkbookException::createForMissingWorksheet($targetSheetName);
        }

        $worksheetXml = $xlsx->getFromName($sheetFileMap[$targetSheetName]);
        if (!$worksheetXml) {
            throw BrokenWorkbookException::createForBrokenWorksheet($targetSheetName);
        }

        // 選択されているセルの番地を調べる
        $dom = new \DOMDocument;
        $dom->loadXML($worksheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $activeCell = null;
        $selections = $xpath->query('//s:worksheet/s:sheetViews/s:sheetView/s:selection');
        if ($selections->length == 1) {
            /** @var \DOMElement $selection */
            $selection = $selections->item(0);
            $activeCell = $selection->getAttribute('activeCell');
        }

        return $activeCell ? : 'A1';
    }
}
