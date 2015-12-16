<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\XmlNamespace;

/**
 * Class CellSelectionSetter
 * 指定シート内で指定セルを選択状態にする
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Setter
 */
class CellSelectionSetter implements SetterInterface
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
     * @param \ZipArchive $xlsx
     * @param string $targetSheetName
     * @param string $targetCell
     */
    public function set(\ZipArchive $xlsx, $targetSheetName = null, $targetCell = null)
    {
        // $targetSheetNameのシートxmlファイルを取得
        $sheetFileMap = $this->bookUtil->makeSheetFileMap($xlsx);
        if (!isset($sheetFileMap[$targetSheetName])) {
            throw new \RuntimeException(sprintf('Worksheet "%s" not found', $targetSheetName));
        }
        $path = $sheetFileMap[$targetSheetName];

        $worksheetXml = $xlsx->getFromName($path);
        if (!$worksheetXml) {
            throw new \RuntimeException(sprintf('Worksheet "%s" is empty or broken.', $targetSheetName));
        }

        // selectionがあればactiveCellセット、なければselectionを作ってからactiveCellセット
        $dom = new \DOMDocument;
        $dom->loadXML($worksheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $selections = $xpath->query('//s:worksheet/s:sheetViews/s:sheetView/s:selection');
        if ($selections->length == 1) {
            /** @var \DOMElement $selection */
            $selection = $selections->item(0);
            $selection->setAttribute('activeCell', $targetCell);
            $selection->setAttribute('sqref', $targetCell);
        } else {
            $selection = $dom->createElement('sheetView');
            $selection->setAttribute('activeCell', $targetCell);
            $selection->setAttribute('sqref', $targetCell);

            $sheetView = $xpath->query('//s:worksheet/s:sheetViews/s:sheetView')->item(0);
            $sheetView->appendChild($selection);
        }

        // zipに書き込む
        $xlsx->addFromString($path, $dom->saveXML());
    }
}
