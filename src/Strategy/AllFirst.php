<?php


namespace Nanaweb\ExcelSelectionSetter\Strategy;

use Nanaweb\ExcelSelectionSetter\Util\Setter\CellSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Util\Setter\SheetSelectionSetter;
use Nanaweb\ExcelUtil\Book as BookUtil;

/**
 * Class AllFirst
 * 全てのシートでA1を選択状態にし、最初のシートを選択する
 *
 * @package Nanaweb\ExcelSelectionSetter\Strategy
 */
class AllFirst implements StrategyInterface
{
    /**
     * @var SheetSelectionSetter
     */
    private $sheetSetter;

    /**
     * @var CellSelectionSetter
     */
    private $cellSetter;

    /**
     * @var BookUtil
     */
    private $bookUtil;

    /**
     * @var string
     */
    private $firstCell;

    public function __construct($firstCell = 'A1', SheetSelectionSetter $sheetSetter = null, CellSelectionSetter $cellSetter = null, BookUtil $bookUtil = null)
    {
        $this->firstCell = $firstCell;
        $this->sheetSetter = $sheetSetter ? $sheetSetter : new SheetSelectionSetter();
        $this->cellSetter = $cellSetter ? $cellSetter : new CellSelectionSetter();
        $this->bookUtil = $bookUtil ? $bookUtil : new BookUtil();
    }

    public function setSelection(\ZipArchive $xlsx)
    {
        $sheets = $this->bookUtil->makeSheetMap($xlsx);
        foreach ($sheets as $sheetName) {
            $this->cellSetter->set($xlsx, $sheetName, $this->firstCell);
        }

        $firstSheet = reset($sheets);
        $this->sheetSetter->set($xlsx, $firstSheet);
    }

    public function getName()
    {
        return 'all_first';
    }
}
