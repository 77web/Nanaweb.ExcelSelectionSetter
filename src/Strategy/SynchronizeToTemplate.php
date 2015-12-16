<?php

namespace Nanaweb\ExcelSelectionSetter\Strategy;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\ZipArchive;
use Nanaweb\ExcelSelectionSetter\Util\Reader\CellSelectionReader;
use Nanaweb\ExcelSelectionSetter\Util\Reader\SheetSelectionReader;
use Nanaweb\ExcelSelectionSetter\Util\Setter\CellSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Util\Setter\SheetSelectionSetter;

/**
 * Class SynchronizeToTemplate
 * 全シートのセル選択状態＋シート選択状態をテンプレートxlsxファイルの選択状態と同じにする
 *
 * @package Nanaweb\ExcelSelectionSetter\Strategy
 */
class SynchronizeToTemplate implements StrategyInterface
{
    /**
     * @var SheetSelectionReader
     */
    private $sheetReader;

    /**
     * @var CellSelectionReader
     */
    private $cellReader;

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

    public function __construct(SheetSelectionReader $sheetReader = null, CellSelectionReader $cellReader = null, SheetSelectionSetter $sheetSetter = null, CellSelectionSetter $cellSetter = null, BookUtil $bookUtil = null)
    {
        $this->sheetReader = $sheetReader ? $sheetReader : new SheetSelectionReader();
        $this->cellReader = $cellReader ? $cellReader : new CellSelectionReader();
        $this->sheetSetter = $sheetSetter ? $sheetSetter : new SheetSelectionSetter();
        $this->cellSetter = $cellSetter ? $cellSetter : new CellSelectionSetter();
        $this->bookUtil = $bookUtil ? $bookUtil : new BookUtil();
    }

    /**
     * @param ZipArchive $xlsx
     * @param ZipArchive|null $template
     */
    public function setSelection(ZipArchive $xlsx, ZipArchive $template = null)
    {
        if (!($template instanceOf ZipArchive)) {
            throw new \RuntimeException('Invalid template');
        }

        // テンプレートで選択されているシートを取得
        $selectedSheet = $this->sheetReader->read($template);

        // テンプレートの各シートで選択されているセルを取得
        $sheets = $this->bookUtil->makeSheetMap($template);
        $selectedCells = [];
        foreach ($sheets as $sheetId => $sheetName) {
            $selectedCells[$sheetName] = $this->cellReader->read($template, $sheetName);
        }

        // 選択シートを書き込む
        $this->sheetSetter->set($xlsx, $selectedSheet);

        // 各シートの選択セルを書き込む
        foreach ($selectedCells as $sheetName => $cellPosition) {
            $this->cellSetter->set($xlsx, $sheetName, $cellPosition);
        }
    }

    public function getName()
    {
        return 'synchronize_to_template';
    }
}
