<?php

namespace Nanaweb\ExcelSelectionSetter\Strategy;

use Nanaweb\ExcelSelectionSetter\Util\Setter\CellSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Util\Setter\SheetSelectionSetter;
use Nanaweb\ExcelUtil\Book as BookUtil;

class AllFirstTest extends \PHPUnit_Framework_TestCase
{
    public function test()
    {
        $target = $this->getMock(\ZipArchive::class);

        $sheetSetter = $this->getMock(SheetSelectionSetter::class);
        $cellSetter = $this->getMock(CellSelectionSetter::class);
        $bookUtil = $this->getMock(BookUtil::class);

        $bookUtil->expects($this->once())
            ->method('makeSheetMap')
            ->with($target)
            ->will($this->returnValue([
                1 => 'シート1',
                2 => 'シート2',
            ]))
        ;
        $sheetSetter->expects($this->once())
            ->method('set')
            ->with($target, 'シート1')
        ;
        $cellSetter->expects($this->exactly(2))
            ->method('set')
            ->with($target, $this->logicalOr('シート1', 'シート2'), 'A1')
        ;

        $strategy = new AllFirst('A1', $sheetSetter, $cellSetter, $bookUtil);
        $strategy->setSelection($target);
    }
}
