<?php


namespace Nanaweb\ExcelSelectionSetter\Strategy;

use Nanaweb\ExcelSelectionSetter\Util\Reader\CellSelectionReader;
use Nanaweb\ExcelSelectionSetter\Util\Reader\SheetSelectionReader;
use Nanaweb\ExcelSelectionSetter\Util\Setter\CellSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Util\Setter\SheetSelectionSetter;
use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\ZipArchive;

class SynchronizeToTemplateTest extends \PHPUnit_Framework_TestCase
{
    public function test()
    {
        $target = $this->getMockBuilder(ZipArchive::class)->disableOriginalConstructor()->getMock();
        $template = $this->getMockBuilder(ZipArchive::class)->disableOriginalConstructor()->getMock();

        $sheetReader = $this->getMock(SheetSelectionReader::class);
        $cellReader = $this->getMock(CellSelectionReader::class);
        $sheetSetter = $this->getMock(SheetSelectionSetter::class);
        $cellSetter = $this->getMock(CellSelectionSetter::class);
        $bookUtil = $this->getMock(BookUtil::class);

        $sheetReader->expects($this->once())
            ->method('read')
            ->with($target)
            ->will($this->returnValue('シート2'))
        ;
        $bookUtil->expects($this->once())
            ->method('makeSheetMap')
            ->with($template)
            ->will($this->returnValue([
                1 => 'シート1',
                2 => 'シート2',
            ]))
        ;
        $cellReader->expects($this->exactly(2))
            ->method('read')
            ->with($template, $this->logicalOr('シート1', 'シート2'))
            ->will($this->returnValue('A100'))
        ;
        $sheetSetter->expects($this->once())
            ->method('set')
            ->with($target, 'シート2')
        ;
        $cellSetter->expects($this->exactly(2))
            ->method('set')
            ->with($target, $this->logicalOr('シート1', 'シート2'), 'A100')
        ;

        $strategy = new SynchronizeToTemplate($sheetReader, $cellReader, $sheetSetter, $cellSetter, $bookUtil);
        $strategy->setSelection($target, $template);
    }
}
