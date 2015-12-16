<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\ZipArchive;
use Nanaweb\ExcelUtil\Book as BookUtil;

class SheetSelectionSetterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @param string $xmlFixtureFile
     * @dataProvider provideData
     */
    public function test($xmlFixtureFile)
    {
        $bookXml = file_get_contents(__DIR__.'/../../data/xml/'.$xmlFixtureFile);
        $sheetXml = file_get_contents(__DIR__.'/../../data/xml/worksheet.xml');

        $zip = $this->getMockBuilder(ZipArchive::class)->disableOriginalConstructor()->getMock();
        $bookUtil = $this->getMock(BookUtil::class);

        $bookUtil->expects($this->atLeastOnce())
            ->method('makeSheetMap')
            ->with($zip)
            ->will($this->returnValue([1 => 'シート1', 2 => 'シート2']))
        ;
        $bookUtil->expects($this->once())
            ->method('makeSheetFileMap')
            ->with($zip)
            ->will($this->returnValue(['シート1' => 'xl/sheets/sheet1.xml', 'シート2' => 'xl/sheets/sheet2.xml']))
        ;
        
        $zip->expects($this->exactly(3))
            ->method('getFromName')
            ->with($this->logicalOr('xl/workbook.xml', 'xl/sheets/sheet1.xml', 'xl/sheets/sheet2.xml'))
            ->will($this->returnCallback(function($arg) use ($bookXml, $sheetXml){
                return $arg == 'xl/workbook.xml' ? $bookXml : $sheetXml;
            }))
        ;
        $zip->expects($this->exactly(3))
            ->method('addFromString')
            ->with($this->logicalOr('xl/workbook.xml', 'xl/sheets/sheet1.xml', 'xl/sheets/sheet2.xml'), $this->callback(function($xml){
                return strpos($xml, 'activeTab="1"') !== false || strpos($xml, 'tabSelected="1"') !== false || strpos($xml, 'tabSelected="0"') !== false;
            }))
        ;

        $setter = new SheetSelectionSetter($bookUtil);
        $setter->set($zip, 'シート2');
    }

    public function provideData()
    {
        return [
            ['workbook_with_activeTab.xml'],
            ['workbook_without_activeTab.xml'],
        ];
    }
}
