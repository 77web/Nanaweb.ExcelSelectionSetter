<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Reader;

use Nanaweb\ExcelUtil\Book as BookUtil;
use Nanaweb\ExcelUtil\ZipArchive;

class CellSelectionReaderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @param string $xmlFixtureFile
     * @param string $expect
     * @dataProvider provideData
     */
    public function test($xmlFixtureFile, $expect)
    {
        $xml = file_get_contents(__DIR__.'/../../data/xml/'.$xmlFixtureFile);

        $zip = $this->getMockBuilder(ZipArchive::class)->disableOriginalConstructor()->getMock();
        $zip->expects($this->atLeastOnce())
            ->method('getFromName')
            ->with('xl/worksheet/sheet1.xml')
            ->will($this->returnValue($xml))
        ;

        $bookUtil = $this->getMock(BookUtil::class);
        $dummySheetMap = [
            'テスト' => 'xl/worksheet/sheet1.xml',
        ];
        $bookUtil->expects($this->once())
            ->method('makeSheetFileMap')
            ->with($zip)
            ->will($this->returnValue($dummySheetMap))
        ;

        $reader = new CellSelectionReader($bookUtil);
        $this->assertEquals($expect, $reader->read($zip, 'テスト'));
    }

    public function provideData()
    {
        return [
            ['worksheet_with_activeCell.xml', 'A2'],
            ['worksheet_without_activeCell.xml', 'A1'],
        ];
    }
}
