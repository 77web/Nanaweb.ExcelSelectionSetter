<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\Book as BookUtil;

class CellSelectionSetterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @param string $xmlFixtureFile
     * @dataProvider provideData
     */
    public function test($xmlFixtureFile)
    {
        $xml = file_get_contents(__DIR__.'/../../data/xml/'.$xmlFixtureFile);

        $zip = $this->getMock(\ZipArchive::class);
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

        $setter = new CellSelectionSetter($bookUtil);
        $setter->set($zip, 'テスト', 'A1');
    }

    public function provideData()
    {
        return [
            ['worksheet_with_activeCell.xml'],
            ['worksheet_without_activeCell.xml'],
        ];
    }
}
