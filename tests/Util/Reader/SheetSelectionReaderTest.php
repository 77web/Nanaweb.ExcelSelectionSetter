<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Reader;


class SheetSelectionReaderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @param string $xmlFixtureFile
     * @param string $expect
     * @dataProvider provideData
     */
    public function test($xmlFixtureFile, $expect)
    {
        $xml = file_get_contents(__DIR__.'/../../data/xml/'.$xmlFixtureFile);

        $zip = $this->getMock(\ZipArchive::class);
        $zip->expects($this->once())
            ->method('getFromName')
            ->with('xl/workbook.xml')
            ->will($this->returnValue($xml))
        ;

        $reader = new SheetSelectionReader();
        $this->assertEquals($expect, $reader->read($zip));
    }

    public function provideData()
    {
        return [
            ['workbook_with_activeTab.xml', 'シート3'],
            ['workbook_without_activeTab.xml', 'シート1'],
        ];
    }
}
