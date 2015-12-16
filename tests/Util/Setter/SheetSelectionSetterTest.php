<?php


namespace Nanaweb\ExcelSelectionSetter\Util\Setter;


class SheetSelectionSetterTest extends \PHPUnit_Framework_TestCase
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
            ->with('xl/workbook.xml')
            ->will($this->returnValue($xml))
        ;
        $zip->expects($this->once())
            ->method('addFromString')
            ->with('xl/workbook.xml', $this->callback(function($xml){
                return strpos($xml, 'activeTab="1"') !== false;
            }))
        ;

        $setter = new SheetSelectionSetter();
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
