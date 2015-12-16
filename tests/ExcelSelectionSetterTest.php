<?php

namespace Nanaweb\ExcelSelectionSetter;

use Nanaweb\ExcelSelectionSetter\Strategy\StrategyInterface;
use Nanaweb\ExcelUtil\ZipArchive;

class ExcelSelectionSetterTest extends \PHPUnit_Framework_TestCase
{
    public function test()
    {
        $arg = [];
        $path = __DIR__.'/data/dummy.xlsx';

        $strategy1 = $this->createStrategy('test01');
        $strategy2 = $this->createStrategy('test02');

        $strategy1->expects($this->once())
            ->method('setSelection')
            ->with($this->isInstanceOf(ZipArchive::class), $arg)
        ;
        $strategy2->expects($this->never())
            ->method('setSelection')
        ;

        $service = new ExcelSelectionSetter();
        $service->addStrategy($strategy1);
        $service->addStrategy($strategy2);

        $service->execute($path, 'test01', $arg);
    }

    private function createStrategy($name)
    {
        $strategy = $this->getMock(StrategyInterface::class);
        $strategy->expects($this->atLeastOnce())
            ->method('getName')
            ->will($this->returnValue($name))
        ;

        return $strategy;
    }
}
