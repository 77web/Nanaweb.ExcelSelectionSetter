<?php

use Nanaweb\ExcelSelectionSetter\ExcelSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Strategy\AllFirst;
use Nanaweb\ExcelSelectionSetter\Strategy\SynchronizeToTemplate;
use Nanaweb\ExcelUtil\ZipArchive;

class FunctionalTest extends \PHPUnit_Framework_TestCase
{
    public function test()
    {
        $service = new ExcelSelectionSetter();
        $service->addStrategy(new AllFirst());
        $service->addStrategy(new SynchronizeToTemplate());

        $originalFile = __DIR__ . '/data/functional/target.xlsx';
        $templateFile = new ZipArchive();
        $templateFile->open(__DIR__ . '/data/functional/template.xlsx');

        // allfirst
        $allFirstTarget = __DIR__ . '/data/functional/all_first.xlsx';
        copy($originalFile, $allFirstTarget);
        $service->execute($allFirstTarget, 'all_first');


        // synchro
        $syncTarget = __DIR__ . '/data/functional/sync.xlsx';
        copy($originalFile, $syncTarget);
        $service->execute($syncTarget, 'synchronize_to_template', $templateFile);

        $this->assertTrue(true);
    }
}
