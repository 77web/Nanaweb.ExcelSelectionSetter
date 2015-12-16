<?php

namespace Nanaweb\ExcelSelectionSetter\Strategy;

use Nanaweb\ExcelUtil\ZipArchive;

interface StrategyInterface
{
    public function getName();

    public function setSelection(ZipArchive $xlsx);
}
