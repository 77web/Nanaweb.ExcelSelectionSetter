<?php

namespace Nanaweb\ExcelSelectionSetter\Strategy;

interface StrategyInterface
{
    public function getName();

    public function setSelection(\ZipArchive $xslx);
}
