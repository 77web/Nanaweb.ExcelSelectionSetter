<?php

namespace Nanaweb\ExcelSelectionSetter\Exception;

class BrokenWorkbookException extends \Exception
{
    public static function createForMissingWorksheet($sheetName)
    {
        return new self(sprintf('worksheet "%s" was not found', $sheetName));
    }

    public static function createForBrokenWorksheet($sheetName)
    {
        return new self(sprintf('worksheet "%s" is broken', $sheetName));
    }
}
