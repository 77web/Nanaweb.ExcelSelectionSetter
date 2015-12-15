<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Reader;

/**
 * Interface ReaderInterface
 * 現在アクティブな位置を読み取る機能のinterface
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Reader
 */
interface ReaderInterface
{
    /**
     * @param \ZipArchive $xlsx
     * @return string
     */
    public function read(\ZipArchive $xlsx);
}
