<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

/**
 * Interface SetterInterface
 * アクティブな位置を設定する機能のinterface
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Setter
 */
interface SetterInterface
{
    /**
     * @param \ZipArchive $xlsx
     */
    public function set(\ZipArchive $xlsx);
}
