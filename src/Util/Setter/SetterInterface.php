<?php

namespace Nanaweb\ExcelSelectionSetter\Util\Setter;

use Nanaweb\ExcelUtil\ZipArchive;

/**
 * Interface SetterInterface
 * アクティブな位置を設定する機能のinterface
 *
 * @package Nanaweb\ExcelSelectionSetter\Util\Setter
 */
interface SetterInterface
{
    /**
     * @param ZipArchive $xlsx
     */
    public function set(ZipArchive $xlsx);
}
