<?php

namespace Nanaweb\ExcelSelectionSetter;

use Nanaweb\ExcelSelectionSetter\Strategy\StrategyInterface;
use Nanaweb\ExcelUtil\ZipArchive;

class ExcelSelectionSetter
{
    /**
     * @var StrategyInterface[]
     */
    private $strategies;

    public function __construct()
    {
        $this->strategies = [];
    }

    /**
     * @param StrategyInterface $strategy
     * @return ExcelSelectionSetter
     */
    public function addStrategy(StrategyInterface $strategy)
    {
        $this->strategies[$strategy->getName()] = $strategy;

        return $this;
    }

    /**
     * @param string $xlsxPath
     * @param string $strategyName
     * @param mixed $args
     */
    public function execute($xlsxPath, $strategyName, $args = null)
    {
        $xlsx = new ZipArchive($xlsxPath);

        if (!isset($this->strategies[$strategyName])) {
            throw new \RuntimeException(sprintf('Strategy "%s" not found', $strategyName));
        }

        $this->strategies[$strategyName]->setSelection($xlsx, $args);

        $xlsx->close();
    }
}
