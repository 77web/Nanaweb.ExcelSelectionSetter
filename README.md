[![Build Status](https://travis-ci.org/77web/Nanaweb.ExcelSelectionSetter.svg?branch=master)](https://travis-ci.org/77web/Nanaweb.ExcelSelectionSetter)

# Nanaweb.ExcelSelectionSetter

Sets active sheet and active cells of xlsx files.

# Installation

Use composer.

```bash
$ composer require nanaweb/excel-selection-setter
```

# Usage

## AllFirst

Select first cell in each sheet, first sheet of the workbook.

```php
<?php

use Nanaweb\ExcelSelectionSetter\ExcelSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Strategy\AllFirst;


$setter = new ExcelSelectionSetter();
$setter->addStrategy(new AllFirst());

$setter->execute($pathToYourXlsx, 'all_first');

```

## SynchronizeToTemplate

Select the sheet which is selected in the template, cells which are selected in the template.

```php
<?php

use Nanaweb\ExcelSelectionSetter\ExcelSelectionSetter;
use Nanaweb\ExcelSelectionSetter\Strategy\SynchronizeToTemplate;


$setter = new ExcelSelectionSetter();
$setter->addStrategy(new SynchronizeToTemplate());

$setter->execute($pathToYourXlsx, 'synchronize_to_template');

```


## Create your own strategy

Any classes which implement `StrategyInterface` are accepted as strategy.
