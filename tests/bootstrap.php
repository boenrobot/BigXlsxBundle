<?php

namespace Bassim\BigXlsxBundle\Tests;

use RuntimeException;

$file = __DIR__ . '/../vendor/autoload.php';
if (!file_exists($file)) {
    throw new RuntimeException('Install dependencies to run test suite. "php composer.phar install --dev"');
}
require_once $file;
