<?php

declare(strict_types=1);

namespace Bassim\BigXlsxBundle\Tests\Entity;

use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use PHPUnit_Framework_TestCase;

class SharedStringXmlTest extends PHPUnit_Framework_TestCase
{

    public function testWrite(): void
    {
        $sharedStringXml = new SharedStringXml();

        $pos = $sharedStringXml->addString("1");
        static::assertEquals(0, $pos);

        $pos = $sharedStringXml->addString("1");
        static::assertEquals(1, $pos);

        $file = $sharedStringXml->getFile();
        $xml = simplexml_load_string(file_get_contents($file));
        static::assertEquals("1", $xml->si[0]->t);
    }
}
