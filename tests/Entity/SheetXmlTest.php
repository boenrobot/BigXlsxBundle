<?php

declare(strict_types=1);

namespace Bassim\BigXlsxBundle\Tests\Entity;

use Bassim\BigXlsxBundle\Entity\SharedStringXml;
use Bassim\BigXlsxBundle\Entity\SheetXml;
use PHPUnit_Framework_TestCase;

class SheetXmlTest extends PHPUnit_Framework_TestCase
{

    public function testWrite(): void
    {
        $data = array(
            array("1", "2", "3")
        );

        $sharedStringXml = new SharedStringXml();
        $sheetXml = new SheetXml($sharedStringXml);
        $sheetXml->addRow($data[0]);


        $file = $sheetXml->getFile();
        $xml = simplexml_load_string(file_get_contents($file));

        static::assertEquals(0, (string)$xml->sheetData->row[0]->c[0]->v);
    }
}
