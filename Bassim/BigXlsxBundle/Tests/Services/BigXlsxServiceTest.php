<?php

declare(strict_types=1);

namespace Bassim\BigXlsxBundle\Tests\Services;

use Bassim\BigXlsxBundle\Services\BenchmarkService;
use Bassim\BigXlsxBundle\Services\BigXlsxService;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
use PHPUnit_Framework_TestCase;

class BigXlsxServiceTest extends PHPUnit_Framework_TestCase
{

    private $rowCount = 100;

    /**
     * @throws SpreadsheetException
     * @throws WriterException
     * @throws ReaderException
     */
    public function testService(): void
    {
        $data = [];
        $service = new BigXlsxService(); //get('bassim_big_xlsx.service');

        $data[] = array("id", "name");
        for ($i = 0; $i < $this->rowCount; $i++) {
            $data[] = array("1_a_" . $i, "1_b_" . $i, "1_c_" . $i);
        }

        $service->addSheet(0, "test Sheet_0", $data);
        $data[] = array("id2", "name2");
        for ($i = 0; $i < $this->rowCount; $i++) {
            $data[] = array("2_a_" . $i, "2_b_" . $i);
        }

        $service->addSheet(1, "test Sheet_1", $data);
        $file = $service->getFile();

        $reader = new Xlsx();
        $reader->load($file);
        $worksheetNames = $reader->listWorksheetNames($file);
        static::assertCount(2, $worksheetNames);


    }

    /**
     * @throws SpreadsheetException
     * @throws WriterException
     * @throws ReaderException
     */
    public function testServiceAddCustomSheet(): void
    {
        $data = [];
        $service = new BigXlsxService(); //get('bassim_big_xlsx.service');

        $data[] = array("id", "name");
        for ($i = 0; $i < 1; $i++) {
            $data[] = array("1_a_" . $i, "1_b_" . $i, "1_c_" . $i);
        }

        $service->addSheet(0, "test Sheet_0", $data);

        $data[] = array("id2", "name2");
        for ($i = 0; $i < 1; $i++) {
            $data[] = array("2_a_" . $i, "2_b_" . $i);
        }

        $service->addSheet(1, "test Sheet_1", $data);

        $objPHPExcel = $service->getPHPExcel();

        //add custom sheet
        $objPHPExcel->createSheet(2);
        $objPHPExcel->setActiveSheetIndex(2);
        $objPHPExcel->getActiveSheet()->setTitle("test");

        $file = $service->getFile();

        $objPHPExcel2 = new Xlsx();
        $objPHPExcel2->load($file);

        $worksheetNames = $objPHPExcel2->listWorksheetNames($file);
        static::assertCount(3, $worksheetNames);

    }


    /**
     * @throws SpreadsheetException
     * @throws WriterException
     * @throws ReaderException
     */
    public function testBenchmarkService(): void
    {
        $service = new BenchmarkService(); //->get('bassim_big_xlsx.benchmark.service');
        $service->create();

        $data = array();
        for ($i = 0; $i < $this->rowCount; $i++) {
            $data[] = array("1_a_" . $i, "1_b_" . $i, "1_c_" . $i);
        }

        $service->addSheet(0, "test Sheet_0", array("id", "name"), $data);
        $data = array();
        for ($i = 0; $i < $this->rowCount; $i++) {
            $data[] = array("2_a_" . $i, "2_b_" . $i);
        }

        $service->addSheet(1, "test Sheet_1", array("id2", "name2"), $data);
        $file = $service->get();

        $reader = new Xlsx();
        $reader->load($file);
        $worksheetNames = $reader->listWorksheetNames($file);
        unlink($file);
        static::assertCount(2, $worksheetNames);
    }
}
