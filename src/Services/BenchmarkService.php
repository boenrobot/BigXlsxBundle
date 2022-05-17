<?php

declare(strict_types=1);

namespace Bassim\BigXlsxBundle\Services;

use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class BenchmarkService
{

    /** @var $objPHPExcel Spreadsheet */
    private $objPHPExcel;
    private $columnName;
    private $sheets = [];

    public function __construct()
    {
    }


    public function create(): void
    {
        $this->columnName = null;

        $this->objPHPExcel = new Spreadsheet();

    }

    /**
     * @param int $sheetNumber
     * @param string $name
     * @param iterable $columns
     * @param iterable $data
     * @return void
     * @throws SpreadsheetException
     */
    public function addSheet(int $sheetNumber, string $name, iterable $columns, iterable $data): void
    {
        if ($sheetNumber > 0) {
            $this->objPHPExcel->createSheet($sheetNumber);
        }

        $this->objPHPExcel->setActiveSheetIndex($sheetNumber);
        $this->objPHPExcel->getActiveSheet()->setTitle($name);

        $this->sheets[$sheetNumber] = array_merge(array($columns), $data);

        //set headers
        foreach ($columns as $row) {
            if (is_null($this->columnName)) {
                $this->columnName = 'a';
            }
            $cellName = strtoupper($this->columnName++ . "1");
            $this->objPHPExcel->getActiveSheet()->setCellValue($cellName, $row);
        }

        $this->columnName = null;

        //set data
        foreach ($data as $rowIndex => $row) {
            foreach ($row as $cell) {
                if (is_null($this->columnName)) {
                    $this->columnName = 'a';
                }
                $cellName = strtoupper($this->columnName++ . "" . ($rowIndex + 2));
                $this->objPHPExcel->getActiveSheet()->setCellValue($cellName, $cell);
            }
            $this->columnName = null;
        }
    }

    /**
     * @return string
     * @throws Exception
     */
    public function get(): string
    {
        // Save Excel 2007 file
        $objWriter = new Xlsx($this->objPHPExcel);
        $filename = tempnam(sys_get_temp_dir(), 'BCH');

        $objWriter->save($filename);

        return $filename;
    }

    /**
     * @return array
     */
    public function getSheets(): array
    {
        return $this->sheets;
    }
}
