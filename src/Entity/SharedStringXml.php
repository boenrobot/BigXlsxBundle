<?php

declare(strict_types=1);

namespace Bassim\BigXlsxBundle\Entity;

class SharedStringXml
{
    private $string;
    private $lineCount;
    private $position = -1;
    private $strings = array();
    private $filePointer;
    private $sharedStringsFile;

    public function __construct()
    {
        $this->sharedStringsFile = tempnam(sys_get_temp_dir(), 'sss');
        $this->filePointer = fopen($this->sharedStringsFile, 'wb');
    }


    public function addString(string $string): int
    {
//        $pos = array_search($string, $this->strings);
//        if ($pos !== false) {
//            return null;
//        }

        $this->position++;
        $this->strings[] = $string;
        $string = htmlspecialchars($string, ENT_NOQUOTES | ENT_XML1 | ENT_IGNORE, 'UTF-8');

        $this->write("<si><t>" . $string . "</t></si>");
        return $this->position;
    }

    public function getFile(): string
    {
        $this->flush();
        $this->prependHeader();
        $this->appendFooter();
        return $this->sharedStringsFile;
    }

    private function prependHeader(): void
    {
        /** @noinspection HttpUrlsUsage */
        $string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="' . count($this->strings) . '">
';

        $context = stream_context_create();
        $filePointer = fopen($this->sharedStringsFile, 'rb', true, $context);
        $tmpName = tempnam(sys_get_temp_dir(), 'PHT');
        file_put_contents($tmpName, $string);
        file_put_contents($tmpName, $filePointer, FILE_APPEND);
        fclose($filePointer);
        unlink($this->sharedStringsFile);
        rename($tmpName, $this->sharedStringsFile);
    }

    private function appendFooter(): void
    {
        $context = stream_context_create();
        $filePointer = fopen($this->sharedStringsFile, 'ab', true, $context);
        fwrite($filePointer, '</sst>');
        fclose($filePointer);
    }

    private function write($string): void
    {
        $this->lineCount++;
        $this->string .= $string;
    }

    private function flush(): void
    {
        fwrite($this->filePointer, $this->string);
        fclose($this->filePointer);
    }
}
