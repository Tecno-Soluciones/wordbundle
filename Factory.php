<?php

namespace Tecno\WordBundle;

use PhpOffice\PhpWord\PhpWord;

/**
 * Factory for PHPWord objects, StreamedResponse, and PHPWord_Writer_IWriter.
 *
 * @package Tecno\WordBundle
 */
class Factory
{
    private $phpWordIO;

    public function __construct($phpWordIO = 'IOFactory::createWriter')
    {
        $this->phpWordIO = $phpWordIO;
    }
    /**
     * Creates an empty PHPWord Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return PhpWord
     */
    public function createPHPWordObject($filename =  null)
    {
        if (null == $filename) {
            $phpWordObject = new PhpWord();

            return $phpWordObject;
        }

        return call_user_func(array($this->phpWordIO, 'load'), $filename);
    }
}
