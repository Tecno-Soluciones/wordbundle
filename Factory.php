<?php

namespace Tecno-Soluciones\WordBundle;

use Symfony\Component\HttpFoundation\StreamedResponse;

/**
 * Factory for PHPWord objects, StreamedResponse, and PHPWord_Writer_IWriter.
 *
 * @package Tecno\WordBundle
 */
class Factory
{
    private $phpWordIO;

    public function __construct($phpWordIO = '\PHPWord_IOFactory')
    {
        $this->phpWordIO = $phpWordIO;
    }
    /**
     * Creates an empty PHPWord Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return \PHPWord
     */
    public function createPHPWordObject($filename =  null)
    {
        if (null == $filename) {
            $phpWordObject = new \PHPWord();

            return $phpWordObject;
        }

        return call_user_func(array($this->phpWordIO, 'load'), $filename);
    }

    /**
     * Create a writer given the PHPWordObject and the type,
     *   the type coul be one of PHPWord_IOFactory::$_autoResolveClasses
     *
     * @param \PHPWord $phpWordObject
     * @param string    $type
     *
     *
     * @return \PHPWord_Writer_IWriter
     */
    public function createWriter(\PHPWord $phpWordObject, $type = 'Word5')
    {
        return call_user_func(array($this->phpWordIO, 'createWriter'), $phpWordObject, $type);
    }

    /**
     * Stream the file as Response.
     *
     * @param \PHPWord_Writer_IWriter $writer
     * @param int                      $status
     * @param array                    $headers
     *
     * @return StreamedResponse
     */
    public function createStreamedResponse(\PHPWord_Writer_IWriter $writer, $status = 200, $headers = array())
    {
        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }
}