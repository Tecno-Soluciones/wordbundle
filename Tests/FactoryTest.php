<?php

namespace Tecno\WordBundle\Tests;

use Tecno\WordBundle\Factory;

class FactoryTest extends \PHPUnit_Framework_TestCase
{
    public function testCreate()
    {
        $factory =  new Factory();
        $this->assertInstanceOf('\PHPWord', $factory->createPHPWordObject());
    }

    public function testCreateStreamedResponse()
    {
        $writer = $this->getMock('\PHPWord_Writer_IWriter');
        $writer->expects($this->once())
            ->method('save')
            ->with('php://output');

        $factory =  new Factory();
        $factory->createStreamedResponse($writer)->sendContent();
    }
}
