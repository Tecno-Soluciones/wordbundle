<?php

namespace  Tecno\WordBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;

class FakeController extends Controller
{
    public function streamAction()
    {
        // create an empty object
        $phpExcelObject = $this->createXSLObject();
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Excel5');
        // create the response
        $response = $this->get('phpword')->createStreamedResponse($writer);
        // adding headers
        $response->headers->set('Content-Type', 'text/vnd.ms-word; charset=utf-8');
        $response->headers->set('Content-Disposition', 'attachment;filename=stream-file.dos');
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');

        return $response;
    }

    public function storeAction()
    {
        // create an empty object
        $phpExcelObject = $this->createDOCObject();
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Word2007');
        $filename = tempnam(sys_get_temp_dir(), 'xls-') . '.xls';
        // create filename
        $writer->save($filename);

        return new Response($filename, 201);
    }

    public function readAndSaveAction()
    {
        $filename = $this->container->getParameter('xls_fixture_absolute_path');
        // create an object from a filename
        $phpExcelObject = $this->createDOCObject($filename);
        // create the writer
        $writer = $this->get('phpword')->createWriter($phpExcelObject, 'Word2007');
        $filename = tempnam(sys_get_temp_dir(), 'doc-') . '.doc';
        // create filename
        $writer->save($filename);

        return new Response($filename, 201);
    }

    /**
     * utility class
     * @return mixed
     */
    private function createDOCObject()
    {
        $phpExcelObject = $this->get('phpword')->createPHPExcelObject();

		$section = $phpExcelObject->createSection();
		
        $section->addText('Hello world!');

		// You can directly style your text by giving the addText function an array:
		$section->addText('Hello world! I am formatted.', array('name'=>'Tahoma', 'size'=>16, 'bold'=>true));

		// If you often need the same style again you can create a user defined style to the word document
		// and give the addText function the name of the style:
		$PHPWord->addFontStyle('myOwnStyle', array('name'=>'Verdana', 'size'=>14, 'color'=>'1B2232'));
		$section->addText('Hello world! I am formatted by a user defined style', 'myOwnStyle');

		// You can also putthe appended element to local object an call functions like this:
		$myTextElement = $section->addText('Hello World!');
		$myTextElement->setBold();
		$myTextElement->setName('Verdana');
		$myTextElement->setSize(22);

        return $phpExcelObject;
    }
}
