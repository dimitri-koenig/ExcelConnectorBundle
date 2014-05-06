<?php

namespace spec\Pim\Bundle\ExcelConnectorBundle\Excel\Reader;

use PhpSpec\ObjectBehavior;
use Prophecy\Argument;

class ContentCacheSpec extends ObjectBehavior
{
    function let()
    {
        $this->beConstructedWith(__DIR__ . '/../../fixtures/sharedStrings.xml');
    }

    function it_is_initializable()
    {
        $this->shouldHaveType('Pim\Bundle\ExcelConnectorBundle\Excel\Reader\ContentCache');
    }

    function it_returns_shared_strings()
    {
        $this->get(0)->shouldReturn('value1');
        $this->get(2)->shouldReturn('  ');
        $this->get(4)->shouldReturn('value3');
        $this->get(5)->shouldReturn('value4');
    }
}
