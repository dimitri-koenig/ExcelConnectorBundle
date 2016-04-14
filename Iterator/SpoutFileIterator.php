<?php

namespace Pim\Bundle\ExcelConnectorBundle\Iterator;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Iterator;
use Symfony\Component\DependencyInjection\ContainerAwareInterface;
use Symfony\Component\DependencyInjection\ContainerInterface;
use Symfony\Component\OptionsResolver\OptionsResolver;

/**
 * Spout File iterator
 *
 * @author    JM Leroux <jean-marie.leroux@akeneo.com>
 * @copyright 2013 Akeneo SAS (http://www.akeneo.com)
 * @license   http://opensource.org/licenses/osl-3.0.php  Open Software License (OSL 3.0)
 */
class SpoutFileIterator extends AbstractFileIterator implements ContainerAwareInterface
{
    /** @var ContainerInterface */
    protected $container;

    /** @var Iterator */
    protected $worksheetIterator;

    /** @var Iterator */
    protected $valuesIterator;

    /** @var ReaderInterface */
    private $spout;

    /** @var string[] */
    protected $labels;
    
    /**
     * {@inheritdoc}
     */
    public function current()
    {
        return $this->valuesIterator->current();
    }

    /**
     * {@inheritdoc}
     */
    public function key()
    {
        return sprintf('%s/%s', $this->worksheetIterator->current(), $this->valuesIterator->key());
    }

    /**
     * {@inheritdoc}
     */
    public function next()
    {
        $this->valuesIterator->next();

        if (!$this->valuesIterator->valid()) {
            $this->worksheetIterator->next();
            if ($this->worksheetIterator->valid()) {
                $this->initializeValuesIterator();
            }
        }
    }

    /**
     * {@inheritdoc}
     */
    public function rewind()
    {
        $this->worksheetIterator = $this->getWorksheetIterator();
        $this->worksheetIterator->rewind();
        if ($this->worksheetIterator->valid()) {
            $this->initializeValuesIterator();
        } else {
            $this->valuesIterator = null;
        }
    }

    /**
     * Returns the associated Excel object
     *
     * @return ReaderInterface
     */
    public function getExcelObject()
    {
        if (!$this->spout) {
            $this->spout = ReaderFactory::create(Type::XLSX);
            $this->spout->open($this->filePath);
        }

        return $this->spout;
    }

    /**
     * @inheritdoc
     */
    public function valid()
    {
        return $this->worksheetIterator->valid() && $this->valuesIterator && $this->valuesIterator->valid();
    }

    /**
     * {@inheritdoc}
     */
    public function setContainer(ContainerInterface $container = null)
    {
        $this->container = $container;
    }

    private function getWorksheetIterator()
    {
        // TODO imlements CallbackFilterIterator to filter worksheets
        return $this->getExcelObject()->getSheetIterator();
    }

    /**
     * Initializes the current worksheet
     */
    protected function initializeValuesIterator()
    {
        $this->valuesIterator = $this->createValuesIterator();

        if (!$this->valuesIterator->valid()) {
            $this->valuesIterator = null;
            $this->worksheetIterator->next();
            if ($this->worksheetIterator->valid()) {
                $this->initializeValuesIterator();
            }
        }
    }

    /**
     * Returns true if the worksheet should be read
     *
     * @param string $title
     *
     * @return boolean
     */
    protected function isReadableWorksheet($title)
    {
        return $this->isIncludedWorksheet($title) && !$this->isExcludedWorksheet($title);
    }

    /**
     * Returns true if the worksheet should be indluded
     *
     * @param string $title The title of the worksheet
     *
     * @return boolean
     */
    protected function isIncludedWorksheet($title)
    {
        if (!isset($this->options['include_worksheets'])) {
            return true;
        }

        foreach ($this->options['include_worksheets'] as $regexp) {
            if (preg_match($regexp, $title)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Returns true if the worksheet should be excluded
     *
     * @param string $title The title of the worksheet
     *
     * @return boolean
     */
    protected function isExcludedWorksheet($title)
    {
        foreach ($this->options['exclude_worksheets'] as $regexp) {
            if (preg_match($regexp, $title)) {
                return true;
            }
        }

        return false;
    }

    /**
     * {@inheritdoc}
     */
    protected function setDefaultOptions(OptionsResolver $resolver)
    {
        $resolver->setDefaults(
            array(
                'exclude_worksheets' => array(),
                'parser_options' => array()
            )
        );
        $resolver->setDefined(array('include_worksheets'));
    }

    /**
     * Creates the value iterator
     *
     * @return Iterator
     */
    protected function createValuesIterator()
    {
        // TODO manage $this->options['parser_options']
        $sheet = $this->getWorksheetIterator()->current();

        $iterator = $sheet->getRowIterator();
        $iterator->rewind();
        $this->options['label_row'] = 1;
        $this->options['data_row'] = 2;
        while ($iterator->valid() && ((int) $this->options['label_row'] > $iterator->key())) {
            $iterator->next();
        }
        
        $this->labels = $iterator->current();
        
        while ($iterator->valid() && ((int) $this->options['data_row'] > $iterator->key())) {
            $iterator->next();
        }

        return $iterator;
    }
}
