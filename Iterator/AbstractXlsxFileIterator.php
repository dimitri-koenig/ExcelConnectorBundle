<?php

namespace Pim\Bundle\ExcelConnectorBundle\Iterator;

use CallbackFilterIterator;
use Iterator;
use Symfony\Component\DependencyInjection\ContainerAwareInterface;
use Symfony\Component\DependencyInjection\ContainerInterface;
use Symfony\Component\OptionsResolver\OptionsResolverInterface;

/**
 * XLSX File iterator
 *
 * @author    Antoine Guigan <antoine@akeneo.com>
 * @copyright 2013 Akeneo SAS (http://www.akeneo.com)
 * @license   http://opensource.org/licenses/osl-3.0.php  Open Software License (OSL 3.0)
 */
abstract class AbstractXlsxFileIterator extends AbstractFileIterator implements ContainerAwareInterface,
    InitializableIteratorInterface
{
    /**
     * @var ContainerInterface
     */
    protected $container;

    /**
     * @var XlsxFileIterator
     */
    protected $xls;

    /**
     * @var Iterator
     */
    protected $worksheetIterator;

    /**
     * @var Iterator
     */
    protected $valuesIterator;

    /**
     * Constructor
     *
     * @param string $filePath
     * @param array  $options
     */
    public function __construct($filePath, array $options = array())
    {
        parent::__construct($filePath, $options);

        $reader = new \PHPExcel_Reader_Excel2007();
        $this->xls = $reader->load($filePath);
    }

    public function initialize()
    {
        $this->worksheetIterator = new CallbackFilterIterator(
            $this->xls->getWorksheetIterator(),
            function ($worksheet) {
                return $this->isIncludedWorksheet($worksheet);
            }
        );
        $this->rewind();
    }

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
        return sprintf('%s/%s', $this->worksheetIterator->key(), $this->valuesIterator->key());
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
        $this->worksheetIterator->rewind();
        if ($this->worksheetIterator->valid()) {
            $this->initializeValuesIterator();
        } else {
            $this->valuesIterator = null;
        }
    }

    /**
     * Returns the associate PHPExcel object
     *
     * @return \PHPExcel
     */
    public function getExcelObject()
    {
        return $this->xls;
    }

    /**
     * @inheritdoc
     */
    public function valid()
    {
        return $this->worksheetIterator->valid() && $this->valuesIterator && $this->valuesIterator->valid();
    }

    /**
     * Initializes the current record
     *
     * @return type
     */
    protected function initializeValuesIterator()
    {
        $this->valuesIterator = $this->createValuesIterator($this->worksheetIterator->current());
        if (!$this->valuesIterator->valid()) {
            $this->valuesIterator = null;
            $this->worksheetIterator->next();
            if ($this->worksheetIterator->valid()) {
                $this->initializeValuesIterator();
            }
        }
    }

    /**
     * {@inheritdoc}
     */
    public function setContainer(ContainerInterface $container = null)
    {
        $this->container = $container;
    }

    /**
     * Returns true if the worksheet should be included
     *
     * @param \PHPExcel_Worksheet $worksheet
     *
     * @return boolean
     */
    protected function isIncludedWorksheet(\PHPExcel_Worksheet $worksheet)
    {
        $title = $worksheet->getTitle();

        if (isset($this->options['include_worksheets'])) {
            $included = false;
            foreach ($this->options['include_worksheets'] as $regexp) {
                if (preg_match($regexp, $title)) {
                    $included = true;
                    break;
                }
            }

            if (!$included) {
                return false;
            }
        }

        foreach ($this->options['exclude_worksheets'] as $regexp) {
            if (preg_match($regexp, $title)) {
                return false;
            }
        }

        return true;
    }

    /**
     * {@inheritdoc}
     */
    protected function setDefaultOptions(OptionsResolverInterface $resolver)
    {
        $resolver->setDefaults(
            array(
                'exclude_worksheets'  => array()
            )
        );
        $resolver->setOptional(array('include_worksheets'));
    }

    /**
     * Returns the Excel Helper service
     *
     * @return \Pim\Bundle\ExcelConnectorBundle\Excel\ExcelHelper
     */
    protected function getExcelHelper()
    {
        return $this->container->get('pim_excel_connector.excel.helper');
    }

    /**
     * @param \PHPExcel_Worksheet $worksheet
     *
     * @return Iterator
     */
    abstract protected function createValuesIterator(\PHPExcel_Worksheet $worksheet);
}