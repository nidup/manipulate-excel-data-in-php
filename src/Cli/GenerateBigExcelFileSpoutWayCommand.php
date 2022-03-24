<?php

declare(strict_types=1);

namespace App\Cli;

use Box\Spout\Common\Entity\Cell;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Stopwatch\Stopwatch;

class GenerateBigExcelFileSpoutWayCommand extends Command
{
    /** @var Stopwatch */
    private $stopwatch;

    public function __construct(Stopwatch $stopwatch)
    {
        parent::__construct();
        $this->stopwatch = $stopwatch;
    }

    protected function configure()
    {
        $this->setName('nidup:excel:generate-big-file')
            ->setDescription('Generate a 1M lines excel file (with box/spout)');
    }

    /**
     * @param InputInterface $input
     * @param OutputInterface $output
     * @return int
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $path = 'data/movies-100.xlsx';
        # open the file
        $reader = ReaderEntityFactory::createXLSXReader();
        $reader->open($path);
        # read each cell of each row of each sheet
        $rowsData = [];
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $rowsData[]= $row->toArray();
            }
        }
        $reader->close();

        // generate 1M lines
        $pathBig = 'data/movies-1000000.xlsx';
        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($pathBig);
        for ($idx = 0; $idx < 10000; $idx++) {
            foreach ($rowsData as $row) {
                $rowFromValues = WriterEntityFactory::createRowFromArray($row);
                $writer->addRow($rowFromValues);
            }
        }
        $writer->close();
        $output->writeln("I generate 1M rows in the Excel File ".$pathBig);

        return 0;
    }
}
