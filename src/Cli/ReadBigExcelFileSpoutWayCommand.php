<?php

declare(strict_types=1);

namespace App\Cli;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use League\Csv\UnavailableStream;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Stopwatch\Stopwatch;

class ReadBigExcelFileSpoutWayCommand extends Command
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
        $this->setName('nidup:excel:read-big-file')
            ->setDescription('Read a big excel file and measure time and memory (with box/spout)');
    }

    /**
     * @param InputInterface $input
     * @param OutputInterface $output
     * @return int
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    protected function execute(InputInterface $input, OutputInterface $output)
    {
        try {
            $section = 'read_excel_file';
            $this->stopwatch->start($section);
            $path = 'data/movies-1000000.xlsx';
            $reader = ReaderEntityFactory::createXLSXReader();
            $reader->open($path);
            # read each cell of each row of each sheet
            foreach ($reader->getSheetIterator() as $sheet) {
                foreach ($sheet->getRowIterator() as $row) {
                    $cells = $row->getCells();
                    foreach ($cells as $cell) {
                        // we do nothing, but we want to ensure we browse each row
                    }
                }
            }
            $this->stopwatch->stop($section);
            $output->writeln("I read 1.000.000 rows from the Excel File ".$path);
            $output->writeln((string) $this->stopwatch->getEvent($section));
            return 0;
        } catch (UnavailableStream $exception) {
            $output->writeln("File ".$path." does not exist, it has to be generated with the command 'nidup:excel-spout:generate-big-excel-file'");
            return -1;
        }
    }
}
