<?php

declare(strict_types=1);

namespace App\Cli;

use Box\Spout\Common\Entity\Cell;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;

class ReadExcelFileBoxSpoutWayCommand extends Command
{
    protected function configure()
    {
        $this->setName('nidup:excel:read-file')
            ->setDescription('Read an excel file (with box/spout)');
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
        $path = 'data/movies-100.xlsx';
        # open the file
        $reader = ReaderEntityFactory::createXLSXReader();
        $reader->open($path);
        # read each cell of each row of each sheet
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $cells = $row->getCells();
                /**@var Cell $cell */
                foreach ($cells as $cell) {
                    var_dump($cell->getValue());
                }
            }
        }
        $reader->close();
        $output->writeln("I read the Excel File ".$path);

        return 0;
    }
}
