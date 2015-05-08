<?php namespace App\Console\Commands\PhpExcel;

use Illuminate\Console\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputOption;

class TestImage extends Command
{
    /**
     * The console command name.
     *
     * @var string
     */
    protected $name = 'phpexcel:test-image';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Extract image from .xls file with PhpExcel';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function fire()
    {
        $path = storage_path() . '/app/phpexcel/template.xlsx';
        $this->info('Path: ' . $path);
        $phpExcel = \PHPExcel_IOFactory::load($path);
        foreach ($phpExcel->getWorksheetIterator() as $worksheet) {
            $drawings = [];
            /** @var \PHPExcel_Worksheet_BaseDrawing[] */
            $drawingCollection = $worksheet->getDrawingCollection();
            foreach ($drawingCollection as $drawing) {
                $this->info('  Drawing: ' . $drawing->getCoordinates());
                $drawings[$drawing->getCoordinates()] = $drawing;
            }
            $this->info('  Worksheet: ' . $worksheet->getTitle());
            foreach ($worksheet->getRowIterator() as $row) {
                $this->info('    Row number: ' . $row->getRowIndex());
                foreach ($row->getCellIterator() as $cell) {
                    if ($cell) {
                        if (isset($drawings[$cell->getCoordinate()])) {
                            $this->info('      Cell: ' . $cell->getCoordinate() . '; Value: drawing');
                        } else {
                            $this->info('      Cell: ' . $cell->getCoordinate() . '; Value: ' . $cell->getCalculatedValue());
                        }
                    } else {
                        $this->error('!!!!!!Cell: ' . $cell->getCoordinate() . '; Value: null');
                    }
                }
            }
        }
    }
}
