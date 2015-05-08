<?php namespace App\Console\Commands\PhpExcel;

use Illuminate\Console\Command;

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
            $this->info('  Worksheet: ' . $worksheet->getTitle());

            $drawings = [];
            /** @var \PHPExcel_Worksheet_BaseDrawing[] */
            $drawingCollection = $worksheet->getDrawingCollection();
            foreach ($drawingCollection as $drawing) {
                $this->info('  Drawing: ' . $drawing->getCoordinates());
                $drawings[$drawing->getCoordinates()] = $drawing;
            }

            foreach ($worksheet->getRowIterator() as $row) {
                $this->info('    Row number: ' . $row->getRowIndex());
                foreach ($row->getCellIterator() as $cell) {
                    if ($cell) {
                        if (isset($drawings[$cell->getCoordinate()])) {
                            $this->info('      Cell: ' . $cell->getCoordinate() . '; Value: drawing');
                            $drawing = $drawings[$cell->getCoordinate()];
                            $drawingPath = public_path() . '/images/catalog/';
                            if (!file_exists($drawingPath)) {
                                mkdir($drawingPath, 0777, true);
                            }
                            $drawingIndexedFilename = $drawingPath . $drawing->getIndexedFilename();
                            $this->info('      DrawingIndexedFilename: ' . $drawingIndexedFilename);
                            if ($drawing instanceof \PHPExcel_Worksheet_Drawing) {
                                $this->info('      DrawingPath: ' . $drawing->getPath());
                                copy($drawing->getPath(), $drawingIndexedFilename);
                            } elseif ($drawing instanceof \PHPExcel_Worksheet_MemoryDrawing) {
                                // we have a memory drawing (case xls)
                                $image = $drawing->getImageResource();
                                // save image to disk
                                $renderingFunction = $drawing->getRenderingFunction();
                                switch ($renderingFunction) {
                                    case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG:
                                        imagejpeg($image, $drawingIndexedFilename);
                                        break;

                                    case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_GIF:
                                        imagegif($image, $drawingIndexedFilename);
                                        break;

                                    case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_PNG:
                                    case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_DEFAULT:
                                        imagepng($image, $drawingIndexedFilename);
                                        break;
                                }
                            }
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
