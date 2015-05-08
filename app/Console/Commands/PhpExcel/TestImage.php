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
        $this->info(storage_path());
    }
}
