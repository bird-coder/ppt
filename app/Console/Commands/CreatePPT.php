<?php
/*
 * @Description: 
 * @Author: yuanshisan
 * @Date: 2023-10-24 11:40:46
 * @LastEditTime: 2023-11-04 16:56:01
 * @LastEditors: yuanshisan
 */

namespace App\Console\Commands;

use App\Services\PPTService;
use App\Services\SpiderService;
use Illuminate\Console\Command;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\PowerPoint2007;

class CreatePPT extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'ppt:create {--filepath=}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = '根据ppt模板及爬取的内容, 生成ppt';

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
     * @return int
     */
    public function handle()
    {
        $filepath = $this->option('filepath');
        
        // $this->getData();

        $this->createPPT();
    }

    private function createPPT() {
        $ppt = new PPTService('Pet hair dryer');
        $ppt->init();
        $ppt->addTable();
        $ppt->addChart();
        $ppt->addChart();
        $ppt->addContent();
        $ppt->save();
    }

    private function getData() {
        $spider = new SpiderService();
        $patents = $spider->patents();
        var_dump($patents);
        foreach ($patents as $item) {
            $patent_id = $item['patent_id'];
            if (file_exists(storage_path('/ppt/'. $item['pn']))) {
                continue;
            }
            $patent = $spider->getPatent($patent_id);
            $patent['pdf'] = $spider->getPDF($patent_id);
            if (empty($patent) || empty($patent['pdf'])) {
                echo 'token 过期';
                break;
            }
            $spider->saveData($patent);
            sleep(1);
        }
    }
}
