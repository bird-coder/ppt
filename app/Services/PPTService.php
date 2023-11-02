<?php

namespace App\Services;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\PowerPoint2007;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;

class PPTService
{
    private $filepath;

    /**
     * @var PhpPresentation
     */
    private $ppt;

    /**
     * @var PhpPresentation
     */
    private $tpl;

    private $savePath;

    private $configPath;

    public function __construct($filepath) {
        $this->filepath = $filepath;
        $this->ppt = new PhpPresentation();
        if (!file_exists(storage_path('ppt/output'))) {
            mkdir(storage_path('ppt/output'), 0777, true);
        }
        $this->savePath = storage_path('ppt/output');
        $this->configPath = resource_path('template/ppt'). '/config.json';
    }

    public function init() {
        $reader = new PowerPoint2007();
        if ($reader->canRead($this->filepath)) {
            $this->tpl = $reader->load($this->filepath);
        }
    }

    public function copy($indexes = []) {
        if ($this->tpl) {
            $this->ppt->removeSlideByIndex();
            
            foreach ($indexes as $index) {
                $slide = $this->tpl->setActiveSlideIndex($index);
                $this->ppt->addSlide($slide->copy());
            }
            
        }
    }

    // public function addContent() {
    //     if ($this->tpl) {
    //         $slide = $this->tpl->setActiveSlideIndex(8);
    //         //获取所有形状格式内容
    //         foreach ($slide->getShapeCollection() as $one => $shape) {
    //             try {
    //                 if ($shape instanceof RichText) {
    //                     $paragraphs = $shape->getParagraphs();
    //                     foreach ($paragraphs as $two => $paragraph) {
    //                         foreach ($paragraph->getRichTextElements() as $three => $richText) {
    //                             $text = $richText->getText();

    //                             switch ((string)$one.$two) {
    //                                 case 'value':
    //                                     # code...
    //                                     break;
                                    
    //                                 default:
    //                                     # code...
    //                                     break;
    //                             }
    //                         }
    //                     }
    //                 } elseif ($shape instanceof Gd) {

    //                 }
                    
    //             } catch (\Exception $e) {
    //                 var_dump($e);
    //             }
    //         }
    //         // $this->save();
    //     }
    // }

    public function addContent() {
        $path = storage_path('ppt');
        $dirs = $this->loadDataFile($path);
        foreach ($dirs as $dir) {
            if (strpos($dir['json'], 'CN201810934292.4') === false) {
                continue;
            }
            $this->addSlide($dir);
        }
        $this->save();
    }

    private function addSlide($dir) {
        $slide = $this->ppt->createSlide();

        $data = json_decode(file_get_contents($dir['json']), true);
        $configs = json_decode(file_get_contents($this->configPath), true);
        foreach ($configs as $key=>$config) {
            if ($key == 'title') {
                $shape = $this->initRichText($slide, $config);
                $paragraph = $shape->getActiveParagraph();
                $paragraph->createTextRun($data['title']);
            } elseif ($key == 'subTitle') {
                $shape = $this->initRichText($slide, $config);
                $paragraph = $shape->getActiveParagraph();
                $paragraph->createTextRun($data['number']);
            } elseif ($key == 'body') {
                $shape = $this->initRichText($slide, $config);
                $paragraph = $shape->getActiveParagraph();
                $paragraph->createTextRun('Applicants: ');
                $paragraph->createTextRun($data['applicants'])
                    ->getFont()->setName($config['text']['name']);
                $paragraph->createBreak();
                $paragraph->createTextRun('Inventors: ');
                $paragraph->createTextRun($data['inventors'])
                    ->getFont()->setName($config['text']['name']);
                $paragraph->createBreak();
                $paragraph->createTextRun('Application Time: ' . $data['application_time']);
                $paragraph->createBreak();
                $paragraph->createTextRun('Legal status: ' . $data['legal_status']);
                $paragraph->createBreak();
                $paragraph->createTextRun('Patent type: ' . $data['patent_type']);
                $paragraph->createBreak();
                $paragraph->createTextRun('The current obligee: ');
                $paragraph->createTextRun($data['obligee'])
                    ->getFont()->setName($config['text']['name']);
            } elseif ($key == 'desc') {
                $shape = $this->initRichText($slide, $config);
                $paragraph = $shape->getActiveParagraph();
                $paragraph->createTextRun('Abstract:');
                $paragraph->createBreak();
                $paragraph->createTextRun($data['abstract'])
                    ->getFont()->setName($config['text']['name']);
            } elseif ($key == 'img') {
                $this->addImage($slide, $config, $dir['image']);
            }
        }

    }

    private function loadDataFile($path = '') {
        $od = @opendir($path);
        $dirs = [];
        while (($file = readdir($od)) !== false) {
            if ($file == '.' || $file == '..' || $file == 'output') {
                continue;
            }
            $subDir = $path . '/' . $file;
            if (is_dir($subDir)) {
                $files = scandir($subDir);
                $tmp = ['json' => '', 'image' => ''];
                foreach ($files as $f) {
                    if (strpos($f, 'json') !== false) {
                        $tmp['json'] = $subDir . '/' . $f;
                    } elseif (strpos($f, 'png') !== false) {
                        $tmp['image'] = $subDir . '/' . $f;
                    } elseif (strpos($f, 'pdf') !== false) {
                        copy($subDir . '/' . $f, $this->savePath . '/' . $f);
                    }
                }
                $dirs[] = $tmp;
            }
        }
        return $dirs;
    }

    private function initRichText(Slide $slide, $config) {
        $shape = $slide->createRichTextShape();
        $shape->setWidthAndHeight($config['p']['width'], $config['p']['height']);
        $shape->setOffsetX($config['p']['offsetX']);
        $shape->setOffsetY($config['p']['offsetY']);

        $paragraph = $shape->getActiveParagraph();
        $paragraph->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_LEFT)
            ->setVertical(Alignment::VERTICAL_TOP)
            ->setTextDirection(Alignment::TEXT_DIRECTION_HORIZONTAL);

        if ($config['p']['lineSpacing'] > 0) $paragraph->setLineSpacing($config['p']['lineSpacing']);

        $paragraph->getFont()
            ->setName($config['text']['name_en'])
            ->setSize($config['text']['size'])
            ->setColor(new Color($config['text']['color']))
            ->setBold($config['text']['bold']);

        return $shape;
    }

    private function addImage(Slide $slide, $config, $imgpath) {
        $shape = $slide->createDrawingShape();
        $shape->setPath($imgpath)
            ->setResizeProportional(true)
            ->setWidth($config['width'])
            ->setOffsetX($config['offsetX'])
            ->setOffsetY($config['offsetY']);

        if ($shape->getHeight() > $config['height']) {
            $shape->setHeight($config['height']);
        }
    }

    private function save() {
        if ($this->ppt) {
            $writer = IOFactory::createWriter($this->ppt, 'PowerPoint2007');
            $writer->save($this->savePath.'/output.ppt');
        }
    }

    public function download() {
        header('Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation');
        header('Content-Disposition: attachment; filename=');
        $writer = IOFactory::createWriter($this->ppt, 'PowerPoint2007');
        $writer->save('php://output');
        unlink($this->savePath.'/output.ppt');
        exit;
    }
}