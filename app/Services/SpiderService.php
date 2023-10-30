<?php
/*
 * @Description: 
 * @Author: yuanshisan
 * @Date: 2023-10-24 16:06:44
 * @LastEditTime: 2023-10-25 22:14:25
 * @LastEditors: yuanshisan
 */

namespace App\Services;

use Illuminate\Support\Facades\Http;

class SpiderService
{
    private $searchHost = 'https://search-service.zhihuiya.com/core-search-api/search/';

    private $headers = [
        'Origin'        => 'https://analytics.zhihuiya.com/',
        'Referer'       => 'https://analytics.zhihuiya.com/',
        'X-Api-Version' => '2.0',
        'X-Patsnap-From'=> 'w-analytics-search-result',
    ];

    private $token = 'eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIzZDA2YjY5MTUyZmY0NDM1OWZiYzQ2NzI2ZmMwZjM3MyIsInByb2R1Y3QiOiJwcm8iLCJzZXNzaW9uIjoiZWI3OGU5ZTBjY2U5NDJlZWFkZDlmZWNhZGI0YmM4MTYiLCJjb25zb2xlX2FjY2VzcyI6ZmFsc2UsImlzcyI6InBhdHNuYXAiLCJjbGllbnRfaWQiOiJmNThiYmRmZGQ2MzU0OWRiYjY0ZmVkNGI4MTZjOGJmYyIsImF1dGhvcml0aWVzIjpbImV5bGYiLCIxeWV5IiwiMXc0bCIsImEwNnEiLCJleTJnIiwiZzIwbSIsIjdqNjYiLCIxam4xIiwiODAwMDIiLCI4MDIwMCIsImIwMDA4IiwiODAwMDciLCIxMWNlIiwiZXVscyIsIjF5dzEiLCJleXBwIiwiYjAyMDMiLCIxeWlmIiwiYjAwMDQiLCJiMDIwMSIsIjIwMzAxIiwiZTJ3diIsIjIwMTAxIiwiZTM5NiIsImdjMGsiLCIyMDUwNyIsIjIwOTAzIiwiZW4zcCIsIjF0ajkiLCJhMGtjIiwiZTk3aCIsIjFqcXMiLCIyMDEwOCIsIjIwNTA0IiwiMjA3MDIiLCJlMTNqIiwiMjA1MDEiLCJjb2c3IiwiZ3B5MiIsIjFwbjAiLCIxajJ5IiwiZTllZyIsImgwMDAwIiwiaDAwMDMiLCIxdzRyIiwiMXc0dSIsIjIwMTE1IiwiMTE4bSIsIjd1YzciLCIxeTF6IiwiMjA1MTgiLCIyMDUxNyIsIjIwNTE1IiwiMXBuaCIsIjIwNTEyIiwiMWNjeCIsImEwNWMiLCJnbjBrIiwiN2o2cSIsImd0bmUiLCIxY2d2IiwiMXBnNiIsIjF3MW0iLCIyMDIwMyIsIjIwNDAxIiwiMWNlaSIsIjIwMDA0IiwiZTI5OCIsIjc5bWUiLCI3cXEwIiwiMjAwMDEiLCIxcDMxIiwiMjA2MDUiLCIyMDgwMyIsIjgwMDA5IiwiZTlmdiIsIjIwMDA5IiwiN2o1OCIsImU5angiLCIyMDAwNyIsIjIwMjA1IiwiMjA2MDEiLCI3ajd4IiwiZzJrYSIsIjF3aGoiLCI3YWdmIiwiZDAwMDAiLCIxeWRrIiwiYzBjbSIsIjF3dHEiLCJnbW90IiwiMWoxYiIsIjdhc3AiLCIxeWhnIl0sInByb2R1Y3RzIjpbIjhjN2ZhZDk3MmEwZTQwMDNiNjMwOTU2MjY4NjQ0OTAzIiwiZWM0YmRmOWI3NmE5NGE4M2FkZTVjMDYyMmMzNDYwNjIiLCJwcm8iLCJyZXBvcnQiLCJmMjA3Y2E2MGVjMDY0ZWRiYmNiYzI2OTk3OTg1NDAwYSIsImMxMDgzMjM0OWY4NzRjOGZhYmQwZTljYTc0YTJkOTJlIl0sImZpcnN0bG9naW5fcHJvZHVjdCI6ImFuYWx5dGljcyIsInVzZXJfdHlwZSI6IlRSSUFMIiwidXNlcl9pZCI6IjBkNDE3YjlhMGYwZTQxN2Q4ZDg5OWM3MTlkOTY0MjUwIiwiZXhwIjoxNjk4MTU4NjgzLCJpYXQiOjE2OTgxNTY4ODMsImp0aSI6IjdmYWQ2MWFjLWNhMDgtNDUwZC1hODEyLTdjNWRiYzkyMjY1MyJ9.RuSQRJ2BSfN11H9CpGmv7cn3jDApkMHl_QL88XOxxAYhZ2Be-wzwENNSOAzY7eSpfA1Wl4vCosUoNi_xjCEgXmbzzZ3m9XtwQteoN49bj_kRGq_jimZEpP5F-w9LIX0Xu5r6kUFpTHTaImoHAlgEmVfpnozBpiDbHgJLcNC1Y_g';

    protected $legal_status = [
        1 => '',
        2 => 'Under review',
        3 => 'Authorized',
    ];
    protected $patent_types = [
        'U' => 'Utility model',
        'A' => 'Invention',
    ];

    public function __construct($headers = [], $token = '')
    {
        if (!empty($headers)) $this->headers = $headers;
        if (!empty($token)) $this->token = $token;
    }

    private function submit() {
        $url = $this->searchHost . 'submit';
        $data = array(
            'q' => '宠物 and (辅食 or 辅食机 or 料理机)'
        );
        $res = $this->post($url, $data);
        if (empty($res) || !$res['status']) {
            return false;
        }
        return true;
    }

    public function srpInit() {
        $url = $this->searchHost . 'filter/srp-init';
        $data = array(
            'isSpecialQuery' => false,
            'limit' => '100',
            'page' => '1',
            'originQuery' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'q' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'query' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'search_query' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'search_mode' => 'unset',
            'sort' => 'desc',
            'viewtype' => 'tablelist',
            '_type' => 'query'
        );
        $res = $this->post($url, $data);
        if (empty($res) || !$res['status']) {
            return false;
        }
        return true;
    }

    public function patents() {
        // if (!$this->submit()) {
        //     return false;
        // }
        // if (!$this->srpInit()) {
        //     return false;
        // }
        $url = $this->searchHost . 'srp/patents';
        $this->headers['X-Patsnap-From'] = 'w-analytics-search-result';
        $data = array(
            'efq' => "LEGAL_STATUS:(\"1\" OR \"2\" OR \"3\")",
            'history_id' => 'NEW',
            'job_id' => '',
            'limit' => 100,
            'page' => 1,
            'q' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'redirect' => '',
            'search_mode' => 'unset',
            'semantic_id' => '',
            'sn' => '',
            'sort' => 'desc',
            'special_query' => false,
            'use_async' => true,
            'view_type' => 'tablelist',
            'with_count' => true,
            'with_retry' => true,
            '_type' => 'query',
        );
        $res = $this->post($url, $data);
        if (empty($res) || !$res['status']) {
            return [];
        }
        $list = [];
        foreach ($res['data']['patent_data'] as $patent) {
            //命中
            $title = $patent['TITLE'];
            if (strpos($title, 'span') === false) {
                continue;
            } elseif (!preg_match('/(宠物|辅食).*(装置|机器|料理机|器)/', $title)) {
                continue;
            }
            $tmp = array(
                'pn' => $patent['PN'],
                'patent_id' => $patent['PATENT_ID'],
            );
            array_push($list, $tmp);
        }
        return $list;
    }

    public function getPatent($patent_id) {
        $url = $this->searchHost . 'patent/id/' . $patent_id . '/basic?highlight=true';
        $this->headers['X-Patsnap-From'] = 'w-analytics-search-view';
        $data = array(
            'efq' => "LEGAL_STATUS:(\"1\" OR \"2\" OR \"3\")",
            'limit' => 100,
            'page' => 1,
            'q' => '宠物 and (辅食 or 辅食机 or 料理机)',
            'query' => [],
            'rows' => 100,
            'search_mode' => 'unset',
            'selected' => [],
            'sort' => 'desc',
            'source_type' => 'search_result',
            '_type' => 'query',
        );
        $res = $this->post($url, $data);
        if (empty($res) || !$res['status']) {
            return [];
        }
        return array(
            'applicants' => $this->filterHtml($res['data']['AN']['OFFICIAL'][0]),
            'inventors' => implode(',', $res['data']['INS']),
            'application_time' => date('Y-m-d', strtotime($res['data']['APD'])),
            'legal_status' => $this->legal_status[$res['data']['LEGAL_STATUS'][0]] ?? $this->legal_status[3],
            'patent_type' => $this->patent_types[$res['data']['PATENT_TYPE']] ?? $this->patent_types['U'],
            'obligee' => $this->filterHtml($res['data']['ANC']['OFFICIAL'][0]),
            'abstract' => $this->filterHtml($res['data']['ABST']['CN']),
            'title' => $this->filterHtml($res['data']['title_default']),
            'number' => $res['data']['APN'],
            'pn' => $res['data']['PN'],
            'image' => $res['data']['PATSNAP_IMAGE']['url'],
        );
    }

    public function getPDF($patent_id) {
        $url = $this->searchHost . 'patent/id/' . $patent_id . '/pdf';
        $this->headers['X-Patsnap-From'] = 'w-analytics-search-view';
        $res = $this->get($url);
        if (empty($res) || !$res['status']) {
            return '';
        }
        return $res['data']['PDF'];
    }

    public function saveData($data) {
        $path = storage_path('/ppt/'. $data['pn']);
        if (file_exists($path)) {
            return;
        }
        mkdir($path);
        $imgPath = $path . '/' . $data['number']. '.png';
        $pdfPath = $path . '/' . $data['number']. '.pdf';
        $jsonPath = $path . '/' . $data['number']. '.json';
        if (!empty($data['image'])) {
            file_put_contents($imgPath, file_get_contents($data['image']));
        }
        if (!empty($data['pdf'])) {
            file_put_contents($pdfPath, file_get_contents($data['pdf']));
        }

        unset($data['image'], $data['pdf']);
        file_put_contents($jsonPath, json_encode($data, JSON_UNESCAPED_UNICODE));
    }

    private function filterHtml($str) {
        return preg_replace('/<[^>]+>/', '', $str);
    }

    private function get($url) {
        $response = Http::withToken($this->token)->withHeaders($this->headers)->get($url);
        if ($response->successful()) {
            return $response->json();
        }
        return [];
    }

    public function post($url, $data = array()) {
        $response = Http::withToken($this->token)->withHeaders($this->headers)->post($url, $data);
        if ($response->successful()) {
            return $response->json();
        }
        return [];
    }
}