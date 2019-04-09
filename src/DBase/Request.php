<?php

namespace Sucata\DBase;

use Sucata\DBase\DBase;

class Request
{
    public static function get($sqlComm = '')
    {
        if (empty($sqlComm)) {
            return [];
        }
        $config = json_encode(['host' => 'localhost','user'=>'root', 'pass'=>'monitor5', 'db'=>'blabel']);
        $dbase = new DBase($config);
        $dbase->connect();
        return $dbase->query($sqlComm);
    }
}
