<?php

namespace Sucata;

use PDO;
use PDOException;
use RuntimeException;

class DBase
{

    private $myparam = 'array(PDO::MYSQL_ATTR_INIT_COMMAND => "SET NAMES utf8"))';
    private $connection;
    private static $instance;

    public static function getInstance()
    {
        if (!self::$instance) {
            self::$instance = new self();
        }
        return self::$instance;
    }

    public function __construct($host = '', $dbname = '', $user = '', $pass = '')
    {
        if (empty($host)) {
            $host = getenv('DB_HOST');
            $dbname = getenv('DB_NAME');
            $user = getenv('DB_USER');
            $pass = getenv('DB_PASS');
        }    
        $dsn = "mysql:host={$host};dbname={$dbname}";

        try {
            $this->connection = new PDO($dsn, $user, $pass);
        } catch (PDOException $e) {
            throw new RuntimeException("Error: Falha na conexão : " . $e->getMessage());
        }
    }

    private function __clone()
    {
        
    }

    public function getConnection()
    {
        return $this->connection;
    }
    
    public function query(
        $sqlComm,
        $data = array()
    ) {
        $aRet = array();
        $properties = array(PDO::ATTR_CURSOR => PDO::CURSOR_FWDONLY);
        try {
            if (!empty($properties) && !empty($data)) {
                $sth = $this->connection->prepare($sqlComm, $properties);
                $sth->execute($data);
                $aRet = $sth->fetchAll();
            } else {
                $resp = $this->connection->query($sqlComm);
                if (!empty($resp)) {        
                    foreach($resp as $row) {
                        $aRet[]=$row;
                    }
                }
            }
        } catch (PDOException $e) {
            throw new RuntimeException("Falha em query: " . $e->getMessage());
        }
        return $aRet;
    }
}
