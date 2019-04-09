<?php

namespace Sucata\DBase;

use PDO;
use PDOException;
use RuntimeException;

class DBase
{
    private $conf;
    private $myparam = 'array(PDO::MYSQL_ATTR_INIT_COMMAND => "SET NAMES utf8"))';
    private $dsn;
    private $dbh = null;
    private $connected = false;
    private $transaction = false;

    /**
     * Construtor da classe 
     */
    public function __construct($config)
    {
        $conf = null;
        if (is_file($config)) {
            $conf = json_decode(file_get_contents($config));
        } elseif ($config != '') {
            $conf = json_decode($config);
        }
        if (is_null($conf)) {
            return null;
        }
        $this->conf = $conf;
        $host = $this->conf->host;
        $db = $this->conf->db;
        $this->dsn = "mysql:host=$host;dbname=$db";
    }
    
    public function __destruct()
    {
        $this->unconnect();
    }
    
    /**
     * Conectar a base de dados
     * @return boolean 
     */
    public function connect()
    {
        if ($this->connected) {
            return true;
        }
        try {
            $this->dbh = new PDO($this->dsn, $this->conf->user, $this->conf->pass);
            $this->connected = true;
        } catch (PDOException $e) {
            throw new RuntimeException("Error: Falha na conexão : " . $e->getMessage());
        }
        return true;
    }
    
    public function unconnect()
    {
        $this->dbh = null;
        $this->connected = false;
    }
    
    public function tableExists($table)
    {
        if (! $this->connected) {
            $this->connect();
        }
        $results = false;
        try {
            $results = $this->dbh->query("SELECT 1 FROM $table");
        } catch (PDOException $e) {
            throw new RuntimeException("Falha em tableExists: " . $e->getMessage());
        }
        return $results;
    }
    
    public function isEmpty($table)
    {
        $sqlComm = "SELECT count(id) FROM $table";
        if (count($this->query($sqlComm) == 0)) {
            return true;
        }
        return false;
    }
    
    public function truncate($table)
    {
        return $this->execute("TRUNCATE TABLE $table;");
    }
    
    public function drop($table)
    {
        return $this->execute("DROP TABLE $table;");
    }
    
    /**
     * Executa uma pesquisa SQL
     * @param string $sqlComm
     * @param array $data
     * @return array 
     */
    public function query(
        $sqlComm,
        $data = array()
    ) {
        if (! $this->connected) {
            $this->connect();
        }
        $aRet = array();
        $properties = array(PDO::ATTR_CURSOR => PDO::CURSOR_FWDONLY);
        try {
            if (!empty($properties) && !empty($data)) {
                $sth = $this->dbh->prepare($sqlComm, $properties);
                $sth->execute($data);
                $aRet = $sth->fetchAll();
            } else {
                $resp = $this->dbh->query($sqlComm);
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
    
    /**
     * Executa comando SQL
     * @param string $sqlComm
     * @param array $data
     * @return boolean
     */
    public function execute($sqlComm, $data = [])
    {
        if (! $this->connected) {
            $this->connect();
        }
        $flag = true;
        try {
            if (is_array($sqlComm)) {
                foreach ($sqlComm as $sql) {
                    if (! empty($data)) {
                        $stmt = $this->dbh->prepare($sql);
                        $flag &= $stmt->execute($data);
                    } else {
                        $stmt = $this->dbh->prepare($sql);
                        $flag &= $stmt->execute();
                    }
                }
            } else {
                $count = 0;
                if (! empty($data)) {
                    $stmt = $this->dbh->prepare($sqlComm);
                    $flag &= $stmt->execute($data);
                } else {
                    //$stmt = $this->dbh->prepare($sqlComm);
                    //$flag &= $stmt->execute();
                    $flag &= $this->dbh->exec($sqlComm);
                }
            }
            if ($this->transaction && ! $flag) {
                $this->rollbackTrans();
                return false;
            }
        } catch (PDOException $e) {
            if ($this->transaction ) {
                $this->rollbackTrans();
            }
            throw new RuntimeException("Falha em execute: " . $e->getMessage());
        } catch (\Exception $e) {
            throw new RuntimeException("Falha em execute: " . $e->getMessage());
        }
        return $flag;
    }

    /**
     * Inicia a Transaction
     */
    public function beginTrans()
    {
        if (! $this->connected) {
            $this->connect();
        }
        $this->dbh->beginTransaction();
        $this->transaction = true;
    }
    
    /**
     * Retorna os comandos da Transaction
     */
    public function rollbackTrans()
    {
        if ($this->transaction && $this->connected) {
           $this->dbh->rollBack();
        }
        $this->transaction = false;
    }

    /**
     * Realiza os comandos da Transaction
     */
    public function commitTrans()
    {
        if ($this->transaction && $this->connected) {
            $this->dbh->commit();
        }
        $this->transaction = false;
    }

    /**
     * Busca o Id do último comando "insert"
     * return integer 
     */
    public function lastId()
    {
        if (! $this->connected) {
            return 0;
        }
        return $this->dbh->lastInsertId();
    }
}
