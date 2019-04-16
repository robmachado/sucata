<?php
ini_set("display_errors", 1);
error_reporting(E_ALL);
require_once '../bootstrap.php';

use Sucata\DBase\DBase;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use Carbon\Carbon;

$base = basename(__FILE__);


$dtini = filter_input(INPUT_POST, 'dtini', FILTER_SANITIZE_STRING);
$dtfim = filter_input(INPUT_POST, 'dtfim', FILTER_SANITIZE_STRING);
$complete = filter_input(INPUT_POST, 'complete', FILTER_SANITIZE_STRING);

if (!empty($dtini) && !empty($dtfim)) {
    $dti = Carbon::createFromFormat('Y-m-d H:i:s', "{$dtini} 00:00:00")->format('Y-m-d H:i:s');
    $dtf = Carbon::createFromFormat('Y-m-d H:i:s', "{$dtfim} 23:59:59")->format('Y-m-d H:i:s');
    if ($complete) {
        $tipo = "COMPLETO";
        $sql = "SELECT * FROM appoint WHERE date >= '$dti' AND date <= '$dtf' ORDER BY orders_id, sector;";
    } else {
        $tipo = "RESUMO";
        $sql = "SELECT sector, type, quality, sum(net) as peso "
            . "FROM appoint "
            . "WHERE date >= '$dti' AND date <= '$dtf' "
            . "GROUP BY sector, type, quality "
            . "ORDER BY sector, type, quality;";
    }
    /**
     * type = tipo de sucata 1-APARA e 2-REFILE
     * quality = qualidade da sucata 1-TRANSPARENTE 2-COLORIDO 3-EVA
     */
    $config = json_encode(
        [
            'host' => 'mercurio',
            'user'=>'etiqueta',
            'pass'=>'forever',
            'db'=>'blabel'
        ]
    );
    $dbase = new DBase($config);
    $dbase->connect();
    $resp = $dbase->query($sql);
    $dados = [];
    $i = 0;
    if (empty($resp)) {
        echo "<center><H1>Não foram localizados registros para esse período.</H1></center>";
        die;
    }
    
    foreach($resp as $lin) {
        $dados[$i]['data'] = isset($lin['date']) ? $lin['date'] : '';
        $dados[$i]['op'] = isset($lin['orders_id']) ? $lin['orders_id'] : '';
        $dados[$i]['setor'] = isset($lin['sector']) ? $lin['sector'] : '';
        $dados[$i]['tipo'] = $lin['type'] == 1 ? 'APARA' : 'REFILE';
        $dados[$i]['qual'] = $lin['quality'] == 1 ? 'TRANSPARENTE' : ($lin['quality'] == 3 ? 'EVA' : 'COLORIDO'); 
        $dados[$i]['peso'] = isset($lin['net']) ? $lin['net'] : $lin['peso'];
        $i++;
    }
    
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $locale = 'pt_br';
    $spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
    $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(8);
    
    $sheet->setCellValue('B2', 'Relatório de Sucata');
    $sheet->setCellValue('B3', "Período: " . (new Carbon($dtini))->format('d/m/Y') . " até " . (new Carbon($dtfim))->format('d/m/Y'));
    $sheet->setCellValue('B4', $tipo);
    
    if (!$complete) {
        $sheet->setCellValue('B6', 'Setor');
        $sheet->setCellValue('C6', 'Tipo');
        $sheet->setCellValue('D6', 'Qualidade');
        $sheet->setCellValue('E6', 'Peso');
        $i = 7;
        foreach($dados as $d) {
            $sheet->setCellValue("B{$i}", $d['setor']);
            $sheet->setCellValue("C{$i}", $d['tipo']);
            $sheet->setCellValue("D{$i}", $d['qual']);
            $sheet->setCellValue("E{$i}", $d['peso']);
            $i++;
        }
    } else {
        $sheet->setCellValue('B6', 'Data');
        $sheet->setCellValue('C6', 'OP');
        $sheet->setCellValue('D6', 'Setor');
        $sheet->setCellValue('C6', 'Tipo');
        $sheet->setCellValue('D6', 'Qualidade');
        $sheet->setCellValue('E6', 'Peso');
        $i = 7;
        foreach($dados as $d) {
            $c = new Carbon($d['data']);
            $sheet->setCellValue("B{$i}", $c->format('d/m/Y'));
            $sheet->setCellValue("C{$i}", $d['op']);
            $sheet->setCellValue("D{$i}", $d['setor']);
            $sheet->setCellValue("C{$i}", $d['tipo']);
            $sheet->setCellValue("D{$i}", $d['qual']);
            $sheet->setCellValue("E{$i}", $d['peso']);
            $i++;
        }
    }
    
    $writer = new Xls($spreadsheet);
    
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header("Content-Disposition: attachment;filename=\"sucata.xls\"");
    header("Cache-Control: max-age=0");
    $writer->save('php://output');
    die;
}

$title = "Sucata";
$body = "
<div class=\"container-fluid\">
    <div class=\"row\">
        <div class=\"col-md-4\"></div>
        <div class=\"col-md-4\">
            <h2>Relatório de Sucata</h2>
        </div>
        <div class=\"col-md-4\"></div>
    </div>
    <form role=\"form\" method=\"POST\" action=\"{$base}\" >
    <div class=\"row\">
        <div class=\"col-md-2\"></div>
        <div class=\"col-md-8\">
            <div class=\"row\">
                <div class=\"col-md-3\">
                    <label for=\"dtini\">Data Inicial</label> 
                    <div class=\"form-group\">
                        <input type=\"date\" class=\"form-control\" id=\"dtini\" name=\"dtini\" autofocus required>
                    </div>
                </div>
                <div class=\"col-md-3\">
                    <label for=\"dtfim\">Data Final</label> 
                    <div class=\"form-group\">
                        <input type=\"date\" class=\"form-control\" id=\"dtfim\" name=\"dtfim\" required>
                    </div>
                </div>
                <div class=\"col-md-3\">
                    <div class=\"form-check vcenter\">
                        <br> 
                        <input type=\"checkbox\" class=\"form-check-input\" id=\"complete\" name=\"complete\">
                        <label class=\"form-check-label\" for=\"complete\">Gerar completo</label>
                    </div>
                </div>
                <div class=\"col-md-3\">
                    <button type=\"submit\" class=\"btn btn-primary\">Gerar</button>
                </div>
            </div>
        </div>
        <div class=\"col-md-2\"></div>
    </div>
    </form>
</div>";


$html = file_get_contents('main.html');
$html = str_replace("{{extras}}", '', $html);
$html = str_replace("{{title}}", $title, $html);
$html = str_replace("{{content}}", $body, $html);
$html = str_replace("{{script}}", "", $html);
echo $html;
