<?php

ini_set("display_errors", 1);
error_reporting(E_ALL);
require_once '../vendor/autoload.php';

use Sucata\DBase;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Carbon\Carbon;

$base = basename(__FILE__);

$db1 = new DBase('mercurio', 'legacy', 'etiqueta', 'forever');

$dtini = filter_input(INPUT_POST, 'dtini', FILTER_SANITIZE_STRING);
$dtfim = filter_input(INPUT_POST, 'dtfim', FILTER_SANITIZE_STRING);

if (!empty($dtini) && !empty($dtfim)) {
    $dti = Carbon::createFromFormat('Y-m-d H:i:s', "{$dtini} 00:00:00")->format('Y-m-d H:i:s');
    $dtf = Carbon::createFromFormat('Y-m-d H:i:s', "{$dtfim} 23:59:59")->format('Y-m-d H:i:s');
    //$sql = "SELECT * FROM apontamentos WHERE data >= '$dti' AND data <= '$dtf' ORDER BY data, shifttimeini, maq;";
    $sql = "SELECT ord.cliente, ord.nome, apo.* FROM apontamentos apo LEFT JOIN ordens ord ON ord.id = apo.numop WHERE apo.data >= '2019-05-01' AND apo.data <= '2019-05-14' ORDER BY apo.maq, apo.data, apo.shifttimeini;";
    /**
     * type = tipo de sucata 1-APARA e 2-REFILE
     * quality = qualidade da sucata 1-TRANSPARENTE 2-COLORIDO 3-EVA
     */
    $dados = $db1->query($sql);
    if (empty($dados)) {
        echo "<center><H1>Não foram localizados registros para esse período.</H1></center>";
        die;
    }
    //echo "<pre>";
    //print_r($dados);
    //echo "</pre>";
    //die;

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $locale = 'pt_br';
    $spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
    $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(8);

    $sheet->setCellValue('B2', 'Relatório de Produção');
    $sheet->setCellValue('B3', "Período: " . (new Carbon($dtini))->format('d/m/Y') . " até " . (new Carbon($dtfim))->format('d/m/Y'));
    //$sheet->setCellValue('B4', $tipo);

    $styleArray = [
        'font' => [
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKGRAY
        ],
    ];

    $sheet->setCellValue('B6', 'Maquina');
    $sheet->setCellValue('C6', 'Data');
    $sheet->setCellValue('D6', 'H.Inicio');
    $sheet->setCellValue('E6', 'H.Fim');
    $sheet->setCellValue('F6', 'Turno');
    $sheet->setCellValue('G6', 'OP');
    $sheet->setCellValue('H6', 'Cod.Parada');
    $sheet->setCellValue('I6', 'Qtdade');
    $sheet->setCellValue('J6', 'Unidade');
    $sheet->setCellValue('K6', 'Fator');
    $sheet->setCellValue('L6', 'Tempo SetUp');
    $sheet->setCellValue('M6', 'N. Operadores');
    $sheet->setCellValue('N6', 'Velocidade');
    $sheet->setCellValue('O6', 'Refile');
    $sheet->setCellValue('P6', 'Aparas');
    $sheet->setCellValue('Q6', 'Peso Total');
    $sheet->setCellValue('R6', 'Metragem');
    $sheet->setCellValue('S6', 'Cliente');
    $sheet->setCellValue('T6', 'Descrição');
    $i = 7;
    foreach ($dados as $d) {
        
        $peso = 0;
        if ($d['uni'] == 'KG') {
            $peso = $d['qtd'];
        } elseif ($d['uni'] == 'PC' && is_numeric($d['qtd']) && is_numeric($d['fator'])) {
            $peso = round($d['fator'] * $d['qtd']/1000, 2);
        }
        $metragem = '';
        if (substr($d['maq'], 0, 1) != 'C'
            && is_numeric($d['qtd'])
            && is_numeric($d['fator'])
            && $d['fator'] > 0
        ) {
            $metragem = round(($d['qtd'] * 1000)/$d['fator'], 0);
        }
        
        $dt = Carbon::createFromFormat('Y-m-d', $d['data'])->format('d/m/Y');
        $sheet->setCellValue("B{$i}", $d['maq']);
        $sheet->setCellValue("C{$i}", $dt);
        $sheet->setCellValue("D{$i}", $d['shifttimeini']);
        $sheet->setCellValue("E{$i}", $d['shifttimefim']);
        $sheet->setCellValue("F{$i}", $d['turno']);
        $sheet->setCellValue("G{$i}", $d['numop']);
        $sheet->setCellValue("H{$i}", $d['parada']);
        $sheet->setCellValue("I{$i}", $d['qtd']);
        $sheet->setCellValue("J{$i}", $d['uni']);
        $sheet->setCellValue("K{$i}", $d['fator']);
        $sheet->setCellValue("L{$i}", $d['setup']);
        $sheet->setCellValue("M{$i}", $d['ops']);
        $sheet->setCellValue("N{$i}", $d['velocidade']);
        $sheet->setCellValue("O{$i}", $d['refile']);
        $sheet->setCellValue("P{$i}", $d['aparas']);
        
        $sheet->setCellValue("Q{$i}", $peso);
        $sheet->setCellValue("R{$i}", $metragem);
        $sheet->setCellValue("S{$i}", $d['cliente']);
        $sheet->setCellValue("T{$i}", $d['nome']);
        $i++;
    }
    //$sheet->getStyle("E7:E{$i}")
    //    ->getNumberFormat()
    //    ->setFormatCode('#,##0.0 "kg"');
    //$sheet->getStyle("B6:P6")
    //    ->getAlignment()
    //    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    $sheet->getStyle('B6:T6')->applyFromArray($styleArray);

    $writer = new Xls($spreadsheet);

    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header("Content-Disposition: attachment;filename=\"relprod.xls\"");
    header("Cache-Control: max-age=0");
    $writer->save('php://output');
    die;
}

$title = "Producao";
$body = "
<div class=\"container-fluid\">
    <div class=\"row\">
        <div class=\"col-md-4\"></div>
        <div class=\"col-md-4\">
            <h2>Relatório de Produção</h2>
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
