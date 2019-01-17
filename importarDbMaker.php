<?php
if ($_SERVER['REQUEST_METHOD'] == 'POST') {

    //Exibe todos os erros do PHP
    error_reporting(E_ALL);
    ini_set('display_errors', true);
    ini_set('display_startup_errors', true);

    //Altera time zone para SP.
    date_default_timezone_set('America/Sao_Paulo');

    //Carrega bibliotecas de leitura para XLS.
    require_once('../../lib/xls-reader/php-excel-reader/excel_reader2.php');
    require_once('../../lib/xls-reader/SpreadsheetReader.php');

    //Carrega configurações de ODBC
    $config = parse_ini_file('/etc/odbc.ini', true);
    //Informa como as querys foram montadas na tela do usuário quando marcado checkbox
    $debug = isset($_POST['debug']) ? true : false;

    //Extenções de arquivos Permitidas
    $extPermitido = ['.xlsx', '.xls'];
    //Pega extenção do arquivo que foi feito upload
    $extensao = strrchr($_FILES["arquivo"]['name'], '.');

    //Verifica a extenção do arquivo com as extenções permitidas
    if (!in_array($extensao, $extPermitido)) {
        echo "Extenção de arquivo inválida apenas <b>.xls, .xlsx</b> é permitido.";
        exit;
    }

    //Diretório que vai ser efetuado o upload do arquivo temporário
    $uploadfile = '../../tmp2/' . basename($_FILES['arquivo']['name']);
    //Move o arquivo para o diretório temporário
    move_uploaded_file($_FILES['arquivo']['tmp_name'], $uploadfile);

    //Verifica se o arquivo foi movido e é válido
    if (!file_exists($uploadfile)) {
        echo json_encode(array(
            'codigo' => 2,
            'mensagem' => utf8_encode('Arquivos não localizados.')
        ));
        exit;
    }

    //Abre xls
    $dadosRecebidos = new SpreadsheetReader($uploadfile);

    //Diz quais campos devem ficar sem aspas
    $tipoSemAspa = array("INTEGER", "SMALLINT", "DECIMAL", "BIGINT");

    //Verifica nome da DB.
    $a = 0;
    foreach ($dadosRecebidos as $linha) {
        $a++;
        if ($a == 2) {
            $db = $linha[0];
            break;
        }
    }

    //Emite erro caso não encontra DB.
    if (!isset($config[$db]['User'])) {
        echo "Database <b>{$db}</b> não encontrada nos drivers ODBC. Necessário cadastrar em <b>/etc/odbc.ini</b><br>";
        exit;
    }

    //Inicia conexão com banco de dados
    $conexao = NULL;
    try {
        $conexao = new PDO("odbc:{$db}", $config[$db]['User'], $config[$db]['Password']);
    } catch (PDOException $e) {
        echo "<pre>";
        var_dump($e);
        $conexao = false;
        exit;
    }

    //Monta cabeçalho com colunas da tabela
    foreach ($dadosRecebidos as $linha) {
        foreach ($linha AS $coluna)
            $cabecalho[] = $coluna;
        break;
    }

    //Busca todas as tabelas que irão ser inseridos registros.
    $c = 0;
    $tabelasInserir = [];
    foreach ($dadosRecebidos as $linha) {
        $c++;
        if ($c == 1)
            continue;
        if($linha[1] == '')
            continue;
        if (!isset($tabelasInserir[$linha[1]]))
            $tabelasInserir[$linha[1]] = $linha[1];
    }
    var_dump($tabelasInserir);

    //Busca no banco de dados todas as tabelas que serão trabalhadas e monta uma lista com tipo de cada coluna
    $tipoColunas = [];
    foreach ($tabelasInserir AS $tabela) {
        $tabela = strtoupper($tabela);
        $dadosTabela = $conexao->query("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='{$tabela}'")->fetchAll();
        foreach ($dadosTabela AS $dado) {
            $tipoColunas[$tabela][trim($dado['COLUMN_NAME'])] = trim($dado['DATA_TYPE']);
        }

        //Valida o cabeçalho verificando se existem colunas que não estão na tabela
        $b = 0;
        foreach ($cabecalho AS $cab) {
            $b++;
            if ($b > 2) {
                $cabPesquisar = trim($cab);
                if (!isset($tipoColunas[$tabela][strtoupper($cabPesquisar)]))
                    $colunasInvalidas[] = '<b>' . trim($cab) . '</b>';
            }
        }

        //Emite erro identificando quais colunas não foram encontradas.
        if (isset($colunasInvalidas)) {
            $colunasInvalidas = implode(', ', $colunasInvalidas);
            echo "Existem colunas no arquivo que não foram localizadas no banco de dados.<br>"
            . "Referente a tabela: <b> {$tabela} </b> <br>"
            . "Colunas não localizadas: {$colunasInvalidas} <br>";
            exit;
        }
    }

    $d = 0;
    $posicao = 1;
    //Inicia transação
    $conexao->beginTransaction();
    //Inicia processo de leitura para gravação
    foreach ($dadosRecebidos as $linha) {
        $d++;

        //Pula linha do cabeçalho
        if ($d == 1)
            continue;
        
        if($linha[1] == '')
            continue;

        $posicao++;

        $i = 0;
        //Combina cabeçalho com dados da linha
        $registro = array_combine($cabecalho, $linha);

        //Informa o inicio da leitura da linha
        if ($debug)
            echo "Iniciando linha <b>{$posicao}</b>:<br>";
        $colunas = [];
        $infos = [];

        //Inicia montagem de estrutura de query
        foreach ($cabecalho AS $cab) {
            if ($registro[$cab] == '')
                continue;
            $i++;
            //Pula as duas primeiras colunas DB, TABELA
            if ($i > 2) {
                $cabPesquisar = trim($cab);
                $colunas[] = strtoupper($cab);
                //Faz tratamento com as colunas de acordo com a necessidade
                if (in_array($tipoColunas[$tabela][strtoupper($cabPesquisar)], $tipoSemAspa)) {
                    $registro[$cab] = ($registro[$cab]{0} == '-' ? -1 : 1) * preg_replace("/[^\d\.]/", '', $registro[$cab]);
                    $valor = $registro[$cab] == '' ? 'NULL' : "{$registro[$cab]}";
                } else if ($tipoColunas[$tabela][strtoupper($cabPesquisar)] == 'DATE') {
                    $valor = substr($registro[$cab], 6, 4) . '-' . substr($registro[$cab], 3, 2) . '-' . substr($registro[$cab], 0, 2);
                } else {
                    $valor = "'{$registro[$cab]}'";
                }
                //Gera array com os valores finais
                $infos[] = $valor;

                //Informa ao usuário como foi montado os campos
                if ($debug)
                    echo 'Coluna: ' . strtoupper($cabPesquisar) . ' Tipo: ' . $tipoColunas[$tabela][strtoupper($cabPesquisar)] . ' Valor: ' . $valor . '<BR>';
            }
        }

        //Junta colunas
        $colunas = implode(',', $colunas);
        //Junta Informações da linha
        $infos = implode(',', $infos);
        //Monta query com base nos resultados
        $query = 'INSERT INTO ' . $tabela . ' (' . $colunas . ') VALUES (' . $infos . ')';

        //Informa query montada ao usuário
        if ($debug)
            echo $query . '<br>';

        //Executa query e faz validação se houve algum erro para que retorne ao usuário.
        try {
            $executar = $conexao->prepare("{$query}");
            $executar->execute();
        } catch (PDOException $e) {
            echo "<pre>";
            var_dump($e);
            //Se ocorrer erro da rollback em tudo.
            $conexao->rollback();
            $conexao = false;
            exit;
        }

        //Informa ao usuário que finalizou a linha e inseriu com sucesso.
        if ($debug)
            echo "Finalizada linha <b>{$posicao}</b><br><br>";
    }

    //Commita a transação salvando os dados;
    $conexao->commit();
    //Apaga o arquivo temporário.
    unlink($uploadfile);
    //Informa ao usuário que finalizou o processo.
    echo "Processo concluído.";
}
?>

<form method="POST" action="importarDbMaker.php" enctype="multipart/form-data">
    <label>Arquivo</label>
    <input type="file" name="arquivo"><br>
    <input type="checkbox" name="debug"> Mostrar histórico<br><br>
    <input type="submit" value="Importar">
</form>