<?php

namespace Chuva\Php\WebScrapping;

use Box\Spout\Common\Type;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

class Main
{
    public static function run()
    {
        // Carregar o conteudo HTML
        $htmlContent = file_get_contents('C:\\Users\\nicol\\OneDrive\\Área de Trabalho\\Nova pasta\\exercicios-2024\\php\\assets\\origin.html');

        // Criar um novo objeto
        $dom = new \DOMDocument();
        $dom->loadHTML($htmlContent);

        // Inicializar um array para armazenar os dados
        $trabalhos = [];

        // Encontrar todos os links <a> com a classe "paper-card"
        $links = $dom->getElementsByTagName('a');
        foreach ($links as $link) {
            // Verificar se o link possui a classe "paper-card"
            if ($link->getAttribute('class') === 'paper-card p-lg bd-gradient-left') {
                // Extrair as informações do título e do ID do link
                $titulo = $link->getElementsByTagName('h4')[0]->textContent;
                $id = $link->getElementsByTagName('div')[1]->getElementsByTagName('div')[1]->getElementsByTagName('div')[1]->textContent;

                // Extrair o tipo do trabalho
                $tipo = '';
                $tags = $link->getElementsByTagName('div');
                foreach ($tags as $tag) {
                    if ($tag->getAttribute('class') === 'tags mr-sm') {
                        $tipo = $tag->textContent;
                        break;
                    }
                }

                // Extrair os autores e instituiçoes
                $autores = [];
                $instituicoes = [];
                $autoresNodes = $link->getElementsByTagName('span');
                foreach ($autoresNodes as $index => $autorNode) {
                    if ($index % 2 === 0) {
                        $autores[] = $autorNode->textContent;
                    } else {
                        $instituicoes[] = $autorNode->getAttribute('title');
                    }
                }

                // Inicializar arrays vazios para autores e instituiçoes
                $autoresData = [];
                $instituicoesData = [];

                // Preencher os arrays de autores e instituiçoes com dados nulos
                for ($i = 0; $i < 9; $i++) {
                    $autoresData[] = isset($autores[$i]) ? $autores[$i] : '';
                    $instituicoesData[] = isset($instituicoes[$i]) ? $instituicoes[$i] : '';
                }

                // Montar o array com os dados
                $trabalho = [
                    'ID' => $id,
                    'Title' => $titulo,
                    'Type' => $tipo,
                    'Author 1' => $autoresData[0],
                    'Author 1 Institution' => $instituicoesData[0],
                    'Author 2' => $autoresData[1],
                    'Author 2 Institution' => $instituicoesData[1],
                    'Author 3' => $autoresData[2],
                    'Author 3 Institution' => $instituicoesData[2],
                    'Author 4' => $autoresData[3],
                    'Author 4 Institution' => $instituicoesData[3],
                    'Author 5' => $autoresData[4],
                    'Author 5 Institution' => $instituicoesData[4],
                    'Author 6' => $autoresData[5],
                    'Author 6 Institution' => $instituicoesData[5],
                    'Author 7' => $autoresData[6],
                    'Author 7 Institution' => $instituicoesData[6],
                    'Author 8' => $autoresData[7],
                    'Author 8 Institution' => $instituicoesData[7],
                    'Author 9' => $autoresData[8],
                    'Author 9 Institution' => $instituicoesData[8],
                ];

                // Adicionar os dados
                $trabalhos[] = $trabalho;
            }
        }

        // Criar nova planilha Excel
        $writer = WriterEntityFactory::createXLSXWriter();
        $filePath = 'C:\\Users\\nicol\\OneDrive\\Área de Trabalho\\Nova pasta\\exercicios-2024\\php\\assets\\model.xlsx';
        $writer->openToFile($filePath);

        // Adicionar cabeçalhos a planilha
        $headerRow = WriterEntityFactory::createRowFromArray(['ID', 'Title', 'Type', 'Author 1', 'Author 1 Institution', 'Author 2', 'Author 2 Institution', 'Author 3', 'Author 3 Institution', 'Author 4', 'Author 4 Institution', 'Author 5', 'Author 5 Institution', 'Author 6', 'Author 6 Institution', 'Author 7', 'Author 7 Institution', 'Author 8', 'Author 8 Institution', 'Author 9', 'Author 9 Institution']);
        $writer->addRow($headerRow);

        // Adicionar os dados a planilha
        foreach ($trabalhos as $trabalho) {
            // Adicionar uma linha na planilha
            $rowData = WriterEntityFactory::createRowFromArray([
                $trabalho['ID'],
                $trabalho['Title'],
                $trabalho['Type'],
                $trabalho['Author 1'],
                $trabalho['Author 1 Institution'],
                $trabalho['Author 2'],
                $trabalho['Author 2 Institution'],
                $trabalho['Author 3'],
                $trabalho['Author 3 Institution'],
                $trabalho['Author 4'],
                $trabalho['Author 4 Institution'],
                $trabalho['Author 5'],
                $trabalho['Author 5 Institution'],
                $trabalho['Author 6'],
                $trabalho['Author 6 Institution'],
                $trabalho['Author 7'],
                $trabalho['Author 7 Institution'],
                $trabalho['Author 8'],
                $trabalho['Author 8 Institution'],
                $trabalho['Author 9'],
                $trabalho['Author 9 Institution'],
            ]);
            $writer->addRow($rowData);
        }

        // Fechar a planilha
        $writer->close();

        echo "Planilha criada com sucesso em: $filePath";
    }
}

Main::run();
