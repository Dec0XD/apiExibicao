const express = require('express');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// Função para carregar dados do Excel
const carregarDadosOPEC = () => {
    const caminhoArquivo = path.join(__dirname, 'data/RELAT FALHA CML OPEC TVBA-CCAST 2024 - Copia de teste api.xlsx');
    const workbook = xlsx.readFile(caminhoArquivo);
    const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
    const dados = xlsx.utils.sheet_to_json(primeiraAba, { raw: false, dateNF: 'dd/mm/yyyy' });
    return dados;
};
const carregarDadosPGM_TVBA = () => {
    const caminhoArquivo = path.join(__dirname, 'data/Planilha PGM Exibição - TVBA 2024.xlsm');
    const workbook = xlsx.readFile(caminhoArquivo);
    const primeiraAba = workbook.Sheets[workbook.SheetNames[2]];
    const dados = xlsx.utils.sheet_to_json(primeiraAba, { raw: false, dateNF: 'dd/mm/yyyy' });
    return dados;
};

// Rota para fornecer dados do Excel
app.get('/dadosOPEC', (req, res) => {
    try {
        const dados = carregarDadosOPEC();
        res.json(dados);
    } catch (erro) {
        res.status(500).json({ mensagem: 'Erro ao carregar dados do Excel', erro });
    }
});

app.get('/dadosPGM_TVBA', (req, res) => {
    try {
        const dados = carregarDadosPGM_TVBA();
        res.json(dados);
    } catch (erro) {
        res.status(500).json({ mensagem: 'Erro ao carregar dados do Excel', erro });
    }
});

app.listen(port, () => {
    console.log(`API rodando em http://localhost:${port}`);
});
