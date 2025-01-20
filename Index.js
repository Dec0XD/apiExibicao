const express = require('express');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// Função para carregar dados do Excel
const carregarDadosExcel = () => {
    const caminhoArquivo = path.join(__dirname, 'data/RELAT FALHA CML OPEC TVBA-CCAST 2024.xlsx');
    const workbook = xlsx.readFile(caminhoArquivo);
    const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
    const dados = xlsx.utils.sheet_to_json(primeiraAba, { raw: false, dateNF: 'dd/mm/yyyy' });
    return dados;
};

// Rota para fornecer dados do Excel
app.get('/dados', (req, res) => {
    try {
        const dados = carregarDadosExcel();
        res.json(dados);
    } catch (erro) {
        res.status(500).json({ mensagem: 'Erro ao carregar dados do Excel', erro });
    }
});

app.listen(port, () => {
    console.log(`API rodando em http://localhost:${port}`);
});
