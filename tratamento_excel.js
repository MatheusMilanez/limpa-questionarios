const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");

// Configurações
const pastaInput = path.join(__dirname, "input");
const pastaOutput = path.join(__dirname, "output");
if (!fs.existsSync(pastaInput)) fs.mkdirSync(pastaInput);
if (!fs.existsSync(pastaOutput)) fs.mkdirSync(pastaOutput);

// Gabaritos por questionário
const gabaritos = {
  94: [0, 1, 0, 1, 0, 0, 1, 0, 1, 0],
  95: [2, 3, 0, 1, 3, 2, 1, 0, 3, 0],
};

function tratarPlanilha(arquivo) {
  const workbook = XLSX.readFile(arquivo);
  const abas = ["94", "95"];

  abas.forEach((nomeAba) => {
    const aba = workbook.Sheets[nomeAba];
    if (!aba) return;

    let dados = XLSX.utils.sheet_to_json(aba, { defval: "" });

    // Limpar colunas Unnamed
    dados = dados.map((linha) => {
      const novaLinha = {};
      Object.keys(linha).forEach((coluna) => {
        if (!coluna.toLowerCase().includes("unnamed")) {
          novaLinha[coluna.trim()] = linha[coluna];
        }
      });
      return novaLinha;
    });

    const resultados = [];

    if (nomeAba === "95") {
      const gabarito = gabaritos[nomeAba];

      dados.forEach((linha, index) => {
        const email =
          linha["Email"] ||
          linha["email"] ||
          linha["Aluno"] ||
          linha["aluno"] ||
          "Aluno sem e-mail";

        const numeroQuestao = index % 10;
        const pergunta = `Q${numeroQuestao + 1}`;
        const respostaBruta = linha["answer"] || linha["Answer"] || "";

        let resposta =
          respostaBruta !== "" && !isNaN(respostaBruta)
            ? parseInt(respostaBruta)
            : null;

        const alternativaCorreta = gabarito[numeroQuestao];
        const acertou = resposta === alternativaCorreta;

        resultados.push({
          Email: email,
          Questionario: nomeAba,
          Pergunta: pergunta,
          Resposta: resposta !== null ? resposta : "Sem resposta",
          Gabarito: alternativaCorreta,
          Acertou: acertou ? "Sim" : "Não",
        });
      });
    } else {
      // Questionário 94 — cada linha tem a questão e a resposta
      dados.forEach((linha) => {
        const email =
          linha["Email"] ||
          linha["email"] ||
          linha["Aluno"] ||
          linha["aluno"] ||
          "Aluno sem e-mail";

        let questao = parseInt(linha["Questão"] || linha["Questao"]);
        if (isNaN(questao)) questao = 0;

        const resposta =
          linha["Resposta"] !== "" && !isNaN(linha["Resposta"])
            ? parseInt(linha["Resposta"])
            : null;

        const gabarito = gabaritos[nomeAba];
        const alternativaCorreta = gabarito[questao];
        const acertou = resposta === alternativaCorreta;

        resultados.push({
          Email: email,
          Questionario: nomeAba,
          Pergunta: `Q${questao + 1}`,
          Resposta: resposta !== null ? resposta : "Sem resposta",
          Gabarito: alternativaCorreta,
          Acertou: acertou ? "Sim" : "Não",
        });
      });
    }

    gerarExcel(nomeAba, resultados);
    gerarPDF(nomeAba, resultados);
  });

  console.log("✅ Tratamento finalizado.");
}

function gerarExcel(nomeAba, dados) {
  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Questionario_${nomeAba}`);

  const caminho = path.join(pastaOutput, `resultado_${nomeAba}.xlsx`);
  XLSX.writeFile(wb, caminho);
  console.log(`📊 Excel salvo: ${caminho}`);
}

function gerarPDF(nomeAba, dados) {
  const caminhoPDF = path.join(pastaOutput, `resultado_${nomeAba}.pdf`);
  const doc = new PDFDocument({ margin: 30 });
  doc.pipe(fs.createWriteStream(caminhoPDF));

  doc
    .fontSize(16)
    .text(`Relatório - Questionário ${nomeAba}`, { align: "center" })
    .moveDown();

  dados.forEach((linha) => {
    doc
      .fontSize(12)
      .text(
        `Email: ${linha.Email} | Questão: ${linha.Pergunta} | Resposta: ${linha.Resposta} | Gabarito: ${linha.Gabarito} | Acertou: ${linha.Acertou}`
      );
  });

  doc.end();
  console.log(`📄 PDF gerado: ${caminhoPDF}`);
}

// Execução
const arquivos = fs
  .readdirSync(pastaInput)
  .filter((file) => file.endsWith(".xlsx"));

if (arquivos.length === 0) {
  console.error("❌ Nenhum arquivo .xlsx encontrado na pasta input.");
} else {
  const caminho = path.join(pastaInput, arquivos[0]);
  console.log(`📥 Iniciando processamento: ${caminho}`);
  tratarPlanilha(caminho);
}
