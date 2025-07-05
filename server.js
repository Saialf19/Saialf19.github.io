require('dotenv').config();
const express = require('express');
const sql = require('mssql');
const path = require('path');
const ExcelJS = require('exceljs');
const app = express();

app.use(express.json({ limit: '10mb' })); // Aumentar límite
app.use(express.static(path.join(__dirname, 'public')));

const dbConfig = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER || 'localhost',
  database: process.env.DB_NAME,
  options: { trustServerCertificate: true }
};

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/guardarDatosDinamicos', async (req, res) => {
  try {
    await sql.connect(dbConfig);
    const body = req.body;
    const criterios = Object.entries(body).filter(([k]) => k.startsWith('criterio')).map(([_, v]) => v);
    const alternativas = Object.entries(body).filter(([k]) => k.startsWith('alternativa')).map(([_, v]) => v);

    for (let crit of criterios) {
      const check = await sql.query`SELECT COUNT(*) AS count FROM Criterios WHERE nombre = ${crit}`;
      if (check.recordset[0].count === 0) {
        await sql.query`INSERT INTO Criterios(nombre) VALUES (${crit})`;
      }
    }

    for (let alt of alternativas) {
      const check = await sql.query`SELECT COUNT(*) AS count FROM Alternativas WHERE nombre = ${alt}`;
      if (check.recordset[0].count === 0) {
        await sql.query`INSERT INTO Alternativas(nombre) VALUES (${alt})`;
      }
    }

    res.send("Datos dinámicos guardados correctamente.");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error al guardar datos dinámicos.");
  }
});

app.post('/exportarAHP', async (req, res) => {
  try {
    // Validar que se reciban todos los datos necesarios
    const {
      criterios,
      alternativas,
      matrizCriterios,
      pesosCriterios,
      resultadoFinal,
      pesosAlternativasPorCriterio
    } = req.body;

    if (!criterios || !alternativas || !matrizCriterios || !pesosCriterios || !resultadoFinal) {
      return res.status(400).send("Faltan datos esenciales.");
    }

    const workbook = new ExcelJS.Workbook();
    // Solo UNA hoja para todo
    const hoja = workbook.addWorksheet('AHP Dinámico');

    // 1) Matriz de comparación de criterios
    hoja.addRow(['Matriz de comparación de criterios']);
    hoja.addRow(['', ...criterios]);
    matrizCriterios.forEach((fila, i) => {
      hoja.addRow([criterios[i], ...fila]);
    });

    // 2) Suma columnas
    const sumaCols = criterios.map((_, j) =>
      matrizCriterios.reduce((sum, fila) => sum + fila[j], 0)
    );
    hoja.addRow([]);
    hoja.addRow(['Suma columnas', ...sumaCols.map(x => Number(x.toFixed(4)))]);

    // 3) Matriz normalizada
    hoja.addRow([]);
    hoja.addRow(['Matriz normalizada']);
    hoja.addRow(['', ...criterios]);
    const matrizNorm = matrizCriterios.map(fila =>
      fila.map((v, j) => v / sumaCols[j])
    );
    matrizNorm.forEach((fila, i) => {
      hoja.addRow([criterios[i], ...fila.map(x => Number(x.toFixed(4)))]);
    });

    // 4) Pesos (vector de autoridad)
    hoja.addRow([]);
    hoja.addRow(['Pesos (vector de autoridad)', ...pesosCriterios.map(x => Number(x.toFixed(4)))]);

    // 5) Vector A·w
    const Aw = matrizCriterios.map(fila =>
      fila.reduce((sum, v, j) => sum + v * pesosCriterios[j], 0)
    );
    hoja.addRow([]);
    hoja.addRow(['Vector A·w', ...Aw.map(x => Number(x.toFixed(4)))]);

    // 6) Vector A·w / w
    const AwDivW = Aw.map((v, i) => v / pesosCriterios[i]);
    hoja.addRow(['Vector A·w / w', ...AwDivW.map(x => Number(x.toFixed(4)))]);

    // 7) Consistencia
    const lambdaMax = AwDivW.reduce((a,b) => a + b, 0) / criterios.length;
    const IC = (lambdaMax - criterios.length) / (criterios.length - 1);
    const RI_VALUES = {1:0,2:0,3:0.58,4:0.90,5:1.12,6:1.24,7:1.32,8:1.41,9:1.45};
    const RI = RI_VALUES[criterios.length] ?? 1.49;
    const RC = RI === 0 ? 0 : IC / RI;
    hoja.addRow([]);
    hoja.addRow(['λ_max', Number(lambdaMax.toFixed(4))]);
    hoja.addRow(['Índice de Consistencia (IC)', Number(IC.toFixed(4))]);
    hoja.addRow(['Índice Aleatorio (RI)', Number(RI.toFixed(2))]);
    hoja.addRow(['Razón de Consistencia (RC)', Number(RC.toFixed(4))]);
    hoja.addRow([RC < 0.1
      ? 'La matriz es consistente (RC < 0.1)'
      : 'La matriz NO es consistente (RC >= 0.1)']);

    // 8) Pesos de alternativas por criterio
    hoja.addRow([]);
    Object.entries(pesosAlternativasPorCriterio).forEach(([crit, arrPesos]) => {
      hoja.addRow([`Pesos de alternativas para el criterio: ${crit}`]);
      hoja.addRow(['Alternativa', 'Peso']);
      alternativas.forEach((alt, i) => {
        hoja.addRow([alt, Number(arrPesos[i].toFixed(4))]);
      });
      hoja.addRow([]);
    });

    // 9) Resultado final
    hoja.addRow(['Resultado Final']);
    hoja.addRow(['Alternativa', 'Puntaje']);
    alternativas.forEach((alt, i) => {
      hoja.addRow([alt, Number(resultadoFinal[i].toFixed(4))]);
    });

    // Devolver XLSX
    const buffer = await workbook.xlsx.writeBuffer();
    res
      .header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      .header('Content-Disposition', 'attachment; filename=AHP_Dinamico.xlsx')
      .send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Error generando Excel dinámico");
  }
});


const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`✅ Servidor corriendo en el puerto ${PORT}`);
});

