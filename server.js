require('dotenv').config();
const express = require('express');
const sql = require('mssql');
const path = require('path');
const ExcelJS = require('exceljs');
const app = express();

app.use(express.json());
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
    const {
      criterios,
      alternativas,
      matrizCriterios,
      pesosCriterios,
      subcriteriosPorCriterio,
      matricesSubcriterios,
      pesosSubcriteriosPorCriterio,
      matricesAlternativas,
      pesosAlternativasPorSubcriterio,
      pesosAlternativasPorCriterio,
      resultadoFinal
    } = req.body;

    const workbook = new ExcelJS.Workbook();
    const hoja = workbook.addWorksheet('AHP Completo');

    // --- MATRIZ DE CRITERIOS ---
    hoja.addRow(['Matriz de comparación de criterios']);
    hoja.addRow(['', ...criterios]);
    matrizCriterios.forEach((row, i) => {
      hoja.addRow([criterios[i], ...row]);
    });
    hoja.addRow([]);

    // Suma columnas
    const colSums = Array(criterios.length).fill(0);
    for (let j = 0; j < criterios.length; j++)
      for (let i = 0; i < criterios.length; i++)
        colSums[j] += matrizCriterios[i][j];
    hoja.addRow(['Suma columnas', ...colSums.map(x => Number(x.toFixed(4)))]);
    hoja.addRow([]);

    // Matriz normalizada
    const normMatrix = matrizCriterios.map(row =>
      row.map((val, j) => val / colSums[j])
    );
    hoja.addRow(['Matriz normalizada']);
    hoja.addRow(['', ...criterios]);
    normMatrix.forEach((row, i) => {
      hoja.addRow([criterios[i], ...row.map(x => Number(x.toFixed(4)))]);
    });
    hoja.addRow([]);

    // Pesos (vector de autoridad)
    hoja.addRow(['Pesos (vector de autoridad)', ...pesosCriterios.map(x => Number(x.toFixed(4)))]);
    hoja.addRow([]);

    // Vector Aw
    const Aw = matrizCriterios.map(row =>
      row.reduce((sum, val, j) => sum + val * pesosCriterios[j], 0)
    );
    hoja.addRow(['Vector A·w', ...Aw.map(x => Number(x.toFixed(4)))]);
    // Vector Aw/w
    const Aw_div_w = Aw.map((val, i) => val / pesosCriterios[i]);
    hoja.addRow(['Vector A·w / w', ...Aw_div_w.map(x => Number(x.toFixed(4)))]);
    hoja.addRow([]);

    // Consistencia
    const lambdaMax = Aw_div_w.reduce((a, b) => a + b, 0) / criterios.length;
    const IC = (lambdaMax - criterios.length) / (criterios.length - 1);
    const RI_VALUES = {
      1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49
    };
    const RI = RI_VALUES[criterios.length] || 1.49;
    const RC = RI === 0 ? 0 : IC / RI;
    hoja.addRow([`λ_max`, lambdaMax.toFixed(4)]);
    hoja.addRow([`Índice de Consistencia (IC)`, IC.toFixed(4)]);
    hoja.addRow([`Índice Aleatorio (RI)`, RI.toFixed(2)]);
    hoja.addRow([`Razón de Consistencia (RC)`, RC.toFixed(4)]);
    hoja.addRow([RC < 0.1 ? 'La matriz es consistente (RC < 0.1)' : 'La matriz NO es consistente (RC >= 0.1)']);
    hoja.addRow([]);

    // --- SUBCRITERIOS ---
    if (subcriteriosPorCriterio && Object.keys(subcriteriosPorCriterio).length > 0) {
      Object.entries(subcriteriosPorCriterio).forEach(([crit, subs], cidx) => {
        hoja.addRow([`Subcriterios para ${crit}:`, ...subs]);
        if (matricesSubcriterios && matricesSubcriterios[crit]) {
          hoja.addRow([`Matriz de comparación de subcriterios para ${crit}`]);
          hoja.addRow(['', ...subs]);
          matricesSubcriterios[crit].forEach((row, i) => {
            hoja.addRow([subs[i], ...row]);
          });
          hoja.addRow([]);
        }
        if (pesosSubcriteriosPorCriterio && pesosSubcriteriosPorCriterio[crit]) {
          hoja.addRow([`Pesos de subcriterios para ${crit}:`, ...pesosSubcriteriosPorCriterio[crit].map(x => Number(x.toFixed(4)))]);
          hoja.addRow([]);
        }
      });
    }

    // --- MATRICES DE ALTERNATIVAS ---
    if (subcriteriosPorCriterio && Object.keys(subcriteriosPorCriterio).length > 0 && pesosAlternativasPorSubcriterio) {
      // Con subcriterios
      Object.entries(subcriteriosPorCriterio).forEach(([crit, subs], cidx) => {
        subs.forEach((sub, sidx) => {
          hoja.addRow([`Matriz de alternativas para subcriterio: ${sub} (${crit})`]);
          hoja.addRow(['', ...alternativas]);
          if (matricesAlternativas && matricesAlternativas[`${cidx}_${sidx}`]) {
            matricesAlternativas[`${cidx}_${sidx}`].forEach((row, i) => {
              hoja.addRow([alternativas[i], ...row]);
            });
          }
          hoja.addRow([]);
          hoja.addRow([`Pesos de alternativas para subcriterio: ${sub} (${crit})`]);
          if (pesosAlternativasPorSubcriterio && pesosAlternativasPorSubcriterio[`${cidx}_${sidx}`]) {
            hoja.addRow(['Alternativa', 'Peso']);
            alternativas.forEach((alt, i) => {
              hoja.addRow([alt, Number(pesosAlternativasPorSubcriterio[`${cidx}_${sidx}`][i].toFixed(4))]);
            });
          }
          hoja.addRow([]);
        });
      });
    } else if (matricesAlternativas && Array.isArray(matricesAlternativas) && pesosAlternativasPorCriterio) {
      // Sin subcriterios
      matricesAlternativas.forEach((matriz, cIndex) => {
        hoja.addRow([`Matriz de alternativas para el criterio: ${criterios[cIndex]}`]);
        hoja.addRow(['', ...alternativas]);
        matriz.forEach((row, i) => {
          hoja.addRow([alternativas[i], ...row]);
        });
        hoja.addRow([]);
        hoja.addRow([`Pesos de alternativas para el criterio: ${criterios[cIndex]}`]);
        hoja.addRow(['Alternativa', 'Peso']);
        alternativas.forEach((alt, i) => {
          hoja.addRow([alt, Number(pesosAlternativasPorCriterio[cIndex][i].toFixed(4))]);
        });
        hoja.addRow([]);
      });
    }

    // --- RESULTADO FINAL ---
    hoja.addRow(['Resultado Final']);
    hoja.addRow(['Alternativa', 'Puntaje']);
    alternativas.forEach((alt, i) => {
      hoja.addRow([alt, Number(resultadoFinal[i].toFixed(4))]);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=AHP_Resultados.xlsx');
    res.send(buffer);

  } catch (err) {
    console.error('Error al generar Excel:', err);
    res.status(500).send('Error generando Excel');
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`✅ Servidor corriendo en el puerto ${PORT}`);
});

