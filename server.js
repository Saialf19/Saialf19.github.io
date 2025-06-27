const express = require('express');
const sql = require('mssql');
const path = require('path');
const ExcelJS = require('exceljs');
const app = express();

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const dbConfig = {
  user: 'sa',
  password: 'TuContraseñaSQL',
  server: 'localhost',
  database: 'AHPDB',
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
      matrizAlternativas,
      pesosCriterios,
      resultadoFinal
    } = req.body;

    const workbook = new ExcelJS.Workbook();
    const hoja = workbook.addWorksheet('AHP Completo');

    // Título matriz de criterios
    hoja.addRow(['Matriz de comparación de criterios']);
    hoja.addRow(['', ...criterios]);
    matrizCriterios.forEach((row, i) => {
      hoja.addRow([criterios[i], ...row]);
    });
    hoja.addRow([]);

    // Título pesos criterios
    hoja.addRow(['Pesos de Criterios']);
    hoja.addRow(['Criterio', 'Peso']);
    criterios.forEach((c, i) => {
      hoja.addRow([c, pesosCriterios[i]]);
    });
    hoja.addRow([]);

    // Título matrices alternativas
    matrizAlternativas.forEach((pesos, cIndex) => {
      hoja.addRow([`Pesos de alternativas para el criterio: ${criterios[cIndex]}`]);
      hoja.addRow(['Alternativa', 'Peso']);
      alternativas.forEach((alt, i) => {
        hoja.addRow([alt, pesos[i]]);
      });
      hoja.addRow([]);
    });

    // Resultado Final
    hoja.addRow(['Resultado Final']);
    hoja.addRow(['Alternativa', 'Puntaje']);
    alternativas.forEach((alt, i) => {
      hoja.addRow([alt, resultadoFinal[i]]);
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

app.listen(4000, () => console.log('✅ Servidor corriendo en http://localhost:4000'));
