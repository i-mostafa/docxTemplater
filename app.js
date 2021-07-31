const { createReport } = require("docx-templates");
const xlsx = require("node-xlsx");

const fs = require("fs");
const template = fs.readFileSync("input/template.docx");
const workSheetsFromFile = xlsx.parse(`input/input.xlsx`);
const keys = workSheetsFromFile[0].data[0];

const data = [];
const dataArray = workSheetsFromFile[0].data;
for (let i = 1; i < dataArray.length; i++) {
  const patient = {};

  keys.forEach((key, j) => {
    patient[key] = dataArray[i][j] || "";
  });
  data.push(patient);
}

data.forEach(async (patient, i) => {
  const buffer = await createReport({
    template,
    data: patient,
  });
  fs.writeFileSync(
    `output/${patient.patientName}-${patient.accessionNumber}.docx`,
    buffer
  );
});
