const { createReport } = require("docx-templates");
const xlsx = require("node-xlsx");
const fs = require("fs");
const { PDFNet } = require("@pdftron/pdfnet-node");

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
  patient.patientName = patient.patientName.toUpperCase();
  patient.result = patient.result.toUpperCase();

  await PDFNet.initialize();
  const buffer = await createReport({
    template,
    data: patient,
  });
  fs.writeFileSync(
    `output/${patient.patientName}-${patient.accessionNumber}.docx`,
    buffer
  );
  // toPdf(
  //   `output/${patient.patientName}-${patient.accessionNumber}.docx`,
  //   `output/${patient.patientName}-${patient.accessionNumber}.pdf`
  // );

  // perform the conversion with no optional parameters
  const pdfdoc = await PDFNet.Convert.officeToPdfWithPath(
    `output/${patient.patientName}-${patient.accessionNumber}.docx`
  );

  // save the result
  await pdfdoc.save(
    `output/${patient.patientName}-${patient.accessionNumber}.pdf`,
    PDFNet.SDFDoc.SaveOptions.e_linearized
  );
  if (Object.is(data.length - 1, i)) {
    process.exit();
  }
});
