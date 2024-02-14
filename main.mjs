import XLSX from 'xlsx';

function readWorkbook(filePath) {
  console.log('Leitura do arquivo')
  return XLSX.readFile(filePath);
}

function calculateAverage(grades) {

  const totalGrades = grades.reduce((acc, grade) => acc + grade / 10, 0);
  return totalGrades / grades.length;
}

function determineStatus(average, absences) {

  if (absences > 15) {
    return 'Reprovado por Falta';
  } else if (average >= 7) {
    return 'Aprovado';
  } else if (average < 5) {
    return 'Reprovado';
  } else {
    return 'Exame final';
  }
}

function calcularNAF(average) {

  average = Math.max(0, Math.min(10, average));
  const naf = Math.max(0, 10 - average);
  return naf.toFixed(1);
}

function processStudentData(sheet, studentData, i) {
  const { enrollment, studentName, absences, grades } = studentData;

  console.log("Calculando médias")
  const average = calculateAverage(grades);
  
  console.log('Definindo condições de desempenho')
  const status = determineStatus(average, absences);
  sheet[`G${i}`] = { v: status, t: 's' };

  if (status === 'Exame final') {
    console.log('Calculando nota para alcançar aprovação')
    sheet[`H${i}`] = { v: calcularNAF(average), t: 'n' };
  } else {
    sheet[`H${i}`] = { v: 0, t: 'n' };
  }
}
function saveWorkbook(workbook, outputPath) {
  console.log('Salvando arquivo')
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx', bookSST: false, type: 'binary' });
}

// Main function
function main() {
  const filePath = 'Engenharia de Software.xlsx';
  const outputPath = 'Engenharia de Software - Atualizada.xlsx';

  try {
    const workbook = readWorkbook(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    for (let i = 4; sheet[`A${i}`]; i++) {
      const studentData = {
        enrollment: sheet[`A${i}`].v,
        studentName: sheet[`B${i}`].v,
        absences: sheet[`C${i}`].v,
        grades: [sheet[`D${i}`].v, sheet[`E${i}`].v, sheet[`F${i}`].v],
      };

      processStudentData(sheet, studentData, i);
    }

    

    saveWorkbook(workbook, outputPath);
    console.log('Médias calculadas e nova planilha salva.');
  } catch (error) {
    console.error('Erro ao processar a planilha:', error);
  }
}

main();