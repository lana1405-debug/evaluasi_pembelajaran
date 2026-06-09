const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Array data kunci jawaban (20 soal)
const answerKey = ['A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'E'];

// Array data jawaban 40 responden
const students = [
  { name: "Siswa 1 (Absen 1)", scores: ['E', 'D', 'D', 'C', 'C', 'B', 'B', 'C', 'B', 'B', 'B', 'A', 'B', 'A', 'C', 'C', 'C', 'B', 'B', 'A'] },
  { name: "Siswa 2 (Absen 2)", scores: ['C', 'D', 'E', 'E', 'C', 'B', 'B', 'A', 'A', 'A', 'A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'C', 'A'] },
  { name: "Siswa 3 (Absen 3)", scores: ['E', 'D', 'D', 'C', 'C', 'B', 'B', 'C', 'B', 'B', 'B', 'A', 'B', 'A', 'C', 'C', 'C', 'D', 'D', 'E'] },
  { name: "Siswa 4 (Absen 4)", scores: ['E', 'D', 'D', 'C', 'C', 'C', 'B', 'B', 'A', 'A', 'A', 'B', 'B', 'A', 'D', 'C', 'A', 'C', 'D', 'E'] },
  { name: "Siswa 5 (Absen 5)", scores: ['E', 'A', 'C', 'C', 'A', 'D', 'D', 'E', 'E', 'E', 'B', 'A', 'B', 'A', 'D', 'C', 'A', 'C', 'D', 'E'] },
  { name: "Siswa 6 (Absen 6)", scores: ['D', 'D', 'E', 'D', 'D', 'B', 'B', 'C', 'B', 'B', 'B', 'B', 'A', 'D', 'C', 'A', 'C', 'D', 'E', 'E'] },
  { name: "Siswa 7 (Absen 7)", scores: ['C', 'D', 'E', 'E', 'C', 'B', 'B', 'A', 'A', 'C', 'A', 'B', 'B', 'A', 'D', 'C', 'A', 'C', 'C', 'A'] },
  { name: "Siswa 8 (Absen 8)", scores: ['E', 'D', 'D', 'C', 'B', 'B', 'A', 'C', 'B', 'B', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'D', 'E'] },
  { name: "Siswa 9 (Absen 9)", scores: ['D', 'A', 'C', 'C', 'A', 'D', 'D', 'E', 'E', 'E', 'A', 'B', 'B', 'C', 'A', 'C', 'A', 'D', 'E', 'D'] },
  { name: "Siswa 10 (Absen 10)", scores: ['B', 'B', 'A', 'C', 'C', 'B', 'B', 'C', 'B', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'D', 'E'] },
  { name: "Siswa 11 (Absen 11)", scores: ['D', 'E', 'E', 'A', 'A', 'C', 'C', 'A', 'A', 'D', 'A', 'A', 'A', 'B', 'B', 'B', 'C', 'A', 'B', 'D'] },
  { name: "Siswa 12 (Absen 12)", scores: ['C', 'D', 'E', 'E', 'E', 'A', 'A', 'C', 'C', 'A', 'A', 'D', 'A', 'A', 'B', 'B', 'C', 'B', 'C', 'B'] },
  { name: "Siswa 13 (Absen 13)", scores: ['D', 'D', 'E', 'E', 'A', 'A', 'C', 'A', 'A', 'A', 'C', 'A', 'A', 'D', 'B', 'B', 'C', 'A', 'B', 'C'] },
  { name: "Siswa 14 (Absen 14)", scores: ['E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'A', 'A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E'] },
  { name: "Siswa 15 (Absen 15)", scores: ['A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'A'] },
  { name: "Siswa 16 (Absen 16)", scores: ['C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'A', 'A', 'B', 'B', 'A', 'C', 'C', 'A'] },
  { name: "Siswa 17 (Absen 17)", scores: ['A', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'E', 'A', 'C', 'C', 'C', 'A', 'C', 'D', 'E', 'B', 'B'] },
  { name: "Siswa 18 (Absen 18)", scores: ['B', 'B', 'A', 'C', 'C', 'B', 'B', 'A', 'A', 'C', 'A', 'C', 'B', 'A', 'C', 'C', 'D', 'E', 'E', 'D'] },
  { name: "Siswa 19 (Absen 19)", scores: ['C', 'A', 'C', 'D', 'E', 'E', 'B', 'B', 'A', 'E', 'A', 'B', 'B', 'A', 'C', 'C', 'C', 'C', 'A', 'C'] },
  { name: "Siswa 20 (Absen 20)", scores: ['E', 'D', 'D', 'B', 'B', 'A', 'B', 'A', 'A', 'E', 'B', 'A', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D'] },
  { name: "Siswa 1 (Absen 21)", scores: ['B', 'B', 'A', 'B', 'A', 'C', 'C', 'C', 'B', 'A', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'C', 'B', 'B'] },
  { name: "Siswa 2 (Absen 22)", scores: ['A', 'A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'A', 'C', 'D', 'E', 'E', 'C', 'B', 'B', 'A', 'A', 'A'] },
  { name: "Siswa 3 (Absen 23)", scores: ['B', 'B', 'B', 'A', 'C', 'C', 'C', 'A', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'C', 'B', 'B'] },
  { name: "Siswa 4 (Absen 24)", scores: ['A', 'A', 'B', 'B', 'A', 'D', 'C', 'A', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'C', 'B', 'B', 'A', 'E'] },
  { name: "Siswa 5 (Absen 25)", scores: ['B', 'A', 'B', 'A', 'D', 'C', 'A', 'C', 'D', 'E', 'E', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'E'] },
  { name: "Siswa 6 (Absen 26)", scores: ['B', 'B', 'A', 'D', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'D', 'E', 'C', 'B', 'B', 'A', 'C', 'C'] },
  { name: "Siswa 7 (Absen 27)", scores: ['A', 'B', 'B', 'A', 'D', 'C', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'C', 'B', 'B', 'A', 'C', 'B'] },
  { name: "Siswa 8 (Absen 28)", scores: ['C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'D', 'D', 'E', 'D', 'D', 'C', 'A', 'C', 'D', 'E', 'E', 'E'] },
  { name: "Siswa 9 (Absen 29)", scores: ['A', 'B', 'B', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'B', 'B'] },
  { name: "Siswa 10 (Absen 30)", scores: ['A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'E', 'E', 'B', 'B', 'A', 'C', 'C', 'B', 'B', 'C', 'E'] },
  { name: "Siswa 11 (Absen 31)", scores: ['C', 'C', 'A', 'A', 'D', 'A', 'A', 'A', 'B', 'B', 'C', 'B', 'A', 'B', 'C', 'D', 'D', 'E', 'A', 'A'] },
  { name: "Siswa 12 (Absen 32)", scores: ['C', 'A', 'C', 'C', 'A', 'D', 'C', 'A', 'A', 'A', 'A', 'B', 'B', 'A', 'A', 'A', 'A', 'B', 'C', 'E'] },
  { name: "Siswa 13 (Absen 33)", scores: ['A', 'C', 'A', 'A', 'A', 'C', 'A', 'A', 'D', 'B', 'B', 'C', 'A', 'B', 'C', 'D', 'D', 'D', 'E', 'A'] },
  { name: "Siswa 14 (Absen 34)", scores: ['B', 'B', 'A', 'A', 'A', 'A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'A', 'A', 'B', 'B', 'A', 'C'] },
  { name: "Siswa 15 (Absen 35)", scores: ['C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'A', 'C', 'D', 'E', 'E', 'D'] },
  { name: "Siswa 16 (Absen 36)", scores: ['D', 'C', 'C', 'B', 'B', 'A', 'A', 'A', 'A', 'C', 'D', 'D', 'E', 'B', 'B', 'B', 'C', 'A', 'C', 'C'] },
  { name: "Siswa 17 (Absen 37)", scores: ['E', 'E', 'D', 'D', 'E', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'B', 'B', 'A', 'C', 'C'] },
  { name: "Siswa 18 (Absen 38)", scores: ['B', 'B', 'A', 'A', 'C', 'A', 'B', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'B', 'B', 'A', 'C', 'D'] },
  { name: "Siswa 19 (Absen 39)", scores: ['E', 'B', 'B', 'A', 'E', 'A', 'B', 'B', 'A', 'C', 'C', 'C', 'C', 'A', 'C', 'C', 'A', 'C', 'D', 'E'] },
  { name: "Siswa 20 (Absen 40)", scores: ['A', 'B', 'A', 'A', 'E', 'B', 'A', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'D', 'E', 'D', 'B', 'B'] }
];

const wb = XLSX.utils.book_new();
const wsData = [];

// Header
const header = ["No", "Nama Responden"];
for (let q = 0; q < 20; q++) {
  header.push(`S${q + 1}`);
}
header.push("Skor Total");
wsData.push(header);

// Baris Kunci
const keyRow = ["🔑", "KUNCI JAWABAN"];
answerKey.forEach(key => keyRow.push(key));
keyRow.push("-");
wsData.push(keyRow);

// Baris Siswa
students.forEach((s, idx) => {
  const row = [idx + 1, s.name];
  s.scores.forEach(score => row.push(score));
  
  const total = s.scores.reduce((sum, score, qIdx) => {
    return sum + (score.toUpperCase() === answerKey[qIdx].toUpperCase() ? 1 : 0);
  }, 0);
  row.push(total);
  wsData.push(row);
});

// Baris Jumlah Benar di paling bawah
const footer = ["", "Jumlah Benar"];
for (let q = 0; q < 20; q++) {
  let colSum = 0;
  students.forEach(s => {
    if (s.scores[q].toUpperCase() === answerKey[q].toUpperCase()) {
      colSum++;
    }
  });
  footer.push(colSum);
}
const grandTotal = students.reduce((sum, s) => {
  return sum + s.scores.reduce((rowSum, score, qIdx) => {
    return rowSum + (score.toUpperCase() === answerKey[qIdx].toUpperCase() ? 1 : 0);
  }, 0);
}, 0);
footer.push(grandTotal);
wsData.push(footer);

const ws = XLSX.utils.aoa_to_sheet(wsData);
XLSX.utils.book_append_sheet(wb, ws, "Data Butir Pilihan Ganda");

const outputPath = path.join(__dirname, 'data_analisis_soal.xlsx');
XLSX.writeFile(wb, outputPath);

console.log(`Excel file successfully generated at: ${outputPath}`);
