/**
 * Aplikasi Analisis Butir Soal (Arikunto & Sugiyono)
 * Logika Perhitungan Statistik Pilihan Ganda (A-E), Kunci Jawaban, State Management, dan Kontrol UI
 */

// State Aplikasi
let state = {
  students: [],
  answerKey: ['A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'E'], 
  questionsCount: 20,
  studentsCount: 40,
  significanceLevel: 0.05, // 0.05 atau 0.01
  activeTab: 'input-data', // 'input-data', 'hasil-analisis', 'cara-perhitungan', 'cara-reliabilitas'
  selectedQuestionIndex: 0 // Index butir soal yang aktif di detail perhitungan
};

// Konstanta Nama Default
const DEFAULT_STUDENT_NAMES = [];

// Inisialisasi saat halaman dimuat
document.addEventListener('DOMContentLoaded', () => {
  initApp();
  setupEventListeners();
  loadSampleData(); // Muat contoh data secara default untuk memberikan kesan pertama yang bagus
});

/**
 * Inisialisasi State Awal
 */
function initApp() {
  if (state.students.length === 0) {
    generateDefaultData(state.studentsCount, state.questionsCount);
  }
  renderGrid();
  updateGridInputs();
}

/**
 * Membuat data kosong awal
 */
function generateDefaultData(numStudents, numQuestions) {
  state.answerKey = ['A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'E'];
  state.students = [];
  for (let i = 0; i < numStudents; i++) {
    // Penamaan generik seperti Siswa 1 (Absen 1) dst
    const absenNum = i + 1;
    const siswaNum = (i % 20) + 1;
    const name = `Siswa ${siswaNum} (Absen ${absenNum})`;
    const scores = Array(numQuestions).fill(''); 
    state.students.push({
      id: i + 1,
      name: name,
      scores: scores
    });
  }
}

/**
 * Mendaftarkan Event Listeners
 */
function setupEventListeners() {
  // Navigasi Tab Utama
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
      const tabId = item.getAttribute('data-tab');
      switchTab(tabId);
    });
  });

  // Toggle Mode Gelap/Terang
  const themeToggle = document.getElementById('theme-toggle');
  themeToggle.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    const sunIcon = themeToggle.querySelector('.sun-icon');
    const moonIcon = themeToggle.querySelector('.moon-icon');
    if (newTheme === 'dark') {
      sunIcon.style.display = 'none';
      moonIcon.style.display = 'block';
    } else {
      sunIcon.style.display = 'block';
      moonIcon.style.display = 'none';
    }
  });

  // Tombol Kontrol Grid
  document.getElementById('btn-add-row').addEventListener('click', addStudentRow);
  document.getElementById('btn-remove-row').addEventListener('click', removeStudentRow);
  document.getElementById('btn-add-col').addEventListener('click', addQuestionCol);
  document.getElementById('btn-remove-col').addEventListener('click', removeQuestionCol);

  // Setup Signifikansi Level
  const sigSelect = document.getElementById('select-significance');
  if (sigSelect) {
    sigSelect.addEventListener('change', (e) => {
      state.significanceLevel = parseFloat(e.target.value);
      if (state.activeTab !== 'input-data') {
        runAnalysis();
      }
    });
  }

  // Tombol Load Contoh Data
  document.getElementById('btn-load-sample').addEventListener('click', loadSampleData);
  document.getElementById('btn-reset').addEventListener('click', resetToEmpty);

  // Modal Impor Excel
  const importModal = document.getElementById('import-modal');
  const btnOpenImport = document.getElementById('btn-open-import');
  const btnCloseImport = importModal.querySelector('.close-modal-btn');
  const btnCancelImport = document.getElementById('btn-cancel-import');
  const fileInput = document.getElementById('excel-file-input');
  const dropzone = document.getElementById('dropzone');

  btnOpenImport.addEventListener('click', () => importModal.classList.add('active'));
  btnCloseImport.addEventListener('click', () => importModal.classList.remove('active'));
  btnCancelImport.addEventListener('click', () => importModal.classList.remove('active'));

  // Drag and Drop Impor
  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = 'var(--primary)';
    dropzone.style.backgroundColor = 'var(--primary-light)';
  });
  dropzone.addEventListener('dragleave', () => {
    dropzone.style.borderColor = 'var(--border-color)';
    dropzone.style.backgroundColor = 'transparent';
  });
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = 'var(--border-color)';
    dropzone.style.backgroundColor = 'transparent';
    if (e.dataTransfer.files.length > 0) {
      handleImportFile(e.dataTransfer.files[0]);
    }
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleImportFile(e.target.files[0]);
    }
  });

  // Tombol Ekspor
  document.getElementById('btn-export-excel').addEventListener('click', exportToExcel);
  document.getElementById('btn-export-pdf').addEventListener('click', exportToPDF);
}

/**
 * Mengganti Tab Aktif
 */
function switchTab(tabId) {
  state.activeTab = tabId;
  document.querySelectorAll('.nav-item').forEach(item => {
    if (item.getAttribute('data-tab') === tabId) {
      item.classList.add('active');
    } else {
      item.classList.remove('active');
    }
  });

  document.querySelectorAll('.view-panel').forEach(panel => {
    if (panel.id === tabId) {
      panel.classList.add('active');
    } else {
      panel.classList.remove('active');
    }
  });

  // Jika masuk ke tab analisis, jalankan perhitungan statistik
  if (tabId === 'hasil-analisis' || tabId === 'cara-perhitungan' || tabId === 'cara-reliabilitas') {
    runAnalysis();
  }
}

/**
 * Render Grid Spreadsheet Input Data
 */
function renderGrid() {
  const tableHeadRow = document.getElementById('sheet-thead-row');
  const tableBody = document.getElementById('sheet-tbody');
  
  // Clear existing
  tableHeadRow.innerHTML = '';
  tableBody.innerHTML = '';

  // 1. Render Header (No, Nama, Soal 1 s.d. Soal K, Total Skor)
  const thNo = document.createElement('th');
  thNo.textContent = 'No';
  thNo.style.width = '50px';
  tableHeadRow.appendChild(thNo);

  const thNama = document.createElement('th');
  thNama.textContent = 'Nama Responden';
  thNama.className = 'col-sticky';
  tableHeadRow.appendChild(thNama);

  for (let q = 0; q < state.questionsCount; q++) {
    const thQ = document.createElement('th');
    thQ.textContent = `S${q + 1}`;
    thQ.title = `Butir Soal Ke-${q + 1}`;
    tableHeadRow.appendChild(thQ);
  }

  const thTotal = document.createElement('th');
  thTotal.textContent = 'Skor';
  thTotal.title = 'Total Skor Jawaban Benar';
  thTotal.style.width = '70px';
  tableHeadRow.appendChild(thTotal);

  // 2. Render Baris Khusus KUNCI JAWABAN (Posisi di paling atas tbody)
  const trKey = document.createElement('tr');
  trKey.className = 'row-key';

  const tdKeyNo = document.createElement('td');
  tdKeyNo.textContent = '🔑';
  trKey.appendChild(tdKeyNo);

  const tdKeyNama = document.createElement('td');
  tdKeyNama.className = 'col-sticky';
  tdKeyNama.textContent = 'KUNCI JAWABAN';
  trKey.appendChild(tdKeyNama);

  for (let q = 0; q < state.questionsCount; q++) {
    const tdKeyQ = document.createElement('td');
    const inputKeyQ = document.createElement('input');
    inputKeyQ.type = 'text';
    inputKeyQ.maxLength = 1;
    inputKeyQ.value = state.answerKey[q] || 'A';
    inputKeyQ.dataset.keyQuestionIndex = q;

    // Event Kunci Jawaban Diubah
    inputKeyQ.addEventListener('input', (e) => {
      let val = e.target.value.trim().toUpperCase();
      if (val !== 'A' && val !== 'B' && val !== 'C' && val !== 'D' && val !== 'E' && val !== '') {
        inputKeyQ.classList.add('invalid');
      } else {
        inputKeyQ.classList.remove('invalid');
        state.answerKey[q] = val || 'A';
        // Render ulang grid visual warna sel siswa karena kunci berubah
        updateAllStudentCellVisuals();
      }
    });

    tdKeyQ.appendChild(inputKeyQ);
    trKey.appendChild(tdKeyQ);
  }

  const tdKeyTotal = document.createElement('td');
  tdKeyTotal.textContent = '-';
  tdKeyTotal.style.fontWeight = 'bold';
  trKey.appendChild(tdKeyTotal);

  tableBody.appendChild(trKey);

  // 3. Render Baris Siswa (Responden)
  state.students.forEach((student, sIdx) => {
    const tr = document.createElement('tr');

    // Kolom No
    const tdNo = document.createElement('td');
    tdNo.textContent = sIdx + 1;
    tr.appendChild(tdNo);

    // Kolom Nama (Sticky)
    const tdNama = document.createElement('td');
    tdNama.className = 'col-sticky';
    const inputNama = document.createElement('input');
    inputNama.type = 'text';
    inputNama.value = student.name;
    inputNama.style.textAlign = 'left';
    inputNama.addEventListener('change', (e) => {
      student.name = e.target.value;
    });
    tdNama.appendChild(inputNama);
    tr.appendChild(tdNama);

    // Kolom Jawaban Huruf (A-E)
    for (let q = 0; q < state.questionsCount; q++) {
      const tdQ = document.createElement('td');
      const inputQ = document.createElement('input');
      inputQ.type = 'text';
      inputQ.maxLength = 1;
      inputQ.value = student.scores[q] !== undefined ? student.scores[q] : '';
      inputQ.dataset.studentIndex = sIdx;
      inputQ.dataset.questionIndex = q;

      // Berikan kelas visual awal (Benar/Salah)
      updateCellVisual(tdQ, inputQ.value, q);

      // Event listener input & validasi
      inputQ.addEventListener('input', (e) => {
        let val = e.target.value.trim().toUpperCase();
        if (val !== 'A' && val !== 'B' && val !== 'C' && val !== 'D' && val !== 'E' && val !== '') {
          inputQ.classList.add('invalid');
        } else {
          inputQ.classList.remove('invalid');
          student.scores[q] = val;
          // Perbarui style warna sel saat ini
          updateCellVisual(tdQ, val, q);
          // Perbarui total skor responden tersebut
          updateStudentTotalScore(sIdx);
        }
      });

      // Menjadikan navigasi panah keyboard lebih mudah di spreadsheet
      inputQ.addEventListener('keydown', (e) => {
        handleGridKeyboardNavigation(e, sIdx, q);
      });

      tdQ.appendChild(inputQ);
      tr.appendChild(tdQ);
    }

    // Kolom Total Skor (Ditampilkan otomatis)
    const tdTotal = document.createElement('td');
    tdTotal.id = `total-score-${sIdx}`;
    tdTotal.style.fontWeight = 'bold';
    tdTotal.textContent = calculateStudentEvaluatedTotal(student);
    tr.appendChild(tdTotal);

    tableBody.appendChild(tr);
  });

  // 4. Render Baris Jumlah Benar di Paling Bawah
  const trTotalBenar = document.createElement('tr');
  trTotalBenar.className = 'total-row';

  const tdTotalLabelNo = document.createElement('td');
  tdTotalLabelNo.textContent = '';
  trTotalBenar.appendChild(tdTotalLabelNo);

  const tdTotalLabelNama = document.createElement('td');
  tdTotalLabelNama.className = 'col-sticky';
  tdTotalLabelNama.textContent = 'Jumlah Benar';
  trTotalBenar.appendChild(tdTotalLabelNama);

  for (let q = 0; q < state.questionsCount; q++) {
    const tdSumQ = document.createElement('td');
    tdSumQ.id = `sum-correct-q-${q}`;
    tdSumQ.textContent = calculateSumCorrect(q);
    trTotalBenar.appendChild(tdSumQ);
  }

  const tdGrandTotal = document.createElement('td');
  tdGrandTotal.id = 'grand-total-scores';
  tdGrandTotal.textContent = calculateGrandTotal();
  trTotalBenar.appendChild(tdGrandTotal);

  tableBody.appendChild(trTotalBenar);
}

/**
 * Memperbarui kelas warna td (hijau untuk benar, merah tipis untuk salah)
 */
function updateCellVisual(tdCell, studentAns, qIdx) {
  const correctKey = state.answerKey[qIdx] || 'A';
  tdCell.classList.remove('cell-correct', 'cell-incorrect');
  
  if (studentAns === '') {
    // Kosong / belum menjawab, tidak diwarnai khusus
    return;
  }

  if (studentAns.toUpperCase() === correctKey.toUpperCase()) {
    tdCell.classList.add('cell-correct');
  } else {
    tdCell.classList.add('cell-incorrect');
  }
}

/**
 * Memperbarui semua visual sel setelah Kunci Jawaban diubah
 */
function updateAllStudentCellVisuals() {
  state.students.forEach((student, sIdx) => {
    for (let q = 0; q < state.questionsCount; q++) {
      const inputQ = document.querySelector(`input[data-student-index="${sIdx}"][data-question-index="${q}"]`);
      if (inputQ) {
        const tdQ = inputQ.parentElement;
        updateCellVisual(tdQ, student.scores[q] || '', q);
      }
    }
    // Update total skor di kolom samping
    const totalTd = document.getElementById(`total-score-${sIdx}`);
    if (totalTd) {
      totalTd.textContent = calculateStudentEvaluatedTotal(student);
    }
  });

  // Update baris Jumlah Benar di paling bawah
  for (let q = 0; q < state.questionsCount; q++) {
    const sumTd = document.getElementById(`sum-correct-q-${q}`);
    if (sumTd) {
      sumTd.textContent = calculateSumCorrect(q);
    }
  }

  // Update grand total
  const grandTd = document.getElementById('grand-total-scores');
  if (grandTd) {
    grandTd.textContent = calculateGrandTotal();
  }
}

/**
 * Keyboard Navigation di Spreadsheet
 */
function handleGridKeyboardNavigation(e, sIdx, qIdx) {
  let targetInput = null;
  if (e.key === 'ArrowRight' && qIdx < state.questionsCount - 1) {
    targetInput = document.querySelector(`input[data-student-index="${sIdx}"][data-question-index="${qIdx + 1}"]`);
  } else if (e.key === 'ArrowLeft' && qIdx > 0) {
    targetInput = document.querySelector(`input[data-student-index="${sIdx}"][data-question-index="${qIdx - 1}"]`);
  } else if (e.key === 'ArrowDown' && sIdx < state.students.length - 1) {
    targetInput = document.querySelector(`input[data-student-index="${sIdx + 1}"][data-question-index="${qIdx}"]`);
  } else if (e.key === 'ArrowUp') {
    if (sIdx === 0) {
      // Pindah ke baris Kunci Jawaban jika di baris siswa pertama dan tekan panah atas
      targetInput = document.querySelector(`input[data-key-question-index="${qIdx}"]`);
    } else {
      targetInput = document.querySelector(`input[data-student-index="${sIdx - 1}"][data-question-index="${qIdx}"]`);
    }
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (sIdx < state.students.length - 1) {
      targetInput = document.querySelector(`input[data-student-index="${sIdx + 1}"][data-question-index="${qIdx}"]`);
    }
  }

  // Tambahkan navigasi jika dari kunci jawaban ditekan panah bawah
  if (e.target.dataset.keyQuestionIndex !== undefined && e.key === 'ArrowDown') {
    const keyQIdx = parseInt(e.target.dataset.keyQuestionIndex);
    targetInput = document.querySelector(`input[data-student-index="0"][data-question-index="${keyQIdx}"]`);
  }

  if (targetInput) {
    targetInput.focus();
    targetInput.select();
  }
}

// Tambahkan navigasi panah keyboard khusus untuk baris kunci jawaban
document.addEventListener('keydown', (e) => {
  if (e.target.dataset.keyQuestionIndex !== undefined) {
    const qIdx = parseInt(e.target.dataset.keyQuestionIndex);
    let targetInput = null;
    if (e.key === 'ArrowRight' && qIdx < state.questionsCount - 1) {
      targetInput = document.querySelector(`input[data-key-question-index="${qIdx + 1}"]`);
    } else if (e.key === 'ArrowLeft' && qIdx > 0) {
      targetInput = document.querySelector(`input[data-key-question-index="${qIdx - 1}"]`);
    } else if (e.key === 'ArrowDown') {
      targetInput = document.querySelector(`input[data-student-index="0"][data-question-index="${qIdx}"]`);
    }
    
    if (targetInput) {
      targetInput.focus();
      targetInput.select();
    }
  }
});

/**
 * Sinkronisasi jumlah input kontrol dengan state
 */
function updateGridInputs() {
  document.getElementById('input-students-count').value = state.studentsCount;
  document.getElementById('input-questions-count').value = state.questionsCount;
}

/**
 * Menghitung skor total responden yang telah dibandingkan dengan kunci jawaban (Biner)
 */
function calculateStudentEvaluatedTotal(student) {
  return student.scores.reduce((sum, score, idx) => {
    const key = state.answerKey[idx] || 'A';
    const isCorrect = score && score.toUpperCase() === key.toUpperCase();
    return sum + (isCorrect ? 1 : 0);
  }, 0);
}

/**
 * Mengupdate skor total satu siswa di UI
 */
function updateStudentTotalScore(sIdx) {
  const student = state.students[sIdx];
  const totalTd = document.getElementById(`total-score-${sIdx}`);
  if (totalTd) {
    totalTd.textContent = calculateStudentEvaluatedTotal(student);
  }

  // Update jumlah benar per butir soal
  for (let q = 0; q < state.questionsCount; q++) {
    const sumTd = document.getElementById(`sum-correct-q-${q}`);
    if (sumTd) {
      sumTd.textContent = calculateSumCorrect(q);
    }
  }

  // Update grand total
  const grandTd = document.getElementById('grand-total-scores');
  if (grandTd) {
    grandTd.textContent = calculateGrandTotal();
  }
}

/**
 * Menghitung total jawaban benar untuk butir soal q (mencocokkan jawaban dengan kunci)
 */
function calculateSumCorrect(qIdx) {
  const key = state.answerKey[qIdx] || 'A';
  return state.students.reduce((sum, student) => {
    const score = student.scores[qIdx] || '';
    const isCorrect = score && score.toUpperCase() === key.toUpperCase();
    return sum + (isCorrect ? 1 : 0);
  }, 0);
}

/**
 * Menghitung total skor seluruh responden (jumlah jawaban benar gabungan)
 */
function calculateGrandTotal() {
  return state.students.reduce((sum, student) => sum + calculateStudentEvaluatedTotal(student), 0);
}

/**
 * Fungsi Kontrol Grid: Tambah Siswa
 */
function addStudentRow() {
  state.studentsCount++;
  const name = `Siswa ${state.studentsCount}`;
  state.students.push({
    id: state.studentsCount,
    name: name,
    scores: Array(state.questionsCount).fill('')
  });
  renderGrid();
  updateGridInputs();
}

/**
 * Fungsi Kontrol Grid: Hapus Siswa
 */
function removeStudentRow() {
  if (state.students.length <= 3) {
    alert("Jumlah siswa minimal adalah 3 responden untuk keperluan statistik.");
    return;
  }
  state.studentsCount--;
  state.students.pop();
  renderGrid();
  updateGridInputs();
}

/**
 * Fungsi Kontrol Grid: Tambah Soal
 */
function addQuestionCol() {
  state.questionsCount++;
  state.answerKey.push('A');
  state.students.forEach(student => {
    student.scores.push('');
  });
  renderGrid();
  updateGridInputs();
}

/**
 * Fungsi Kontrol Grid: Hapus Soal
 */
function removeQuestionCol() {
  if (state.questionsCount <= 2) {
    alert("Jumlah soal minimal adalah 2 butir.");
    return;
  }
  state.questionsCount--;
  state.answerKey.pop();
  state.students.forEach(student => {
    student.scores.pop();
  });
  renderGrid();
  updateGridInputs();
}

/**
 * Reset grid ke nilai kosong semua
 */
function resetToEmpty() {
  if (confirm("Apakah Anda yakin ingin mengosongkan seluruh jawaban responden?")) {
    state.students.forEach(student => {
      student.scores = Array(state.questionsCount).fill('');
    });
    renderGrid();
  }
}

/**
 * Memuat Contoh Data Riil Pilihan Ganda (A-E)
 */
function loadSampleData() {
  state.questionsCount = 20;
  state.studentsCount = 40;

  // Set kunci jawaban A-E
  state.answerKey = ['A', 'B', 'B', 'A', 'C', 'C', 'A', 'C', 'D', 'E', 'E', 'D', 'D', 'C', 'C', 'B', 'B', 'A', 'A', 'E'];

  const sampleStudents = [
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

  state.students = sampleStudents.map((s, idx) => {
    return {
      id: idx + 1,
      name: s.name,
      scores: [...s.scores]
    };
  });

  renderGrid();
  updateGridInputs();
}

/**
 * ==========================================
 * LOGIKA MATEMATIKA ANALISIS STATISTIK
 * ==========================================
 */

let analysisResults = {
  items: [],
  reliability: {
    k: 0,
    sumPq: 0,
    varianceTotal: 0,
    r11: 0,
    category: "",
    status: ""
  }
};

/**
 * Jalankan Analisis Lengkap
 */
function runAnalysis() {
  const N = state.students.length;
  const k = state.questionsCount;

  if (N < 3) {
    alert("Responden terlalu sedikit untuk melakukan analisis statistik (min. 3)");
    return;
  }

  // --- LANGKAH PENTING ---
  // Evaluasi lembar jawaban A-E siswa menjadi biner (0 atau 1) secara internal
  // sebelum melakukan analisis statistik, sehingga rumus tetap presisi!
  const binaryStudents = state.students.map(student => {
    return {
      id: student.id,
      name: student.name,
      scores: student.scores.map((score, idx) => {
        const key = state.answerKey[idx] || 'A';
        return (score && score.toUpperCase() === key.toUpperCase()) ? 1 : 0;
      })
    };
  });

  // Hitung Skor Total Siswa (Biner)
  const totalScores = binaryStudents.map(s => s.scores.reduce((a, b) => a + b, 0));
  const sumTotalScores = totalScores.reduce((a, b) => a + b, 0);
  const sumTotalScoresSq = totalScores.reduce((sum, score) => sum + Math.pow(score, 2), 0);

  // Varians Total (Rumus Arikunto): S_t^2 = (sum Y^2 - (sum Y)^2 / N) / N
  const varianceTotal = (sumTotalScoresSq - (Math.pow(sumTotalScores, 2) / N)) / N;

  // Siapkan penampung hasil analisis butir
  analysisResults.items = [];
  let sumPq = 0;

  // Analisis Per Butir Soal (qIdx)
  for (let q = 0; q < k; q++) {
    // 1. Ambil array skor biner X untuk semua responden pada butir q
    const itemScores = binaryStudents.map(s => s.scores[q] || 0);
    const sumX = itemScores.reduce((a, b) => a + b, 0);
    const sumXSq = itemScores.reduce((sum, x) => sum + Math.pow(x, 2), 0); // karena dikotomi, sum X^2 = sum X

    // Hitung perkalian X * Y
    let sumXY = 0;
    for (let i = 0; i < N; i++) {
      sumXY += itemScores[i] * totalScores[i];
    }

    // 2. Korelasi Product Moment Pearson (Validitas r_xy)
    let r_xy = 0;
    const numerator = (N * sumXY) - (sumX * sumTotalScores);
    const denominator = Math.sqrt(
      ((N * sumXSq) - Math.pow(sumX, 2)) * 
      ((N * sumTotalScoresSq) - Math.pow(sumTotalScores, 2))
    );

    if (denominator !== 0) {
      r_xy = numerator / denominator;
    } else {
      r_xy = 0;
    }

    // Ambil r_tabel Pearson
    const df = N - 2;
    const r_tabel = window.rTable.get(df, state.significanceLevel);
    const isValid = r_xy > r_tabel;

    // 3. Tingkat Kesukaran (P)
    const B = sumX; // Jumlah siswa menjawab benar
    const P = B / N;
    const q_val = 1 - P;
    sumPq += P * q_val;

    let diffCategory = "";
    if (P < 0.30) diffCategory = "Sukar";
    else if (P <= 0.70) diffCategory = "Sedang";
    else diffCategory = "Mudah";

    // 4. Daya Pembeda (D)
    // Urutkan siswa berdasarkan skor total dari tinggi ke rendah untuk membagi kelompok atas & bawah
    const sortedStudents = binaryStudents.map((s, idx) => ({
      id: s.id,
      name: s.name,
      total: totalScores[idx],
      scoreOnItem: s.scores[q] || 0
    })).sort((a, b) => b.total - a.total);

    let J_A = 0;
    let J_B = 0;
    let upperGroup = [];
    let lowerGroup = [];

    if (N <= 30) {
      const halfSize = Math.floor(N / 2);
      J_A = halfSize;
      J_B = halfSize;
      upperGroup = sortedStudents.slice(0, halfSize);
      lowerGroup = sortedStudents.slice(sortedStudents.length - halfSize);
    } else {
      const groupSize = Math.round(0.27 * N);
      J_A = groupSize;
      J_B = groupSize;
      upperGroup = sortedStudents.slice(0, groupSize);
      lowerGroup = sortedStudents.slice(sortedStudents.length - groupSize);
    }

    const B_A = upperGroup.reduce((sum, s) => sum + s.scoreOnItem, 0);
    const B_B = lowerGroup.reduce((sum, s) => sum + s.scoreOnItem, 0);

    const P_A = B_A / J_A;
    const P_B = B_B / J_B;
    const D = P_A - P_B;

    let discCategory = "";
    if (N <= 30) {
      if (D < 0.00) discCategory = "Sangat Jelek (Ditolak)";
      else if (D < 0.20) discCategory = "Jelek";
      else if (D < 0.40) discCategory = "Cukup";
      else if (D < 0.70) discCategory = "Baik";
      else discCategory = "Baik Sekali";
    } else {
      // Arikunto (2013): D < 0.20 adalah Jelek (termasuk nilai negatif jika ada)
      if (D < 0.20) discCategory = "Jelek";
      else if (D < 0.40) discCategory = "Cukup";
      else if (D < 0.70) discCategory = "Baik";
      else discCategory = "Baik Sekali";
    }

    // Simpan hasil butir soal
    analysisResults.items.push({
      index: q,
      sumX,
      sumXSq,
      sumXY,
      r_xy: parseFloat(r_xy.toFixed(4)),
      r_tabel,
      isValid,
      P: parseFloat(P.toFixed(3)),
      diffCategory,
      D: parseFloat(D.toFixed(3)),
      discCategory,
      B_A, J_A, B_B, J_B, P_A, P_B,
      upperGroupIds: upperGroup.map(s => s.id),
      lowerGroupIds: lowerGroup.map(s => s.id)
    });
  }

  // 5. Reliabilitas KR-20 (r_11)
  let r11 = 0;
  if (k > 1 && varianceTotal > 0) {
    r11 = (k / (k - 1)) * (1 - (sumPq / varianceTotal));
  } else {
    r11 = 0;
  }

  let relCategory = "";
  if (r11 >= 0.80) relCategory = "Sangat Tinggi";
  else if (r11 >= 0.60) relCategory = "Tinggi";
  else if (r11 >= 0.40) relCategory = "Sedang";
  else if (r11 >= 0.20) relCategory = "Rendah";
  else relCategory = "Sangat Rendah";

  const isReliable = r11 >= 0.60;

  analysisResults.reliability = {
    k,
    sumPq: parseFloat(sumPq.toFixed(4)),
    varianceTotal: parseFloat(varianceTotal.toFixed(4)),
    r11: parseFloat(r11.toFixed(4)),
    category: relCategory,
    status: isReliable ? "Reliabel" : "Tidak Reliabel"
  };

  // Render semua visual & hasil ke UI
  updateDashboardWidgets();
  renderAnalysisResultsTable();
  renderCalculationExplanation();
  renderReliabilityExplanation();
  renderCharts();
}

/**
 * Update Dashboard Metrik Widgets
 */
function updateDashboardWidgets() {
  const rel = analysisResults.reliability;
  
  document.getElementById('widget-respondents').textContent = state.students.length;
  document.getElementById('widget-questions').textContent = state.questionsCount;
  
  const krVal = document.getElementById('widget-kr-val');
  const krSub = document.getElementById('widget-kr-sub');
  const krCard = document.getElementById('widget-kr-card');

  krVal.textContent = rel.r11.toFixed(3);
  krSub.innerHTML = `<i class="lucide-activity"></i> Kategori: ${rel.category} (${rel.status})`;

  // Set warna kartu reliabilitas
  krCard.className = 'widget-card';
  if (rel.r11 >= 0.60) {
    krCard.classList.add('success');
  } else if (rel.r11 >= 0.40) {
    krCard.classList.add('warning');
  } else {
    krCard.classList.add('error');
  }

  // Ringkasan Butir Soal Valid vs Gugur
  const validCount = analysisResults.items.filter(item => item.isValid).length;
  const invalidCount = state.questionsCount - validCount;

  const validWidget = document.getElementById('widget-valid-items');
  validWidget.textContent = `${validCount} Soal`;
  document.getElementById('widget-invalid-sub').textContent = `${invalidCount} Soal Tidak Valid / Gugur`;
}

/**
 * Render Tabel Hasil Analisis Utama
 */
function renderAnalysisResultsTable() {
  const tbody = document.getElementById('results-tbody');
  tbody.innerHTML = '';

  analysisResults.items.forEach(item => {
    const tr = document.createElement('tr');

    // Soal Ke-
    const tdIndex = document.createElement('td');
    tdIndex.innerHTML = `Soal ${item.index + 1} <br><small style="color: var(--text-muted);">Kunci: ${state.answerKey[item.index]}</small>`;
    tdIndex.style.fontWeight = '600';
    tr.appendChild(tdIndex);

    // Jawaban Benar
    const tdBenar = document.createElement('td');
    tdBenar.textContent = item.sumX;
    tr.appendChild(tdBenar);

    // Indeks Kesukaran P
    const tdP = document.createElement('td');
    tdP.innerHTML = `${item.P} <br><span class="badge ${getDifficultyBadgeClass(item.diffCategory)}">${item.diffCategory}</span>`;
    tr.appendChild(tdP);

    // Daya Pembeda D
    const tdD = document.createElement('td');
    tdD.innerHTML = `${item.D} <br><span class="badge ${getDiscriminationBadgeClass(item.D, state.students.length)}">${item.discCategory}</span>`;
    tr.appendChild(tdD);

    // Pearson r_xy
    const tdrxy = document.createElement('td');
    tdrxy.textContent = item.r_xy;
    tr.appendChild(tdrxy);

    // r_tabel
    const tdrtabel = document.createElement('td');
    tdrtabel.textContent = item.r_tabel;
    tr.appendChild(tdrtabel);

    // Status Validitas
    const tdStatus = document.createElement('td');
    tdStatus.innerHTML = item.isValid 
      ? `<span class="badge badge-success"><i class="lucide-check"></i> Valid</span><br><small style="color: var(--text-muted); font-size: 0.75rem; font-weight: 500;">r<sub>xy</sub> (${item.r_xy}) &gt; r<sub>tabel</sub> (${item.r_tabel})</small>`
      : `<span class="badge badge-error"><i class="lucide-x"></i> Gugur</span><br><small style="color: var(--text-muted); font-size: 0.75rem; font-weight: 500;">r<sub>xy</sub> (${item.r_xy}) &le; r<sub>tabel</sub> (${item.r_tabel})</small>`;
    tr.appendChild(tdStatus);

    tbody.appendChild(tr);
  });
}

function getDifficultyBadgeClass(cat) {
  if (cat === 'Mudah') return 'badge-success';
  if (cat === 'Sedang') return 'badge-info';
  return 'badge-warning';
}

function getDiscriminationBadgeClass(dVal, N) {
  if (N <= 30) {
    if (dVal < 0) return 'badge-error';
    if (dVal < 0.20) return 'badge-muted';
    if (dVal < 0.40) return 'badge-warning';
    if (dVal < 0.70) return 'badge-info';
    return 'badge-success';
  } else {
    if (dVal < 0.20) return 'badge-muted';
    if (dVal < 0.40) return 'badge-warning';
    if (dVal < 0.70) return 'badge-info';
    return 'badge-success';
  }
}

/**
 * Render Panel Detail Perhitungan Langkah Demi Langkah
 */
function renderCalculationExplanation() {
  const itemSelector = document.getElementById('calc-item-selector');
  itemSelector.innerHTML = '';

  // 1. Bangun daftar pemilih soal di sebelah kiri
  analysisResults.items.forEach((item, idx) => {
    const btn = document.createElement('button');
    btn.className = `item-select-btn ${idx === state.selectedQuestionIndex ? 'active' : ''}`;
    btn.innerHTML = `<i class="lucide-help-circle"></i> Butir Soal ${idx + 1} (${item.isValid ? 'Valid' : 'Gugur'} | r = ${item.r_xy})`;
    btn.addEventListener('click', () => {
      state.selectedQuestionIndex = idx;
      renderCalculationExplanation();
    });
    itemSelector.appendChild(btn);
  });

  // 2. Tampilkan perhitungan detail soal yang aktif
  const activeItem = analysisResults.items[state.selectedQuestionIndex];
  if (!activeItem) return;

  const N = state.students.length;
  const itemNo = activeItem.index + 1;
  const correctKey = state.answerKey[activeItem.index] || 'A';

  // Update Judul Detail
  document.getElementById('calc-title').innerHTML = `Detail Analisis: Butir Soal ${itemNo} (Kunci Jawaban: <strong style="color: var(--primary);">${correctKey}</strong>)`;

  // render tabel bantu perkalian Pearson
  const helperTableBody = document.getElementById('calc-helper-tbody');
  helperTableBody.innerHTML = '';
  
  let totalX = 0;
  let totalY = 0;
  let totalXSq = 0;
  let totalYSq = 0;
  let totalXY = 0;

  state.students.forEach(s => {
    // Cari jawaban asli siswa (A-E)
    const rawAnswer = s.scores[activeItem.index] || '';
    // Evaluasi biner
    const X = (rawAnswer && rawAnswer.toUpperCase() === correctKey.toUpperCase()) ? 1 : 0;
    const Y = calculateStudentEvaluatedTotal(s);
    const XSq = Math.pow(X, 2);
    const YSq = Math.pow(Y, 2);
    const XY = X * Y;

    totalX += X;
    totalY += Y;
    totalXSq += XSq;
    totalYSq += YSq;
    totalXY += XY;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${s.name}</td>
      <td style="font-weight: 600;">${rawAnswer || '-'} <span style="font-size:0.75rem; color: var(--text-muted);">(${X === 1 ? 'Benar: 1' : 'Salah: 0'})</span></td>
      <td>${Y}</td>
      <td>${XSq}</td>
      <td>${YSq}</td>
      <td style="font-weight: 600;">${XY}</td>
    `;
    helperTableBody.appendChild(tr);
  });

  // Render Baris Jumlah
  const trSum = document.createElement('tr');
  trSum.style.fontWeight = 'bold';
  trSum.style.backgroundColor = 'var(--bg-app)';
  trSum.innerHTML = `
    <td>Jumlah (Σ)</td>
    <td>ΣX = ${totalX} (Benar)</td>
    <td>ΣY = ${totalY}</td>
    <td>ΣX² = ${totalXSq}</td>
    <td>ΣY² = ${totalYSq}</td>
    <td>ΣXY = ${totalXY}</td>
  `;
  helperTableBody.appendChild(trSum);

  // Render Penjelasan Rumus Validitas
  const numerator = (N * totalXY) - (totalX * totalY);
  const sumTotalScoresSq = state.students.reduce((sum, s) => sum + Math.pow(calculateStudentEvaluatedTotal(s), 2), 0);
  const sumTotalScores = totalY;
  const denomX = (N * totalXSq) - Math.pow(totalX, 2);
  const denomY = (N * sumTotalScoresSq) - Math.pow(sumTotalScores, 2);
  const denominator = Math.sqrt(denomX * denomY);

  document.getElementById('val-eq-step-1').innerHTML = `
    r<sub>xy</sub> = [ ${N} &times; ${totalXY} - (${totalX} &times; ${totalY}) ] / √[ (${N} &times; ${totalXSq} - ${totalX}<sup>2</sup>) &times; (${N} &times; ${totalYSq} - ${totalY}<sup>2</sup>) ]
  `;
  document.getElementById('val-eq-step-2').innerHTML = `
    r<sub>xy</sub> = [ ${N * totalXY} - ${totalX * totalY} ] / √[ (${N * totalXSq - Math.pow(totalX, 2)}) &times; (${N * totalYSq - Math.pow(totalY, 2)}) ]
  `;
  document.getElementById('val-eq-step-3').innerHTML = `
    r<sub>xy</sub> = ${numerator} / √[ ${denomX} &times; ${denomY} ] = ${numerator} / √[ ${denomX * denomY} ]
  `;
  document.getElementById('val-eq-step-4').innerHTML = `
    r<sub>xy</sub> = ${numerator} / ${denominator.toFixed(4)} = <strong>${activeItem.r_xy}</strong>
  `;

  // Status Validitas
  const valStatusText = document.getElementById('val-status-explanation');
  if (activeItem.isValid) {
    valStatusText.className = 'alert-box';
    valStatusText.innerHTML = `
      <strong>Analisis Validitas:</strong> Karena r<sub>xy</sub> (${activeItem.r_xy}) &gt; r<sub>tabel</sub> (${activeItem.r_tabel}) pada taraf signifikansi 5% dengan df = ${N-2}, maka butir soal ke-${itemNo} dinyatakan <strong>VALID</strong>.
    `;
  } else {
    valStatusText.className = 'alert-box warning';
    valStatusText.innerHTML = `
      <strong>Analisis Validitas:</strong> Karena r<sub>xy</sub> (${activeItem.r_xy}) &le; r<sub>tabel</sub> (${activeItem.r_tabel}) pada taraf signifikansi 5% dengan df = ${N-2}, maka butir soal ke-${itemNo} dinyatakan <strong>TIDAK VALID (GUGUR)</strong>.
    `;
  }

  // Render Penjelasan Kesukaran
  document.getElementById('diff-eq-step').innerHTML = `
    P = B / JS = ${activeItem.sumX} / ${N} = <strong>${activeItem.P}</strong>
  `;
  document.getElementById('diff-explanation').innerHTML = `
    Indeks kesukaran butir soal adalah <strong>${activeItem.P}</strong>. Berdasarkan kriteria Arikunto, nilai ini termasuk dalam kategori <strong>${activeItem.diffCategory}</strong>.
  `;

  // Render Penjelasan Daya Pembeda
  const N_type = N <= 30 ? "Sampel Kecil (N ≤ 30)" : "Sampel Besar (N > 30)";
  const groupPct = N <= 30 ? "50% atas & 50% bawah" : "27% atas & 27% bawah";
  
  document.getElementById('disc-group-type').textContent = `${N_type} menggunakan pembagian kelompok ${groupPct}`;
  
  document.getElementById('disc-eq-step-1').innerHTML = `
    P<sub>A</sub> = B<sub>A</sub> / J<sub>A</sub> = ${activeItem.B_A} / ${activeItem.J_A} = ${activeItem.P_A.toFixed(3)}<br>
    P<sub>B</sub> = B<sub>B</sub> / J<sub>B</sub> = ${activeItem.B_B} / ${activeItem.J_B} = ${activeItem.P_B.toFixed(3)}
  `;
  document.getElementById('disc-eq-step-2').innerHTML = `
    D = P<sub>A</sub> - P<sub>B</sub> = ${activeItem.P_A.toFixed(3)} - ${activeItem.P_B.toFixed(3)} = <strong>${activeItem.D}</strong>
  `;
  document.getElementById('disc-explanation').innerHTML = `
    Indeks daya pembeda diperoleh sebesar <strong>${activeItem.D}</strong> yang berkategori <strong>${activeItem.discCategory}</strong>. Hal ini menunjukkan soal ini memiliki kemampuan yang <strong>${getDInterpretation(activeItem.D, N)}</strong> untuk membedakan antara siswa pandai dan siswa kurang pandai.
  `;
}

function getDInterpretation(d, N) {
  if (N <= 30) {
    if (d < 0) return "sangat buruk (harus dibuang karena terbalik)";
    if (d < 0.20) return "kurang/jelek (perlu direvisi total)";
    if (d < 0.40) return "cukup (bisa digunakan setelah perbaikan)";
    if (d < 0.70) return "baik (layak digunakan)";
    return "sangat baik (sangat layak digunakan)";
  } else {
    // Arikunto (2013)
    if (d < 0.20) return "jelek (perlu direvisi)";
    if (d < 0.40) return "cukup (bisa digunakan setelah perbaikan)";
    if (d < 0.70) return "baik (layak digunakan)";
    return "baik sekali (sangat layak digunakan)";
  }
}

/**
 * Render Tab Detail Reliabilitas KR-20
 */
function renderReliabilityExplanation() {
  const rel = analysisResults.reliability;
  const N = state.students.length;
  
  // Render tabel p, q, pq per butir soal
  const relTableBody = document.getElementById('rel-table-tbody');
  relTableBody.innerHTML = '';

  analysisResults.items.forEach(item => {
    const q_val = parseFloat((1 - item.P).toFixed(3));
    const pq = parseFloat((item.P * q_val).toFixed(4));

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>Soal ${item.index + 1} (Kunci: ${state.answerKey[item.index]})</td>
      <td>${item.P}</td>
      <td>${q_val}</td>
      <td style="font-weight: 600;">${pq}</td>
    `;
    relTableBody.appendChild(tr);
  });

  // Render baris total pq
  const trTotal = document.createElement('tr');
  trTotal.style.fontWeight = 'bold';
  trTotal.style.backgroundColor = 'var(--bg-app)';
  trTotal.innerHTML = `
    <td>Total Jumlah</td>
    <td>-</td>
    <td>-</td>
    <td>Σpq = ${rel.sumPq}</td>
  `;
  relTableBody.appendChild(trTotal);

  // Penjelasan Rumus Varians
  const totalScores = state.students.map(s => calculateStudentEvaluatedTotal(s));
  const sumY = totalScores.reduce((a, b) => a + b, 0);
  const sumYSq = totalScores.reduce((sum, score) => sum + Math.pow(score, 2), 0);

  document.getElementById('rel-var-step-1').innerHTML = `
    S<sub>t</sub><sup>2</sup> = [ ${sumYSq} - (${sumY}<sup>2</sup> / ${N}) ] / ${N}
  `;
  document.getElementById('rel-var-step-2').innerHTML = `
    S<sub>t</sub><sup>2</sup> = [ ${sumYSq} - (${Math.pow(sumY, 2)} / ${N}) ] / ${N} = [ ${sumYSq} - ${(Math.pow(sumY, 2) / N).toFixed(2)} ] / ${N}
  `;
  document.getElementById('rel-var-step-3').innerHTML = `
    S<sub>t</sub><sup>2</sup> = ${(sumYSq - (Math.pow(sumY, 2) / N)).toFixed(2)} / ${N} = <strong>${rel.varianceTotal}</strong>
  `;

  // Penjelasan Rumus KR-20
  const k = rel.k;
  document.getElementById('rel-kr-step-1').innerHTML = `
    r<sub>11</sub> = [ ${k} / (${k} - 1) ] &times; [ 1 - (${rel.sumPq} / ${rel.varianceTotal}) ]
  `;
  document.getElementById('rel-kr-step-2').innerHTML = `
    r<sub>11</sub> = [ ${(k / (k-1)).toFixed(4)} ] &times; [ 1 - ${(rel.sumPq / rel.varianceTotal).toFixed(4)} ]
  `;
  document.getElementById('rel-kr-step-3').innerHTML = `
    r<sub>11</sub> = ${(k / (k-1)).toFixed(4)} &times; ${(1 - (rel.sumPq / rel.varianceTotal)).toFixed(4)} = <strong>${rel.r11}</strong>
  `;

  // Status kesimpulan reliabilitas
  const statusBox = document.getElementById('rel-status-explanation');
  if (rel.r11 >= 0.60) {
    statusBox.className = 'alert-box';
    statusBox.innerHTML = `
      <strong>Kesimpulan Reliabilitas:</strong> Karena r<sub>11</sub> (${rel.r11}) &ge; 0.60, maka instrumen tes ini dinyatakan <strong>RELIABEL</strong> dengan kategori tingkat reliabilitas <strong>${rel.category}</strong>. Kategori ini sangat ideal untuk instrumen penilaian kelas.
    `;
  } else {
    statusBox.className = 'alert-box warning';
    statusBox.innerHTML = `
      <strong>Kesimpulan Reliabilitas:</strong> Karena r<sub>11</sub> (${rel.r11}) &lt; 0.60, maka instrumen tes ini dinyatakan <strong>BELUM RELIABEL</strong> dengan kategori <strong>${rel.category}</strong>. Disarankan untuk membuang atau memperbaiki butir-butir soal yang tidak valid dan menjalankan analisis ulang.
    `;
  }
}

/**
 * ==========================================
 * VISUALISASI GRAFIK DENGAN CHART.JS
 * ==========================================
 */
let difficultyChartInstance = null;
let validityChartInstance = null;

function renderCharts() {
  const ctxDiff = document.getElementById('chart-difficulty').getContext('2d');
  const ctxVal = document.getElementById('chart-validity').getContext('2d');

  if (difficultyChartInstance) difficultyChartInstance.destroy();
  if (validityChartInstance) validityChartInstance.destroy();

  let diffCounts = { 'Mudah': 0, 'Sedang': 0, 'Sukar': 0 };
  analysisResults.items.forEach(item => {
    diffCounts[item.diffCategory]++;
  });

  let validCounts = { 'Valid': 0, 'Gugur (Tidak Valid)': 0 };
  analysisResults.items.forEach(item => {
    if (item.isValid) validCounts['Valid']++;
    else validCounts['Gugur (Tidak Valid)']++;
  });

  difficultyChartInstance = new Chart(ctxDiff, {
    type: 'bar',
    data: {
      labels: ['Mudah', 'Sedang', 'Sukar'],
      datasets: [{
        label: 'Jumlah Butir Soal',
        data: [diffCounts['Mudah'], diffCounts['Sedang'], diffCounts['Sukar']],
        backgroundColor: [
          'rgba(16, 185, 129, 0.6)',
          'rgba(6, 182, 212, 0.6)',
          'rgba(245, 158, 11, 0.6)'
        ],
        borderColor: [
          'rgb(16, 185, 129)',
          'rgb(6, 182, 212)',
          'rgb(245, 158, 11)'
        ],
        borderWidth: 1.5,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { stepSize: 1 }
        }
      }
    }
  });

  validityChartInstance = new Chart(ctxVal, {
    type: 'doughnut',
    data: {
      labels: ['Valid', 'Gugur (Tidak Valid)'],
      datasets: [{
        data: [validCounts['Valid'], validCounts['Gugur (Tidak Valid)']],
        backgroundColor: [
          'rgba(16, 185, 129, 0.6)',
          'rgba(244, 63, 94, 0.6)'
        ],
        borderColor: [
          'rgb(16, 185, 129)',
          'rgb(244, 63, 94)'
        ],
        borderWidth: 1.5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            padding: 15,
            boxWidth: 12
          }
        }
      },
      cutout: '65%'
    }
  });
}

/**
 * ==========================================
 * PENANGANAN IMPOR & EKSPOR DATA
 * ==========================================
 */

/**
 * Membaca dan memproses file excel/csv hasil upload
 */
function handleImportFile(file) {
  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (rawRows.length < 2) {
        alert("File Excel kosong atau tidak memiliki format yang sesuai.");
        return;
      }

      const headerRow = rawRows[0];
      let nameColIdx = 1;
      let startQColIdx = 2;

      // Cari kolom Nama
      const nameColName = headerRow.find(cell => typeof cell === 'string' && cell.toLowerCase().includes('nama'));
      if (nameColName) {
        nameColIdx = headerRow.indexOf(nameColName);
        startQColIdx = nameColIdx + 1;
      }

      // Cari kolom Total Skor jika ada
      let endQColIdx = headerRow.length;
      const totalColIdx = headerRow.findIndex(cell => typeof cell === 'string' && (cell.toLowerCase().includes('total') || cell.toLowerCase().includes('skor')));
      if (totalColIdx !== -1 && totalColIdx > startQColIdx) {
        endQColIdx = totalColIdx;
      }

      const numQuestions = endQColIdx - startQColIdx;
      if (numQuestions <= 0) {
        alert("Gagal mendeteksi kolom soal. Pastikan kolom soal dinamai seperti 'S1', 'S2', dst.");
        return;
      }

      let importedKey = [];
      const importedStudents = [];

      // Membaca baris-baris data
      for (let r = 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;

        const respondentName = row[nameColIdx] ? String(row[nameColIdx]).trim() : '';
        if (!respondentName) continue;

        // Cek apakah baris ini berisi KUNCI JAWABAN
        if (respondentName.toUpperCase().includes("KUNCI") || respondentName.toUpperCase() === "KEY") {
          // Ekstrak Kunci
          importedKey = [];
          for (let q = startQColIdx; q < endQColIdx; q++) {
            const val = row[q] ? String(row[q]).trim().toUpperCase() : 'A';
            importedKey.push(val);
          }
        } else {
          // Baris siswa biasa
          const scores = [];
          for (let q = startQColIdx; q < endQColIdx; q++) {
            const val = row[q] !== undefined ? String(row[q]).trim().toUpperCase() : '';
            scores.push(val);
          }

          importedStudents.push({
            id: importedStudents.length + 1,
            name: respondentName,
            scores: scores
          });
        }
      }

      if (importedStudents.length === 0) {
        alert("Tidak ada data responden yang berhasil diimpor.");
        return;
      }

      // Atur state
      state.students = importedStudents;
      state.studentsCount = importedStudents.length;
      state.questionsCount = numQuestions;
      
      // Jika kunci jawaban diimpor dari file, gunakan. Jika tidak ada, buat kunci default 'A'
      if (importedKey.length === numQuestions) {
        state.answerKey = importedKey;
      } else {
        state.answerKey = Array(numQuestions).fill('A');
        alert("Kunci jawaban tidak ditemukan di file Excel. Sistem otomatis menetapkan 'A' untuk semua soal. Silakan ubah kunci jawaban di tabel jika diperlukan.");
      }

      renderGrid();
      updateGridInputs();

      // Tutup modal
      document.getElementById('import-modal').classList.remove('active');
      alert(`Berhasil mengimpor ${state.studentsCount} siswa dan ${state.questionsCount} butir soal!`);

    } catch (err) {
      console.error(err);
      alert("Terjadi kesalahan saat mengurai file Excel. Pastikan format file sesuai.");
    }
  };

  reader.readAsArrayBuffer(file);
}

/**
 * Mengekspor data spreadsheet input ke file XLSX Excel
 */
function exportToExcel() {
  const wb = XLSX.utils.book_new();
  const wsData = [];

  // Buat Header
  const header = ["No", "Nama Responden"];
  for (let q = 0; q < state.questionsCount; q++) {
    header.push(`S${q + 1}`);
  }
  header.push("Skor Total");
  wsData.push(header);

  // Buat Baris KUNCI JAWABAN (agar saat diimpor kembali, aplikasi tahu kuncinya)
  const keyRow = ["🔑", "KUNCI JAWABAN"];
  state.answerKey.forEach(key => keyRow.push(key));
  keyRow.push("-");
  wsData.push(keyRow);

  // Buat Baris Data Siswa
  state.students.forEach((s, idx) => {
    const row = [idx + 1, s.name];
    s.scores.forEach(score => row.push(score || ''));
    row.push(calculateStudentEvaluatedTotal(s));
    wsData.push(row);
  });

  // Buat Baris Jumlah Benar di paling bawah
  const footer = ["", "Jumlah Benar"];
  for (let q = 0; q < state.questionsCount; q++) {
    footer.push(calculateSumCorrect(q));
  }
  footer.push(calculateGrandTotal());
  wsData.push(footer);

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "Input Analisis Pilihan Ganda");

  // Jika analisis sudah dihitung, sisipkan sheet ringkasan hasil
  if (analysisResults.items.length > 0) {
    const wsResultsData = [[
      "Butir Soal", "Kunci Jawaban", "Jumlah Jawaban Benar", "Indeks Kesukaran (P)", "Kategori Kesukaran", 
      "Daya Pembeda (D)", "Kategori Daya Pembeda", "Korelasi Pearson (r_xy)", "r_tabel", "Status Validitas"
    ]];

    analysisResults.items.forEach(item => {
      wsResultsData.push([
        `Soal ${item.index + 1}`,
        state.answerKey[item.index],
        item.sumX,
        item.P,
        item.diffCategory,
        item.D,
        item.discCategory,
        item.r_xy,
        item.r_tabel,
        item.isValid ? `Valid (${item.r_xy} > ${item.r_tabel})` : `Gugur (${item.r_xy} <= ${item.r_tabel})`
      ]);
    });

    const wsResults = XLSX.utils.aoa_to_sheet(wsResultsData);
    XLSX.utils.book_append_sheet(wb, wsResults, "Hasil Analisis Butir");
  }

  XLSX.writeFile(wb, `analisis_butir_pilihan_ganda_${Date.now()}.xlsx`);
}

/**
 * Ekspor Hasil Analisis Ke PDF
 */
function exportToPDF() {
  switchTab('hasil-analisis');
  setTimeout(() => {
    window.print();
  }, 300);
}
