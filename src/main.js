import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

// Hardcoded Team Members
const MEMBERS = {
  member1: "김한주",
  member2: "김용석",
  member3: "문경수",
  member4: "홍수민"
};

// Mock Metadata based on user image
const SONG_METADATA = [
  { title: "Syren", artist: "Anyma", cover: "🎵" },
  { title: "Playing with Fire", artist: "BLACKPINK", cover: "🔥" },
  { title: "Call Me Maybe", artist: "Carly Rae Jepsen", cover: "📱" },
  { title: "Sie ergibt sich nicht", artist: "Chang Eun Ah", cover: "🎭" },
  { title: "Lemon Tree", artist: "Fools Garden", cover: "🍋" },
  { title: "Spain", artist: "Jesus Molina", cover: "🇪🇸" },
  { title: "Jane Doe", artist: "Kenshi Yonezu", cover: "👤" },
  { title: "Peligrosa", artist: "Ojos", cover: "⚠️" },
  { title: "Lost Chapter", artist: "Pentakill, Jorn", cover: "📖" },
  { title: "Attack on Titan", artist: "Sawano Hiroyuki", cover: "⚔️" },
  { title: "6 Moments musicaux, Op. 16 : No. 4 in E", artist: "Sergei Rachmaninoff", cover: "🎹" },
  { title: "Shoreditch", artist: "Vard", cover: "🎸" },
  { title: "Look at Me!", artist: "XXTENTACION", cover: "💥" }
];

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];
const LABELS_KO = ['몰입도(Eng)', '흥미도(Int)', '활성도(Exc)', '스트레스(Str)', '이완도(Rel)'];
const SHORT_LABELS = ['En', 'In', 'Ex', 'St', 'Re'];

const app = document.querySelector('#app');

app.innerHTML = `
  <div id="setupScreen" class="screen active">
    <header>
      <h1>NeuroWav Dashboard</h1>
      <p class="subtitle">Upload brainwave data (Excel/CSV) for your team member</p>
    </header>
    
    <div class="user-selector">
      <label>Select Team Member:</label>
      <select id="userSelect">
        <option value="member1">조원 1: ${MEMBERS.member1}</option>
        <option value="member2">조원 2: ${MEMBERS.member2}</option>
        <option value="member3">조원 3: ${MEMBERS.member3}</option>
        <option value="member4">조원 4: ${MEMBERS.member4}</option>
      </select>
    </div>

    <div class="upload-section" id="uploadSection">
      <div class="upload-box" id="dropZone">
        <div class="upload-icon">📁</div>
        <div class="upload-text">Drag & Drop all 13 song data files</div>
        <div class="upload-subtext">Will be automatically sorted 1 to 13</div>
        <input type="file" id="fileInput" multiple accept=".csv, .xlsx" style="display: none;">
      </div>
      <div class="error-message" id="errorMessage"></div>
    </div>
  </div>

  <div id="dashboardScreen" class="screen">
    <header class="dash-header">
      <div class="user-info">Viewing: <span><span id="currentUserDisplay"></span>'s Dashboard</span></div>
    </header>

    <div class="dashboard-layout">
      <!-- Left Panel: Song Info -->
      <div class="panel song-info-panel">
        <div class="album-art" id="albumArt">💿</div>
        <h2 id="songTitle">Select a Song</h2>
        <h3 id="songArtist">-</h3>
        <p id="songDesc" class="song-desc">Click on a radar chart to view detailed NeuroWav trends.</p>
      </div>

      <!-- Right Panel: Radar Grid -->
      <div class="panel radar-grid-panel">
        <div class="radar-grid" id="radarGrid"></div>
      </div>

      <!-- Bottom Panel: Detailed Line Chart -->
      <div class="panel line-chart-panel">
        <div class="chart-header">
          <h3 id="lineChartTitle">Raw NeuroWav Trends (Over Time)</h3>
        </div>
        <div style="position: relative; height: calc(100% - 2rem); width: 100%;">
          <canvas id="lineChart"></canvas>
        </div>
      </div>

      <!-- Bottom Right Panel: Personal Rankings -->
      <div class="panel ranking-panel">
        <div class="ranking-header">🏆 My Top 3 Songs</div>
        <div id="rankingContainer"></div>
      </div>
    </div>
  </div>

  <div id="teamScreen" class="screen">
    <header>
      <h1>Team Comparison: NeuroWav Awards</h1>
      <p class="subtitle">Peak and Total Mass Analysis across all team members</p>
    </header>
    <div class="team-awards-container" id="teamGrid" style="display: flex; flex-direction: column; gap: 2rem;">
      <!-- Award sections injected here -->
    </div>
  </div>

  <div class="loading-overlay" id="loadingOverlay">
    <div class="spinner"></div>
    <div class="loading-text">Processing data...</div>
  </div>
`;

// Navigation Setup
const navUpload = document.getElementById('navUpload');
const navDashboard = document.getElementById('navDashboard');
const navTeam = document.getElementById('navTeam');
const setupScreen = document.getElementById('setupScreen');
const dashboardScreen = document.getElementById('dashboardScreen');
const teamScreen = document.getElementById('teamScreen');

function switchScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-links button').forEach(b => b.classList.remove('active'));
  
  if (screenId === 'setupScreen') { navUpload.classList.add('active'); setupScreen.classList.add('active'); }
  if (screenId === 'dashboardScreen') { navDashboard.classList.add('active'); loadDashboard(); }
  if (screenId === 'teamScreen') { navTeam.classList.add('active'); loadTeamComparison(); }
}

navUpload.addEventListener('click', () => switchScreen('setupScreen'));
navDashboard.addEventListener('click', () => switchScreen('dashboardScreen'));
navTeam.addEventListener('click', () => switchScreen('teamScreen'));

// Global State
let currentUserId = 'member1';
let userData = {}; 
let lineChartInstance = null;

// DOM Elements
const userSelect = document.getElementById('userSelect');
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const loadingOverlay = document.getElementById('loadingOverlay');
const errorMessage = document.getElementById('errorMessage');

const currentUserDisplay = document.getElementById('currentUserDisplay');
const albumArt = document.getElementById('albumArt');
const songTitle = document.getElementById('songTitle');
const songArtist = document.getElementById('songArtist');
const radarGrid = document.getElementById('radarGrid');
const lineChartTitle = document.getElementById('lineChartTitle');
const rankingContainer = document.getElementById('rankingContainer');
const teamGrid = document.getElementById('teamGrid');

userSelect.addEventListener('change', (e) => {
  currentUserId = e.target.value;
});

// File Upload Logic
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
});
['dragenter', 'dragover'].forEach(eventName => {
  dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
});
['dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
});

dropZone.addEventListener('drop', (e) => {
  const files = e.dataTransfer.files;
  if (files.length > 0) handleFiles(Array.from(files));
});
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
  const files = e.target.files;
  if (files.length > 0) handleFiles(Array.from(files));
});

function showError(msg) {
  errorMessage.textContent = msg;
  errorMessage.style.display = 'block';
  loadingOverlay.style.display = 'none';
}

function extractNumber(filename) {
  const match = filename.match(/\d+/);
  return match ? parseInt(match[0], 10) : 9999;
}

async function handleFiles(fileArray) {
  if (!fileArray || fileArray.length === 0) return;
  errorMessage.style.display = 'none';
  loadingOverlay.style.display = 'flex';
  
  fileArray.sort((a, b) => extractNumber(a.name) - extractNumber(b.name));
  const results = [];
  
  try {
    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      const data = await readFile(file);
      const parsedData = processExcelData(data);
      const stats = calculateStats(parsedData, file.name, i + 1);
      if (stats) results.push(stats);
    }
    
    if (results.length === 0) {
      showError('No valid data found in the uploaded files.');
      return;
    }
    
    localStorage.setItem(`brainwaveData_${currentUserId}`, JSON.stringify(results));
    switchScreen('dashboardScreen');
    
  } catch (err) {
    console.error(err);
    showError('Error processing files.');
  } finally {
    loadingOverlay.style.display = 'none';
  }
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
}

function processExcelData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  let allRows = [];
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    if (rows && rows.length > 0) {
      allRows = allRows.concat(rows);
    }
  });
  return allRows;
}

function normalizeKey(key) {
  return typeof key === 'string' ? key.toLowerCase().trim() : '';
}

function calculateStats(rows, fileName, songNumber) {
  if (!rows || rows.length === 0) return null;
  
  let headerRowIdx = -1;
  let columnIndices = {};
  
  const MATCHERS = {
    'engagement': ['engagement', 'eng', '몰입'],
    'interest': ['interest', 'int', '흥미'],
    'excitement': ['excitement', 'exc', '활성'],
    'stress': ['stress', 'str', '스트레스'],
    'relaxation': ['relaxation', 'rel', 'relation', '이완']
  };
  
  // Search for the header row in ALL rows (in case data starts way down)
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    let foundCount = 0;
    const tempMap = {};
    
    row.forEach((cell, idx) => {
      const cellStr = normalizeKey(String(cell));
      TARGET_COLUMNS.forEach(target => {
        const matchers = MATCHERS[target];
        if (matchers.some(m => cellStr.includes(m))) {
           tempMap[target] = idx;
           foundCount++;
        }
      });
    });
    
    if (foundCount >= 1) {
      headerRowIdx = i;
      columnIndices = tempMap;
      if (foundCount >= 3) break; 
    }
  }
  
  if (headerRowIdx === -1) return null;
  
  const sums = {};
  const counts = {};
  const rawData = {};
  
  TARGET_COLUMNS.forEach(col => { sums[col] = 0; counts[col] = 0; rawData[col] = []; });
  
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    TARGET_COLUMNS.forEach(col => {
      const colIdx = columnIndices[col];
      if (colIdx !== undefined && row[colIdx] !== null && row[colIdx] !== undefined) {
        let valStr = String(row[colIdx]).replace(/,/g, '');
        const val = parseFloat(valStr);
        if (!isNaN(val)) {
          sums[col] += val;
          counts[col]++;
          rawData[col].push(val);
        }
      }
    });
  }
  
  const totalValidData = Object.values(counts).reduce((acc, c) => acc + c, 0);
  if (totalValidData === 0) return null;
  
  const stats = { fileName, songNumber, rawData, averages: {} };
  TARGET_COLUMNS.forEach(col => {
    stats.averages[col] = counts[col] > 0 ? (sums[col] / counts[col]) : 0;
  });
  
  return stats;
}

// Personal Dashboard Logic
function loadDashboard() {
  const savedData = localStorage.getItem(`brainwaveData_${currentUserId}`);
  if (!savedData) {
    alert('No data uploaded yet for ' + MEMBERS[currentUserId] + '. Please upload files first.');
    switchScreen('setupScreen');
    return;
  }
  
  dashboardScreen.classList.add('active');
  userData = JSON.parse(savedData);
  currentUserDisplay.textContent = MEMBERS[currentUserId];
  
  renderRadarGrid();
  renderRankings();
  
  if (userData.length > 0) selectSong(0);
}

function renderRadarGrid() {
  radarGrid.innerHTML = '';
  userData.forEach((song, index) => {
    const metaIndex = Math.min(index, SONG_METADATA.length - 1);
    const meta = SONG_METADATA[metaIndex];
    
    const div = document.createElement('div');
    div.className = 'radar-item';
    div.id = `radar-item-${index}`;
    div.innerHTML = `
      <div class="radar-wrapper"><canvas id="radar-${currentUserId}-${index}"></canvas></div>
      <div class="radar-title">${meta.title}</div>
    `;
    div.addEventListener('click', () => selectSong(index));
    radarGrid.appendChild(div);
    
    const ctx = document.getElementById(`radar-${currentUserId}-${index}`).getContext('2d');
    const dataVals = TARGET_COLUMNS.map(col => song.averages[col]);
    
    new Chart(ctx, {
      type: 'radar',
      data: {
        labels: SHORT_LABELS,
        datasets: [{
          data: dataVals,
          backgroundColor: 'rgba(16, 185, 129, 0.4)',
          borderColor: '#10b981',
          pointBackgroundColor: '#10b981',
          borderWidth: 1,
          pointRadius: 0
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: {
          r: {
            angleLines: { color: 'rgba(255,255,255,0.1)' }, grid: { color: 'rgba(255,255,255,0.1)' },
            pointLabels: { display: true, color: 'rgba(255,255,255,0.7)', font: { size: 9, family: 'Outfit' } },
            ticks: { display: false }
          }
        }
      }
    });
  });
}

function renderRankings() {
  rankingContainer.innerHTML = '';
  
  TARGET_COLUMNS.forEach((col, idx) => {
    // Sort songs by average descending
    const sorted = [...userData].sort((a, b) => b.averages[col] - a.averages[col]);
    const top3 = sorted.slice(0, 3);
    
    let listHtml = '';
    top3.forEach((song, i) => {
      const metaIndex = Math.min(song.songNumber - 1, SONG_METADATA.length - 1);
      const meta = SONG_METADATA[metaIndex];
      listHtml += `<li><span>${i+1}. ${meta.title}</span><span>${song.averages[col].toFixed(2)}</span></li>`;
    });
    
    rankingContainer.innerHTML += `
      <div class="ranking-category">
        <h4>${LABELS_KO[idx]} Top 3</h4>
        <ul class="ranking-list">${listHtml}</ul>
      </div>
    `;
  });
}

function selectSong(index) {
  document.querySelectorAll('.radar-item').forEach(el => el.classList.remove('active'));
  document.getElementById(`radar-item-${index}`).classList.add('active');
  
  const song = userData[index];
  const metaIndex = Math.min(index, SONG_METADATA.length - 1);
  const meta = SONG_METADATA[metaIndex];
  
  albumArt.textContent = meta.cover;
  songTitle.textContent = meta.title;
  songArtist.textContent = meta.artist;
  lineChartTitle.textContent = `Raw Trends: ${meta.title}`;
  
  renderLineChart(song);
}

function renderLineChart(song) {
  const ctx = document.getElementById('lineChart').getContext('2d');
  if (lineChartInstance) lineChartInstance.destroy();
  
  const labels = song.rawData.engagement.map((_, i) => i + 1);
  const colors = ['#3b82f6', '#eab308', '#ef4444', '#8b5cf6', '#10b981'];
  
  const datasets = TARGET_COLUMNS.map((col, i) => ({
    label: LABELS_KO[i],
    data: song.rawData[col],
    borderColor: colors[i],
    backgroundColor: 'transparent',
    borderWidth: 2, pointRadius: 0, tension: 0.4
  }));
  
  lineChartInstance = new Chart(ctx, {
    type: 'line', data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'top', labels: { color: '#f8fafc', font: { family: 'Outfit' } } } },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
        y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

// Team Comparison Logic
function loadTeamComparison() {
  teamScreen.classList.add('active');
  teamGrid.innerHTML = '';
  
  const teamStats = {};
  let hasData = false;
  
  // Calculate each member's overall stats for each metric
  Object.keys(MEMBERS).forEach(memberKey => {
    const saved = localStorage.getItem(`brainwaveData_${memberKey}`);
    if (saved) {
      hasData = true;
      const parsed = JSON.parse(saved);
      teamStats[memberKey] = {};
      
      TARGET_COLUMNS.forEach(col => {
        let maxVal = -Infinity;
        let minVal = Infinity;
        let sumVal = 0;
        
        parsed.forEach(song => {
          if (song.rawData && song.rawData[col]) {
            song.rawData[col].forEach(val => {
              if (val > maxVal) maxVal = val;
              if (val < minVal) minVal = val;
              sumVal += val;
            });
          }
        });
        
        // Handle case where no data was processed
        if (maxVal === -Infinity) maxVal = 0;
        if (minVal === Infinity) minVal = 0;
        
        teamStats[memberKey][col] = { max: maxVal, min: minVal, sum: sumVal };
      });
    }
  });
  
  if (!hasData) {
    teamGrid.innerHTML = `<p style="text-align:center; width: 100%; color: #ef4444;">No data available. Please upload data for team members first.</p>`;
    return;
  }

  const EMOTIONS = [
    { col: 'engagement', emoji: '🧠', title: '몰입 (Engagement)' },
    { col: 'interest', emoji: '👀', title: '흥미 (Interest)' },
    { col: 'excitement', emoji: '🔥', title: '활성 (Excitement)' },
    { col: 'stress', emoji: '🤯', title: '스트레스 (Stress)' },
    { col: 'relaxation', emoji: '🧘‍♂️', title: '이완 (Relaxation)' }
  ];

  const AWARDS = [
    { key: 'max', label: 'Highest Peak', findMax: true, desc: '가장 높은 수치를 기록' },
    { key: 'sum', label: 'Total Mass (Highest)', findMax: true, desc: '가장 많이 누적됨' },
    { key: 'min', label: 'Lowest Peak', findMax: false, desc: '가장 낮은 수치를 기록' },
    { key: 'sum', label: 'Total Mass (Lowest)', findMax: false, desc: '가장 적게 누적됨' }
  ];

  let htmlContent = '';

  EMOTIONS.forEach(emotion => {
    htmlContent += `<div style="width: 100%; margin-bottom: 3rem;">
      <h2 style="margin-bottom: 1rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; color: var(--accent-color);">
        ${emotion.emoji} ${emotion.title} 시상식
      </h2>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 1rem;">`;

    AWARDS.forEach(award => {
      let bestVal = award.findMax ? -Infinity : Infinity;
      let winnerKey = null;
      let scoresHtml = '';

      Object.keys(teamStats).forEach(mKey => {
        const val = teamStats[mKey][emotion.col][award.key];
        scoresHtml += `<div><strong>${MEMBERS[mKey]}:</strong> ${val.toLocaleString(undefined, {maximumFractionDigits: 2})}</div>`;
        
        if (award.findMax) {
          if (val > bestVal) { bestVal = val; winnerKey = mKey; }
        } else {
          if (val < bestVal) { bestVal = val; winnerKey = mKey; }
        }
      });

      if (winnerKey) {
        htmlContent += `
          <div class="king-card" style="padding: 1.5rem;">
            <div class="king-title" style="font-size: 1rem; color: #fff;">${award.label}</div>
            <div style="font-size: 0.8rem; color: var(--text-secondary); margin-bottom: 1rem;">${award.desc}</div>
            <div class="king-name" style="font-size: 1.5rem;">${MEMBERS[winnerKey]}</div>
            <div class="king-score" style="font-size: 1rem;">${bestVal.toLocaleString(undefined, {maximumFractionDigits: 2})}</div>
            <div class="king-details" style="padding: 0.8rem; font-size: 0.8rem;">${scoresHtml}</div>
          </div>
        `;
      }
    });

    htmlContent += `</div></div>`;
  });

  teamGrid.innerHTML = htmlContent;
}
