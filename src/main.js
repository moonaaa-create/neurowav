import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import preloadedData from './preloadedData.json';

// Initialize data instantly
Object.keys(preloadedData).forEach(key => {
  localStorage.setItem(`brainwaveData_${key}`, JSON.stringify(preloadedData[key]));
});

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
  <div id="homeScreen" class="screen active" style="position: relative; min-height: 80vh; justify-content: center;">
    <div class="floating-bg">
      <div class="floating-emoji" style="left: 10%; top: 20%; animation: floatWander1 15s ease-in-out infinite;">🎵</div>
      <div class="floating-emoji" style="left: 30%; top: 60%; animation: floatWander2 20s ease-in-out infinite; font-size: 4rem;">🎹</div>
      <div class="floating-emoji" style="left: 50%; top: 30%; animation: floatWander1 18s ease-in-out infinite;">🎸</div>
      <div class="floating-emoji" style="left: 70%; top: 70%; animation: floatWander2 22s ease-in-out infinite; font-size: 5rem;">🧠</div>
      <div class="floating-emoji" style="left: 85%; top: 25%; animation: floatWander1 16s ease-in-out infinite;">🎧</div>
      <div class="floating-emoji" style="left: 15%; top: 80%; animation: floatWander2 19s ease-in-out infinite;">🎶</div>
      <div class="floating-emoji" style="left: 50%; top: 80%; animation: floatWander1 17s ease-in-out infinite; font-size: 3.5rem;">⚡</div>
    </div>
    <div class="hero-section" style="text-align: center; padding: 2rem 1rem; animation: fadeIn 0.8s cubic-bezier(0.25, 1, 0.5, 1); z-index: 10; position: relative;">
      <h1 style="font-size: 5rem; margin-bottom: 1.5rem; background: linear-gradient(to right, #0a84ff, #64d2ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">NeuroWav</h1>
      <p class="subtitle" style="font-size: 1.3rem; max-width: 600px; margin: 0 auto 3.5rem auto; line-height: 1.8;">
        <strong>2026학년도 1학기 자료시각화 과제</strong><br>
        13곡의 음악 장르, 4명의 사용자.<br>
        뇌파 분석을 통한 음악 취향 알아내기!
      </p>
      
      <div style="display: flex; flex-direction: column; gap: 2rem; align-items: center;">
        <div>
          <h3 style="color: var(--text-secondary); margin-bottom: 1.2rem; font-weight: 500; letter-spacing: -0.5px;">멤버별 대시보드 보기</h3>
          <div style="display: flex; gap: 1.5rem; flex-wrap: wrap; justify-content: center;">
            <!-- 문경수 -->
            <div id="btn-member3" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1rem 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 130px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile1.png" style="width: 90px; height: 90px; border-radius: 12px; object-fit: cover; margin-bottom: 0.6rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff;">문경수 님</span>
            </div>
            <!-- 김한주 -->
            <div id="btn-member1" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1rem 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 130px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile2.png" style="width: 90px; height: 90px; border-radius: 12px; object-fit: cover; margin-bottom: 0.6rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff;">김한주 님</span>
            </div>
            <!-- 김용석 -->
            <div id="btn-member2" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1rem 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 130px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile3.png" style="width: 90px; height: 90px; border-radius: 12px; object-fit: cover; margin-bottom: 0.6rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff;">김용석 님</span>
            </div>
            <!-- 홍수민 -->
            <div id="btn-member4" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1rem 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 130px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile4.png" style="width: 90px; height: 90px; border-radius: 12px; object-fit: cover; margin-bottom: 0.6rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff;">홍수민 님</span>
            </div>
          </div>
        </div>
        
        <div style="width: 100%; max-width: 400px; height: 1px; background: rgba(255,255,255,0.1); margin: 1rem 0;"></div>
        
        <div>
          <h3 style="color: var(--text-secondary); margin-bottom: 1.2rem; font-weight: 500; letter-spacing: -0.5px;">종합 비교 및 시상식</h3>
          <div style="display: flex; gap: 1rem; flex-wrap: wrap; justify-content: center;">
            <button id="btn-awards-emotiv" class="hero-btn" style="background: linear-gradient(135deg, #ff9f0a, #ff375f); font-size: 1.1rem; padding: 0.8rem 1.8rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(255, 55, 95, 0.3);">🧠 Awards - Emotiv</button>
            <button id="btn-awards-music" class="hero-btn" style="background: linear-gradient(135deg, #0a84ff, #30d158); font-size: 1.1rem; padding: 0.8rem 1.8rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(48, 209, 88, 0.3);">🎵 Awards - Music</button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div id="dashboardScreen" class="screen">
    <header class="dash-header" style="display: flex; align-items: center; gap: 1rem; margin-bottom: 1.5rem;">
      <img id="headerProfileImg" src="/profile1.png" style="width: 48px; height: 48px; border-radius: 50%; border: 2px solid var(--accent-color); object-fit: cover; box-shadow: 0 2px 10px rgba(0,0,0,0.3);">
      <div class="user-info">
        <select id="userSelect" style="font-size: 1.5rem; font-weight: 700; border: none; background: transparent; color: var(--accent-color); cursor: pointer; outline: none;">
          <option value="member3">문경수's Dashboard</option>
          <option value="member1">김한주's Dashboard</option>
          <option value="member2">김용석's Dashboard</option>
          <option value="member4">홍수민's Dashboard</option>
        </select>
      </div>
    </header>

    <div class="dashboard-layout">
      <!-- Left Panel: Radar Grid + Song Info Merged -->
      <div class="panel radar-grid-panel">
        <div class="radar-grid" id="radarGrid">
          <div class="song-info-inline">
            <div class="album-art" id="albumArt">💿</div>
            <div style="flex-grow: 1;">
              <h2 id="songTitle">Select a Song</h2>
              <h3 id="songArtist">-</h3>
              <p id="songDesc" class="song-desc" style="margin-bottom: 0.5rem; display: none;"></p>
              <a id="spotifyLink" href="#" target="_blank" style="display: none; align-items: center; gap: 0.5rem; color: #1ed760; font-size: 0.85rem; font-weight: 600; text-decoration: none; padding: 0.4rem 0.8rem; background: rgba(30,215,96,0.1); border-radius: 20px; width: fit-content; transition: all 0.2s;">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0C5.4 0 0 5.4 0 12s5.4 12 12 12 12-5.4 12-12S18.66 0 12 0zm5.521 17.34c-.24.359-.66.48-1.021.24-2.82-1.74-6.36-2.101-10.561-1.141-.418.122-.779-.179-.899-.539-.12-.421.18-.78.54-.9 4.56-1.021 8.52-.6 11.64 1.32.42.18.479.659.301 1.02zm1.44-3.3c-.301.42-.841.6-1.262.3-3.239-1.98-8.159-2.58-11.939-1.38-.479.12-1.02-.12-1.14-.6-.12-.48.12-1.021.6-1.141C9.6 9.9 15 10.561 18.72 12.84c.361.181.54.78.241 1.2zm.12-3.36C15.24 8.4 8.82 8.16 5.16 9.301c-.6.179-1.2-.181-1.38-.721-.18-.6.18-1.2.72-1.38 4.26-1.26 11.28-1.02 15.721 1.621.539.3.719 1.02.419 1.56-.299.421-1.02.599-1.559.3z"/></svg>
                Listen on Spotify
              </a>
            </div>
          </div>
          <!-- Radar items injected here -->
        </div>
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
    <header class="team-header">
      <h1 id="awardsHeaderTitle">🏆 NeuroWav Awards</h1>
      <p id="awardsHeaderSubtitle" class="subtitle">Peak and Total Mass Analysis across all team members</p>
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

const navDashboard = document.getElementById('navDashboard');
const navTeamEmotiv = document.getElementById('navTeamEmotiv');
const navTeamMusic = document.getElementById('navTeamMusic');
const homeScreen = document.getElementById('homeScreen');
const dashboardScreen = document.getElementById('dashboardScreen');
const teamScreen = document.getElementById('teamScreen');

function switchScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-links button').forEach(b => b.classList.remove('active'));
  
  if (screenId === 'homeScreen') { homeScreen.classList.add('active'); }
  if (screenId === 'dashboardScreen') { navDashboard.classList.add('active'); loadDashboard(); }
  if (screenId === 'emotivAwards') { navTeamEmotiv.classList.add('active'); loadTeamComparison('emotiv'); }
  if (screenId === 'musicAwards') { navTeamMusic.classList.add('active'); loadTeamComparison('music'); }
}

navDashboard.addEventListener('click', () => switchScreen('dashboardScreen'));
navTeamEmotiv.addEventListener('click', () => switchScreen('emotivAwards'));
navTeamMusic.addEventListener('click', () => switchScreen('musicAwards'));

// Nav Logo click to home
document.querySelector('.nav-logo').style.cursor = 'pointer';
document.querySelector('.nav-logo').addEventListener('click', () => switchScreen('homeScreen'));

// Home screen button listeners
['member1', 'member2', 'member3', 'member4'].forEach(memberId => {
  document.getElementById(`btn-${memberId}`).addEventListener('click', () => {
    currentUserId = memberId;
    document.getElementById('userSelect').value = memberId;
    switchScreen('dashboardScreen');
  });
});

document.getElementById('btn-awards-emotiv').addEventListener('click', () => switchScreen('emotivAwards'));
document.getElementById('btn-awards-music').addEventListener('click', () => switchScreen('musicAwards'));

// Emoji Dragging Logic
let activeEmoji = null;
let emojiOffsetX = 0;
let emojiOffsetY = 0;

document.querySelectorAll('.floating-emoji').forEach(emoji => {
  emoji.addEventListener('mousedown', (e) => {
    activeEmoji = emoji;
    const rect = emoji.getBoundingClientRect();
    emojiOffsetX = e.clientX - rect.left;
    emojiOffsetY = e.clientY - rect.top;
    
    emoji.style.animation = 'none'; // Stop floating
    emoji.style.transform = 'none';
    emoji.style.left = rect.left + 'px';
    emoji.style.top = rect.top + 'px';
    emoji.style.opacity = '0.8';
    emoji.style.zIndex = '1000';
  });
});

document.addEventListener('mousemove', (e) => {
  if (!activeEmoji) return;
  activeEmoji.style.left = (e.clientX - emojiOffsetX) + 'px';
  activeEmoji.style.top = (e.clientY - emojiOffsetY) + 'px';
});

document.addEventListener('mouseup', () => {
  if (activeEmoji) {
    activeEmoji.style.opacity = '0.2';
    activeEmoji.style.zIndex = '';
    // Emojis stay where dropped (animation remains 'none')
    activeEmoji = null;
  }
});

// Global State
let currentUserId = 'member1';
let userData = {}; 
let lineChartInstance = null;

// DOM Elements
const userSelect = document.getElementById('userSelect');
const loadingOverlay = document.getElementById('loadingOverlay');

const albumArt = document.getElementById('albumArt');
const songTitle = document.getElementById('songTitle');
const songArtist = document.getElementById('songArtist');
const radarGrid = document.getElementById('radarGrid');
const lineChartTitle = document.getElementById('lineChartTitle');
const rankingContainer = document.getElementById('rankingContainer');
const teamGrid = document.getElementById('teamGrid');

userSelect.addEventListener('change', (e) => {
  currentUserId = e.target.value;
  loadDashboard();
});

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
  
  // File upload logic removed as data is preloaded.
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
  if (!savedData) return;
  
  dashboardScreen.classList.add('active');
  userData = JSON.parse(savedData);
  
  const headerProfileImg = document.getElementById('headerProfileImg');
  if (headerProfileImg) {
    if (currentUserId === 'member3') headerProfileImg.src = '/profile1.png';
    else if (currentUserId === 'member1') headerProfileImg.src = '/profile2.png';
    else if (currentUserId === 'member2') headerProfileImg.src = '/profile3.png';
    else if (currentUserId === 'member4') headerProfileImg.src = '/profile4.png';
  }
  
  renderRadarGrid();
  renderRankings();
  
  if (userData.length > 0) selectSong(0);
}

function renderRadarGrid() {
  // Clear previous radar items but keep the song-info block
  document.querySelectorAll('.radar-item').forEach(el => el.remove());
  
  // Use a fixed max of 1.0 for all radar charts to ensure they don't look "fat"
  // since the normalized data is between 0 and 1.
  const maxGlobal = 1.0;

  userData.forEach((song, index) => {
    const metaIndex = Math.min(index, SONG_METADATA.length - 1);
    const meta = SONG_METADATA[metaIndex];
    
    const div = document.createElement('div');
    div.className = 'radar-item';
    div.id = `radar-item-${index}`;
    div.innerHTML = `
      <div class="radar-wrapper"><canvas id="radar-${currentUserId}-${index}"></canvas></div>
      <div class="radar-title">${index + 1}. ${meta.title}</div>
    `;
    div.addEventListener('click', () => selectSong(index));
    radarGrid.appendChild(div);
    
    const ctx = document.getElementById(`radar-${currentUserId}-${index}`).getContext('2d');
    const dataVals = TARGET_COLUMNS.map(col => song.averages[col]);
    
    // Find the emotion with the highest average value
    const maxValIdx = dataVals.indexOf(Math.max(...dataVals));
    
    const emotionColors = [
      { border: '#3b82f6', bg: 'rgba(59, 130, 246, 0.35)' }, // Engagement (Blue)
      { border: '#eab308', bg: 'rgba(234, 179, 8, 0.35)' },  // Interest (Yellow)
      { border: '#ef4444', bg: 'rgba(239, 68, 68, 0.35)' },   // Excitement (Red)
      { border: '#8b5cf6', bg: 'rgba(139, 92, 246, 0.35)' },  // Stress (Purple)
      { border: '#10b981', bg: 'rgba(16, 185, 129, 0.35)' }   // Relaxation (Green)
    ];
    const colorSet = emotionColors[maxValIdx];
    
    new Chart(ctx, {
      type: 'radar',
      data: {
        labels: SHORT_LABELS,
        datasets: [{
          data: dataVals,
          backgroundColor: colorSet.bg,
          borderColor: colorSet.border,
          pointBackgroundColor: colorSet.border,
          borderWidth: 1.5,
          pointRadius: 0
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: {
          r: {
            min: 0,
            max: maxGlobal,
            angleLines: { color: 'rgba(255,255,255,0.1)' }, grid: { color: 'rgba(255,255,255,0.1)' },
            pointLabels: { display: true, color: 'rgba(255,255,255,0.7)', font: { size: 9, family: '-apple-system' } },
            ticks: { display: false }
          }
        }
      }
    });
  });
}

function renderRankings() {
  rankingContainer.innerHTML = '';
  
  const medals = ['🥇', '🥈', '🥉'];
  
  TARGET_COLUMNS.forEach((col, idx) => {
    // Sort songs by average descending
    const sorted = [...userData].sort((a, b) => b.averages[col] - a.averages[col]);
    const top3 = sorted.slice(0, 3);
    
    let listHtml = '';
    top3.forEach((song, i) => {
      const metaIndex = Math.min(song.songNumber - 1, SONG_METADATA.length - 1);
      const meta = SONG_METADATA[metaIndex];
      const medal = medals[i] || '🏅';
      listHtml += `<li><span style="display: flex; align-items: center; gap: 6px;"><span>${medal}</span> <span>${song.songNumber}. ${meta.title}</span></span><span>${song.averages[col].toFixed(2)}</span></li>`;
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
  
  const spotifyLink = document.getElementById('spotifyLink');
  if (spotifyLink) {
    spotifyLink.style.display = 'inline-flex';
    spotifyLink.href = `https://open.spotify.com/search/${encodeURIComponent(meta.title + ' ' + meta.artist)}`;
  }
  
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
      plugins: { legend: { position: 'top', labels: { color: '#f8fafc', font: { family: '-apple-system' } } } },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
        y: { min: 0, max: 1, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } }
      }
    },
    plugins: [{
      id: 'hoverLine',
      afterDraw: (chart) => {
        if (chart.tooltip && chart.tooltip.getActiveElements().length > 0) {
          const activeElement = chart.tooltip.getActiveElements()[0];
          const chartCtx = chart.ctx;
          const yAxis = chart.scales.y;
          const xPixel = activeElement.element.x;
          
          chartCtx.save();
          chartCtx.beginPath();
          chartCtx.strokeStyle = '#ff453a'; // Red vertical tracking line
          chartCtx.lineWidth = 1.5;
          chartCtx.setLineDash([4, 4]); // Dashed line
          chartCtx.moveTo(xPixel, yAxis.top);
          chartCtx.lineTo(xPixel, yAxis.bottom);
          chartCtx.stroke();
          chartCtx.restore();
        }
      }
    }]
  });
}

// Team Comparison Logic
function loadTeamComparison(type) {
  teamScreen.classList.add('active');
  teamGrid.innerHTML = '';
  
  const awardsHeaderTitle = document.getElementById('awardsHeaderTitle');
  const awardsHeaderSubtitle = document.getElementById('awardsHeaderSubtitle');

  // Change headers dynamically based on type
  if (type === 'emotiv') {
    awardsHeaderTitle.innerHTML = '🧠 NeuroWav Awards - Emotiv';
    awardsHeaderSubtitle.textContent = '조원들의 EEG 뇌파 수치를 바탕으로 한 멤버 시상식';
  } else {
    awardsHeaderTitle.innerHTML = '🎵 NeuroWav Awards - Music';
    awardsHeaderSubtitle.textContent = '조원 전체의 뇌파 평균치를 바탕으로 한 최고의 곡 시상식';
  }

  const teamStats = {};
  const songStats = {};
  let hasData = false;
  
  // Initialize songStats
  for (let i = 1; i <= 13; i++) {
    songStats[i] = {};
    TARGET_COLUMNS.forEach(col => {
      songStats[i][col] = { sum: 0, count: 0, max: -Infinity };
    });
  }
  
  // Calculate each member's overall stats for each metric
  Object.keys(MEMBERS).forEach(memberKey => {
    const saved = localStorage.getItem(`brainwaveData_${memberKey}`);
    if (saved) {
      hasData = true;
      const parsed = JSON.parse(saved);
      teamStats[memberKey] = {};
      
      TARGET_COLUMNS.forEach(col => {
        let maxVal = -Infinity;
        let maxSongTitle = '';
        let maxSongIdx = 1;
        let maxSec = 0;
        
        let minVal = Infinity;
        let minSongTitle = '';
        let minSongIdx = 1;
        let minSec = 0;
        
        let sumVal = 0;
        
        // To track which song contributed the most/least to the sum
        let bestSongSum = -Infinity;
        let bestSongTitle = '';
        let bestSongIdx = 1;
        let worstSongSum = Infinity;
        let worstSongTitle = '';
        let worstSongIdx = 1;
        
        parsed.forEach(song => {
          const sIdx = song.songNumber;
          const meta = SONG_METADATA[sIdx - 1] || { title: `Song ${sIdx}` };
          
          let songSum = 0;
          let songPoints = 0;
          
          if (song.rawData && song.rawData[col]) {
            song.rawData[col].forEach((val, secIdx) => {
              if (val > maxVal) {
                maxVal = val;
                maxSongTitle = `${sIdx}. ${meta.title}`;
                maxSongIdx = sIdx;
                maxSec = secIdx;
              }
              if (val < minVal) {
                minVal = val;
                minSongTitle = `${sIdx}. ${meta.title}`;
                minSongIdx = sIdx;
                minSec = secIdx;
              }
              sumVal += val;
              songSum += val;
              songPoints++;
              
              if (songStats[sIdx] && val > songStats[sIdx][col].max) {
                songStats[sIdx][col].max = val;
              }
            });
          }
          
          if (songPoints > 0) {
            if (songSum > bestSongSum) {
              bestSongSum = songSum;
              bestSongTitle = `${sIdx}. ${meta.title}`;
              bestSongIdx = sIdx;
            }
            if (songSum < worstSongSum) {
              worstSongSum = songSum;
              worstSongTitle = `${sIdx}. ${meta.title}`;
              worstSongIdx = sIdx;
            }
          }
          
          if (song.averages && song.averages[col] !== undefined && songStats[sIdx]) {
            songStats[sIdx][col].sum += song.averages[col];
            songStats[sIdx][col].count++;
          }
        });
        
        // Handle case where no data was processed
        if (maxVal === -Infinity) maxVal = 0;
        if (minVal === Infinity) minVal = 0;
        
        teamStats[memberKey][col] = {
          max: maxVal,
          maxSong: maxSongTitle,
          maxSongIdx: maxSongIdx,
          maxTime: maxSec,
          min: minVal,
          minSong: minSongTitle,
          minSongIdx: minSongIdx,
          minTime: minSec,
          sum: sumVal,
          bestSong: bestSongTitle,
          bestSongIdx: bestSongIdx,
          worstSong: worstSongTitle,
          worstSongIdx: worstSongIdx
        };
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
    { key: 'max', label: 'Highest Peak', findMax: true, desc: '가장 높은 수치를 기록', type: 'max' },
    { key: 'sum', label: 'Total Mass (Highest)', findMax: true, desc: '가장 많이 누적됨', type: 'sum_high' },
    { key: 'min', label: 'Lowest Peak', findMax: false, desc: '가장 낮은 수치를 기록', type: 'min' },
    { key: 'sum', label: 'Total Mass (Lowest)', findMax: false, desc: '가장 적게 누적됨', type: 'sum_low' }
  ];

  let htmlContent = '';

  if (type === 'emotiv') {
    EMOTIONS.forEach(emotion => {
      htmlContent += `<div style="width: 100%; margin-bottom: 3rem;">
        <h2 style="margin-bottom: 1rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; color: var(--accent-color);">
          ${emotion.emoji} ${emotion.title} 시상식
        </h2>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 1rem;">`;

      AWARDS.forEach(award => {
        const results = [];
        
        Object.keys(teamStats).forEach(mKey => {
          const val = teamStats[mKey][emotion.col][award.key];
          const detail = teamStats[mKey][emotion.col];
          results.push({ mKey, val, detail });
        });

        if (award.findMax) {
          results.sort((a, b) => b.val - a.val);
        } else {
          results.sort((a, b) => a.val - b.val);
        }

        const medals = ['🥇', '🥈', '🥉', '🏅'];
        
        let top3Html = '<div style="margin-top: 1rem; text-align: left; background: rgba(0,0,0,0.3); padding: 1rem; border-radius: var(--radius-md); font-size: 0.95rem;">';
        results.forEach((r, idx) => {
          let detailText = '';
          if (award.type === 'max') {
            detailText = `${r.detail.maxSong} (약 ${r.detail.maxTime}초 부근)`;
          } else if (award.type === 'min') {
            detailText = `${r.detail.minSong} (약 ${r.detail.minTime}초 부근)`;
          } else if (award.type === 'sum_high') {
            detailText = `최다 기여 곡: ${r.detail.bestSong}`;
          } else if (award.type === 'sum_low') {
            detailText = `최소 기여 곡: ${r.detail.worstSong}`;
          }

          const isClickable = award.type === 'max' || award.type === 'min';
          const rowId = isClickable ? `award-row-${emotion.col}-${award.type}-${r.mKey}` : '';
          const clickTip = isClickable ? `<span style="font-size: 0.65rem; color: var(--accent-color); margin-left: 4px;">(클릭 시 분석 차트 보기 📈)</span>` : '';
          const rowStyle = isClickable 
            ? 'cursor: pointer; padding: 0.4rem; border-radius: 6px; transition: background-color 0.2s;' 
            : 'padding: 0.4rem; border-radius: 6px;';
          const hoverAttr = isClickable 
            ? `onmouseover="this.style.backgroundColor='rgba(255,255,255,0.03)'" onmouseout="this.style.backgroundColor='transparent'"` 
            : '';

          let profileImgSrc = '';
          if (r.mKey === 'member3') profileImgSrc = '/profile1.png';
          else if (r.mKey === 'member1') profileImgSrc = '/profile2.png';
          else if (r.mKey === 'member2') profileImgSrc = '/profile3.png';
          else if (r.mKey === 'member4') profileImgSrc = '/profile4.png';

          top3Html += `
            <div id="${rowId}" style="display: flex; flex-direction: column; margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba(255,255,255,0.05); padding-bottom: 0.5rem; ${rowStyle}" ${hoverAttr}>
              <div style="display: flex; justify-content: space-between; align-items: center;">
                <span style="color: var(--text-primary); font-weight: 500; display: flex; align-items: center; gap: 8px;">
                  <span style="width: 28px; text-align: left; font-size: 1.2rem; display: inline-block;">${medals[idx]}</span> 
                  <img src="${profileImgSrc}" style="width: 28px; height: 28px; border-radius: 50%; object-fit: cover; border: 1px solid rgba(255,255,255,0.2);">
                  <span style="display: inline-block;">${MEMBERS[r.mKey]}</span>
                </span>
                <span style="color: var(--text-secondary); font-variant-numeric: tabular-nums;">
                  ${r.val.toLocaleString(undefined, {maximumFractionDigits: 2})}
                </span>
              </div>
              <div style="font-size: 0.78rem; color: #8a99ad; margin-left: 28px; margin-top: 0.15rem; font-weight: 400; display: flex; align-items: center; gap: 4px;">
                <span>${detailText}</span>
                ${clickTip}
              </div>
            </div>`;
        });
        // Remove last margin/border for clean UI
        top3Html = top3Html.replace(/margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba\(255,255,255,0.05\); padding-bottom: 0.5rem;"(?!.*margin-bottom)/, 'margin-bottom: 0;"');
        top3Html += '</div>';

        htmlContent += `
          <div class="king-card" style="padding: 1.5rem;">
            <div class="king-title" style="font-size: 1.05rem; color: #fff; font-weight: 600;">${award.label}</div>
            <div style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 1.5rem;">${award.desc}</div>
            ${top3Html}
          </div>
        `;
      });

      htmlContent += `</div>`; // close grid
      
      // Shared wide chart container below the grid!
      htmlContent += `
        <div id="section-chart-container-${emotion.col}" style="display: none; width: 100%; height: 280px; margin-top: 1.5rem; padding: 1.2rem; background: rgba(0,0,0,0.4); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); position: relative;">
          <h3 id="section-chart-title-${emotion.col}" style="font-size: 0.95rem; color: #fff; margin-bottom: 0.8rem; font-weight: 600; text-align: left;"></h3>
          <div style="height: calc(100% - 2.5rem); position: relative;">
            <canvas id="section-chart-canvas-${emotion.col}"></canvas>
          </div>
        </div>
      </div>`; // close emotion section block
    });
  } else if (type === 'music') {
    EMOTIONS.forEach(emotion => {
      htmlContent += `<div style="width: 100%; margin-bottom: 4rem;">
        <h2 style="margin-bottom: 1rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; color: var(--accent-color);">
          ${emotion.emoji} 최고의 ${emotion.title} 곡
        </h2>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1rem;">`;

      const songResults = [];
      Object.keys(songStats).forEach(sIdx => {
        const stats = songStats[sIdx][emotion.col];
        if (stats.count > 0) {
          songResults.push({
            sIdx: parseInt(sIdx),
            avg: stats.sum / stats.count,
            max: stats.max !== -Infinity ? stats.max : 0
          });
        }
      });
      
      const medals = ['🥇', '🥈', '🥉'];
      
      // Sort for Average
      const avgSorted = [...songResults].sort((a, b) => b.avg - a.avg);
      let avgTop3Html = '<div style="margin-top: 1rem; text-align: left; background: rgba(0,0,0,0.3); padding: 1rem; border-radius: var(--radius-md); font-size: 0.95rem;">';
      avgSorted.slice(0, 3).forEach((r, idx) => {
        const meta = SONG_METADATA[r.sIdx - 1]; 
        avgTop3Html += `
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba(255,255,255,0.05); padding-bottom: 0.5rem;">
            <span style="color: var(--text-primary); font-weight: 500; display: flex; align-items: center;">
              <span style="width: 28px; text-align: left; font-size: 1.2rem; display: inline-block;">${medals[idx]}</span> 
              <span style="display: inline-block;">${r.sIdx}. ${meta ? meta.title : 'Unknown'}</span>
            </span>
            <span style="color: var(--text-secondary); font-variant-numeric: tabular-nums;">
              ${r.avg.toLocaleString(undefined, {maximumFractionDigits: 2})}
            </span>
          </div>`;
      });
      avgTop3Html = avgTop3Html.replace(/margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba\(255,255,255,0.05\); padding-bottom: 0.5rem;"(?!.*margin-bottom)/, 'margin-bottom: 0;"');
      avgTop3Html += '</div>';

      // Sort for Peak (Max)
      const maxSorted = [...songResults].sort((a, b) => b.max - a.max);
      let maxTop3Html = '<div style="margin-top: 1rem; text-align: left; background: rgba(0,0,0,0.3); padding: 1rem; border-radius: var(--radius-md); font-size: 0.95rem;">';
      maxSorted.slice(0, 3).forEach((r, idx) => {
        const meta = SONG_METADATA[r.sIdx - 1]; 
        maxTop3Html += `
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba(255,255,255,0.05); padding-bottom: 0.5rem;">
            <span style="color: var(--text-primary); font-weight: 500; display: flex; align-items: center;">
              <span style="width: 28px; text-align: left; font-size: 1.2rem; display: inline-block;">${medals[idx]}</span> 
              <span style="display: inline-block;">${r.sIdx}. ${meta ? meta.title : 'Unknown'}</span>
            </span>
            <span style="color: var(--text-secondary); font-variant-numeric: tabular-nums;">
              ${r.max.toLocaleString(undefined, {maximumFractionDigits: 2})}
            </span>
          </div>`;
      });
      maxTop3Html = maxTop3Html.replace(/margin-bottom: 0.8rem; border-bottom: 0.5px solid rgba\(255,255,255,0.05\); padding-bottom: 0.5rem;"(?!.*margin-bottom)/, 'margin-bottom: 0;"');
      maxTop3Html += '</div>';

      htmlContent += `
        <div class="king-card" style="padding: 1.5rem;">
          <div class="king-title" style="font-size: 1.05rem; color: #fff; font-weight: 600;">평균 수치 최고 (Average)</div>
          <div style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 1.5rem;">노래 재생 내내 꾸준히 높았던 곡</div>
          ${avgTop3Html}
        </div>
        <div class="king-card" style="padding: 1.5rem;">
          <div class="king-title" style="font-size: 1.05rem; color: #fff; font-weight: 600;">순간 최고치 (Peak)</div>
          <div style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 1.5rem;">특정 순간 엄청난 폭발력을 보여준 곡</div>
          ${maxTop3Html}
        </div>
      `;

      htmlContent += `</div>`;
      
      // AI Insight
      const topAvgSong = SONG_METADATA[avgSorted[0].sIdx - 1]?.title;
      const topMaxSong = SONG_METADATA[maxSorted[0].sIdx - 1]?.title;
      
      let insightText = '';
      if (topAvgSong === topMaxSong) {
        insightText = `💡 <strong>NeuroWav Insight:</strong> 역시 <strong>${topAvgSong}</strong>! 꾸준히 높은 수치를 유지하면서 순간 최고치까지 1위를 차지했습니다. 우리 팀에게 완벽하게 통하는 마성의 곡이네요! 👑`;
      } else {
        insightText = `💡 <strong>NeuroWav Insight:</strong> <strong>${topAvgSong}</strong>이(가) 곡 전체에 걸쳐 잔잔한 여운을 남겼다면, <strong>${topMaxSong}</strong>은(는) 뇌리를 스치는 강력한 킬링 파트가 있는 폭발적인 곡입니다! 🔥`;
      }
      
      htmlContent += `
        <div style="width: 100%; margin-top: 1rem; padding: 1.2rem; background: rgba(10,132,255,0.1); border-left: 4px solid var(--accent-color); border-radius: 8px; color: #fff; font-size: 0.95rem; line-height: 1.5;">
          ${insightText}
        </div>
      </div>`;
    });
  }

  teamGrid.innerHTML = htmlContent;

  if (type === 'emotiv') {
    EMOTIONS.forEach(emotion => {
      let activeTarget = null; // Track active row { awardType, mKey }
      
      AWARDS.forEach(award => {
        if (award.type !== 'max' && award.type !== 'min') return;
        
        Object.keys(teamStats).forEach(mKey => {
          const detail = teamStats[mKey][emotion.col];
          const rowId = `award-row-${emotion.col}-${award.type}-${mKey}`;
          
          const rowEl = document.getElementById(rowId);
          if (rowEl) {
            rowEl.addEventListener('click', () => {
              const container = document.getElementById(`section-chart-container-${emotion.col}`);
              const titleEl = document.getElementById(`section-chart-title-${emotion.col}`);
              const canvas = document.getElementById(`section-chart-canvas-${emotion.col}`);
              
              if (activeTarget && activeTarget.awardType === award.type && activeTarget.mKey === mKey) {
                container.style.display = 'none';
                activeTarget = null;
                return;
              }
              
              const savedData = localStorage.getItem(`brainwaveData_${mKey}`);
              if (!savedData) return;
              
              const parsed = JSON.parse(savedData);
              const songIdx = (award.type === 'max') ? detail.maxSongIdx : detail.minSongIdx;
              const songObj = parsed.find(s => s.songNumber === songIdx);
              if (!songObj) return;
              
              container.style.display = 'block';
              activeTarget = { awardType: award.type, mKey };
              
              const songTitle = (award.type === 'max') ? detail.maxSong : detail.minSong;
              const emotionTitle = (award.type === 'max') ? 'Highest Peak' : 'Lowest Peak';
              titleEl.innerHTML = `📊 <strong>${MEMBERS[mKey]}</strong> 님의 <span style="color: var(--accent-color);">${songTitle}</span> 전체 감정 분석 타임라인 <span style="font-size: 0.8rem; color: var(--text-secondary); font-weight: normal; margin-left: 0.5rem;">(기준 지표: ${emotion.title} - ${emotionTitle})</span>`;
              
              if (canvas.chartInstance) {
                canvas.chartInstance.destroy();
              }
              
              const labels = songObj.rawData.engagement.map((_, i) => `${i}초`);
              const ctx = canvas.getContext('2d');
              const highlightSec = (award.type === 'max') ? detail.maxTime : detail.minTime;
              const colors = ['#3b82f6', '#eab308', '#ef4444', '#8b5cf6', '#10b981'];
              
              const datasets = TARGET_COLUMNS.map((col, i) => {
                const isTargetCol = col === emotion.col;
                return {
                  label: LABELS_KO[i],
                  data: songObj.rawData[col],
                  borderColor: colors[i],
                  backgroundColor: 'transparent',
                  borderWidth: isTargetCol ? 2.5 : 1,
                  pointRadius: 0,
                  tension: 0.35
                };
              });
              
              canvas.chartInstance = new Chart(ctx, {
                type: 'line',
                data: { labels, datasets },
                options: {
                  responsive: true,
                  maintainAspectRatio: false,
                  interaction: { mode: 'index', intersect: false },
                  plugins: {
                    legend: {
                      position: 'top',
                      labels: { color: '#f8fafc', font: { family: '-apple-system', size: 10 } }
                    },
                    tooltip: { enabled: true }
                  },
                  scales: {
                    x: { grid: { color: 'rgba(255,255,255,0.03)' }, ticks: { color: '#94a3b8', font: { size: 9 } } },
                    y: { min: 0, max: 1, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', font: { size: 9 } } }
                  }
                },
                plugins: [{
                  id: 'verticalLine',
                  afterDraw: (chart) => {
                    if (highlightSec >= 0) {
                      const chartCtx = chart.ctx;
                      const xAxis = chart.scales.x;
                      const yAxis = chart.scales.y;
                      
                      const xPixel = xAxis.getPixelForValue(highlightSec);
                      if (xPixel !== undefined) {
                        chartCtx.save();
                        chartCtx.beginPath();
                        chartCtx.strokeStyle = '#ff453a'; // Red vertical line
                        chartCtx.lineWidth = 1.5;
                        chartCtx.setLineDash([5, 5]); // Dashed
                        chartCtx.moveTo(xPixel, yAxis.top);
                        chartCtx.lineTo(xPixel, yAxis.bottom);
                        chartCtx.stroke();
                        
                        // Marker point text
                        chartCtx.fillStyle = '#ff453a';
                        chartCtx.font = 'bold 9px -apple-system';
                        chartCtx.textAlign = 'center';
                        chartCtx.fillText(`${highlightSec}초 사건 지점 📍`, xPixel, yAxis.top - 6);
                        chartCtx.restore();
                      }
                    }
                  }
                }]
              });
            });
          }
        });
      });
    });
  }
}

// Initially show home screen
switchScreen('homeScreen');
