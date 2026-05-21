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
            <div id="btn-member3" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 140px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile1.png" style="width: 95px; height: 95px; border-radius: 12px; object-fit: cover; margin-bottom: 0.8rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff; margin-bottom: 0.3rem;">문경수 님</span>
              <span style="font-size: 0.72rem; color: var(--text-secondary); text-align: center; line-height: 1.3;">휴먼AI공학전공<br>21학번</span>
            </div>
            <!-- 김한주 -->
            <div id="btn-member1" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 140px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile2.png" style="width: 95px; height: 95px; border-radius: 12px; object-fit: cover; margin-bottom: 0.8rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff; margin-bottom: 0.3rem;">김한주 님</span>
              <span style="font-size: 0.72rem; color: var(--text-secondary); text-align: center; line-height: 1.3;">음악학과 작곡과<br>21학번</span>
            </div>
            <!-- 김용석 -->
            <div id="btn-member2" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 140px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile3.png" style="width: 95px; height: 95px; border-radius: 12px; object-fit: cover; margin-bottom: 0.8rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff; margin-bottom: 0.3rem;">김용석 님</span>
              <span style="font-size: 0.72rem; color: var(--text-secondary); text-align: center; line-height: 1.3;">자유전공학부 인문<br>26학번</span>
            </div>
            <!-- 홍수민 -->
            <div id="btn-member4" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255,255,255,0.03); padding: 1.2rem; border-radius: 16px; border: 1px solid rgba(255,255,255,0.08); width: 140px; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
              <img src="/profile4.png" style="width: 95px; height: 95px; border-radius: 12px; object-fit: cover; margin-bottom: 0.8rem; border: 1.5px solid rgba(255,255,255,0.15);">
              <span style="font-size: 0.95rem; font-weight: 600; color: #fff; margin-bottom: 0.3rem;">홍수민 님</span>
              <span style="font-size: 0.72rem; color: var(--text-secondary); text-align: center; line-height: 1.3;">조형예술학과<br>25학번</span>
            </div>
          </div>
        </div>
        
        <div style="width: 100%; max-width: 400px; height: 1px; background: rgba(255,255,255,0.1); margin: 1rem 0;"></div>
        
        <div>
          <h3 style="color: var(--text-secondary); margin-bottom: 1.2rem; font-weight: 500; letter-spacing: -0.5px;">종합 비교 및 시상식</h3>
          <div style="display: flex; gap: 1rem; flex-wrap: wrap; justify-content: center;">
            <button id="btn-song-analysis" class="hero-btn" style="background: linear-gradient(135deg, #bf5af2, #0a84ff); font-size: 1.1rem; padding: 0.8rem 1.8rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(191, 90, 242, 0.3);">🎵 Song Dashboard</button>
            <button id="btn-awards-emotiv" class="hero-btn" style="background: linear-gradient(135deg, #ff9f0a, #ff375f); font-size: 1.1rem; padding: 0.8rem 1.8rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(255, 55, 95, 0.3);">🧠 Awards - Emotiv</button>
            <button id="btn-awards-music" class="hero-btn" style="background: linear-gradient(135deg, #0a84ff, #30d158); font-size: 1.1rem; padding: 0.8rem 1.8rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(48, 209, 88, 0.3);">🎵 Awards - Music</button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div id="dashboardScreen" class="screen">
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.5rem; border-bottom: 0.5px solid var(--border-color); padding-bottom: 1rem; flex-wrap: wrap; gap: 1rem;">
      <header class="dash-header" style="display: flex; align-items: center; gap: 1rem; border-bottom: none; padding-bottom: 0; margin-bottom: 0;">
        <img id="headerProfileImg" src="/profile1.png" style="width: 48px; height: 48px; border-radius: 50%; border: 2px solid var(--accent-color); object-fit: cover; box-shadow: 0 2px 10px rgba(0,0,0,0.3);">
        <div class="user-info">
          <select id="userSelect" style="font-size: 1.5rem; font-weight: 700; border: none; background: transparent; color: var(--accent-color); cursor: pointer; outline: none; margin-bottom: 0.15rem;">
            <option value="member3">문경수's Dashboard</option>
            <option value="member1">김한주's Dashboard</option>
            <option value="member2">김용석's Dashboard</option>
            <option value="member4">홍수민's Dashboard</option>
          </select>
          <div id="headerMemberInfo" style="font-size: 0.82rem; color: var(--text-secondary); padding-left: 0.25rem;"></div>
        </div>
      </header>
      
      <div class="toggle-switch-container">
        <button id="btnViewPersonal" class="toggle-btn active">👤 개인 분석</button>
        <button id="btnViewComparative" class="toggle-btn">👥 조원 비교</button>
      </div>
    </div>

    <div class="dashboard-layout">
      <!-- Left Panel: Radar Grid + Song Info Merged -->
      <div class="panel radar-grid-panel" id="personalRadarPanel">
        <div class="radar-grid" id="radarGrid">
          <div class="song-info-inline">
            <div class="album-art" id="albumArt">💿</div>
            <div style="flex-grow: 1;">
              <h2 id="songTitle">Select a Song</h2>
              <h3 id="songArtist">-</h3>
              <p id="songDesc" class="song-desc" style="margin-bottom: 0.5rem; display: none;"></p>
              
              <div style="display: flex; align-items: center; gap: 1rem; flex-wrap: wrap; margin-top: 0.5rem;">
                <a id="spotifyLink" href="#" target="_blank" style="display: none; align-items: center; gap: 0.5rem; color: #1ed760; font-size: 0.85rem; font-weight: 600; text-decoration: none; padding: 0.4rem 0.8rem; background: rgba(30,215,96,0.1); border-radius: 20px; transition: all 0.2s;">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0C5.4 0 0 5.4 0 12s5.4 12 12 12 12-5.4 12-12S18.66 0 12 0zm5.521 17.34c-.24.359-.66.48-1.021.24-2.82-1.74-6.36-2.101-10.561-1.141-.418.122-.779-.179-.899-.539-.12-.421.18-.78.54-.9 4.56-1.021 8.52-.6 11.64 1.32.42.18.479.659.301 1.02zm1.44-3.3c-.301.42-.841.6-1.262.3-3.239-1.98-8.159-2.58-11.939-1.38-.479.12-1.02-.12-1.14-.6-.12-.48.12-1.021.6-1.141C9.6 9.9 15 10.561 18.72 12.84c.361.181.54.78.241 1.2zm.12-3.36C15.24 8.4 8.82 8.16 5.16 9.301c-.6.179-1.2-.181-1.38-.721-.18-.6.18-1.2.72-1.38 4.26-1.26 11.28-1.02 15.721 1.621.539.3.719 1.02.419 1.56-.299.421-1.02.599-1.559.3z"/></svg>
                  Listen on Spotify
                </a>
              </div>

              <!-- Interactive Mock Audio Player -->
              <div class="mock-player" id="mockAudioPlayer" style="display: none;">
                <button id="playerPlayBtn" class="player-btn">▶ Play</button>
                <button id="playerStopBtn" class="player-btn">■ Stop</button>
                <div class="player-progress-container" id="playerProgressContainer">
                  <div class="player-progress-bar" id="playerProgressBar"></div>
                </div>
                <span class="player-time" id="playerTime">00:00</span>
                <!-- Equalizer visualizer -->
                <div class="eq-container" id="playerEq">
                  <div class="eq-bar"></div>
                  <div class="eq-bar"></div>
                  <div class="eq-bar"></div>
                  <div class="eq-bar"></div>
                </div>
              </div>
            </div>
          </div>
          <!-- Radar items injected here -->
        </div>
      </div>

      <!-- Left Panel (Comparative Mode): Multi-user Line Chart -->
      <div class="panel comparison-panel" id="comparativePanel" style="display: none; grid-column: 1 / 2; grid-row: 1 / 2;">
        <div class="comparison-selector-panel">
          <span style="font-weight: 600; font-size: 0.95rem;">👥 조원 전체 감정 비교</span>
          <select id="compareMetricSelect">
            <option value="engagement">🧠 몰입도 (Engagement)</option>
            <option value="interest">👀 흥미도 (Interest)</option>
            <option value="excitement">🔥 활성도 (Excitement)</option>
            <option value="stress">🤯 스트레스 (Stress)</option>
            <option value="relaxation">🧘‍♂️ 이완도 (Relaxation)</option>
          </select>
        </div>
        <div style="position: relative; height: 350px; width: 100%;">
          <canvas id="compareChart"></canvas>
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

  <div id="songAnalysisScreen" class="screen">
    <header class="dash-header" style="margin-bottom: 1.5rem; border-bottom: 0.5px solid var(--border-color); padding-bottom: 1rem;">
      <div class="user-info" style="display: flex; flex-direction: column; gap: 0.2rem;">
        <span style="color: var(--accent-color); font-size: 1.5rem; font-weight: 700;">🎵 Song Dashboard</span>
        <div style="font-size: 0.82rem; color: var(--text-secondary);">Select a song from the track list to compare 4 members' emotional responses in real-time.</div>
      </div>
    </header>

    <div class="song-analysis-layout">
      <!-- Left Panel: 13 Songs Selection Card List -->
      <div class="panel" style="display: flex; flex-direction: column; gap: 1rem; padding: 1.2rem; max-height: 800px;">
        <h3 style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.2rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem; display: flex; align-items: center; gap: 0.5rem;">
          💿 Tracks <span style="font-size: 0.8rem; font-weight: 400; color: var(--text-secondary);">(13 total)</span>
        </h3>
        <div class="song-card-list" id="songCardList">
          <!-- Dynamically populated song cards -->
        </div>
      </div>

      <!-- Right Panel: Visualizer & AI Summary -->
      <div class="song-analysis-main">
        <!-- Chart Controls -->
        <div class="chart-control-bar">
          <div style="display: flex; flex-direction: column; gap: 0.2rem; min-width: 0;">
            <h2 id="songAnalysisTitle" style="font-size: 1.4rem; font-weight: 700; color: #fff; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">Song Title</h2>
            <div id="songAnalysisArtist" style="font-size: 0.9rem; color: var(--text-secondary); font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">Artist</div>
          </div>
          
          <div class="mini-toggle-container">
            <button id="btnSongChartRadar" class="mini-toggle-btn active">🕸️ Radar Overlap</button>
            <button id="btnSongChartBar" class="mini-toggle-btn">📊 Grouped Bar</button>
          </div>
        </div>

        <!-- Big Comparison Chart Panel -->
        <div class="panel" style="position: relative; padding: 1.5rem; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 420px; width: 100%;">
          <div style="position: relative; height: 380px; width: 100%; max-width: 650px;">
            <canvas id="songAnalysisChart"></canvas>
          </div>
        </div>

        <!-- Emotional Insight Panel (AI Summary) -->
        <div class="panel ai-insight-panel" style="padding: 1.5rem;">
          <h3 style="font-size: 1.1rem; font-weight: 600; display: flex; align-items: center; gap: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem; margin-bottom: 0.5rem;">
            🧠 Emotional Metrics & Top Responders
          </h3>
          <div class="insight-grid" id="insightGrid">
            <!-- Dynamic insight cards for each 5 metrics -->
          </div>
        </div>
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
const navSongAnalysis = document.getElementById('navSongAnalysis');
const navTeamEmotiv = document.getElementById('navTeamEmotiv');
const navTeamMusic = document.getElementById('navTeamMusic');
const homeScreen = document.getElementById('homeScreen');
const dashboardScreen = document.getElementById('dashboardScreen');
const songAnalysisScreen = document.getElementById('songAnalysisScreen');
const teamScreen = document.getElementById('teamScreen');

function switchScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-links button').forEach(b => b.classList.remove('active'));
  
  if (screenId === 'homeScreen') { homeScreen.classList.add('active'); }
  if (screenId === 'dashboardScreen') { navDashboard.classList.add('active'); loadDashboard(); }
  if (screenId === 'songAnalysis') { navSongAnalysis.classList.add('active'); songAnalysisScreen.classList.add('active'); loadSongAnalysis(); triggerConfetti(); }
  if (screenId === 'emotivAwards') { navTeamEmotiv.classList.add('active'); loadTeamComparison('emotiv'); triggerConfetti(); }
  if (screenId === 'musicAwards') { navTeamMusic.classList.add('active'); loadTeamComparison('music'); triggerConfetti(); }
}

navDashboard.addEventListener('click', () => switchScreen('dashboardScreen'));
navSongAnalysis.addEventListener('click', () => switchScreen('songAnalysis'));
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

document.getElementById('btn-song-analysis').addEventListener('click', () => switchScreen('songAnalysis'));
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
let compareChartInstance = null;
let currentDashboardView = 'personal'; // 'personal' or 'comparative'

// Interactive Mock Player State
let isPlaying = false;
let playbackTime = 0; // in seconds
let playbackInterval = null;
let playbackMax = 30; // default max seconds
let activeSongIndex = 0; // index of currently selected song

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

// Newly added DOM Elements for Mock Player & Tab toggles
const mockAudioPlayer = document.getElementById('mockAudioPlayer');
const playerPlayBtn = document.getElementById('playerPlayBtn');
const playerStopBtn = document.getElementById('playerStopBtn');
const playerProgressBar = document.getElementById('playerProgressBar');
const playerProgressContainer = document.getElementById('playerProgressContainer');
const playerTime = document.getElementById('playerTime');
const playerEq = document.getElementById('playerEq');

const btnViewPersonal = document.getElementById('btnViewPersonal');
const btnViewComparative = document.getElementById('btnViewComparative');
const personalRadarPanel = document.getElementById('personalRadarPanel');
const comparativePanel = document.getElementById('comparativePanel');
const compareMetricSelect = document.getElementById('compareMetricSelect');

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
  
  // Reset tabs to Personal view by default on load
  if (btnViewPersonal && btnViewComparative) {
    btnViewPersonal.classList.add('active');
    btnViewComparative.classList.remove('active');
  }
  if (personalRadarPanel && comparativePanel) {
    personalRadarPanel.style.display = 'flex';
    comparativePanel.style.display = 'none';
  }
  currentDashboardView = 'personal';

  const headerProfileImg = document.getElementById('headerProfileImg');
  const headerMemberInfo = document.getElementById('headerMemberInfo');
  
  const memberDetails = {
    member3: { img: '/profile1.png', info: '휴먼AI공학전공 | 21학번' },
    member1: { img: '/profile2.png', info: '음악학과 작곡과 | 21학번' },
    member2: { img: '/profile3.png', info: '자유전공학부 인문 | 26학번' },
    member4: { img: '/profile4.png', info: '조형예술학과 | 25학번' }
  };
  
  if (headerProfileImg && headerMemberInfo) {
    const detail = memberDetails[currentUserId];
    if (detail) {
      headerProfileImg.src = detail.img;
      headerMemberInfo.textContent = detail.info;
    }
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
  activeSongIndex = index;
  
  document.querySelectorAll('.radar-item').forEach(el => el.classList.remove('active'));
  const radarItem = document.getElementById(`radar-item-${index}`);
  if (radarItem) radarItem.classList.add('active');
  
  const song = userData[index];
  if (!song) return;

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

  // Display and reset player for this song
  if (mockAudioPlayer) {
    mockAudioPlayer.style.display = 'flex';
  }
  
  // Set total seconds duration matching actual raw data length
  if (song.rawData && song.rawData.engagement) {
    playbackMax = song.rawData.engagement.length - 1;
  } else {
    playbackMax = 30;
  }
  
  // Stop previous playback and reset timeline
  stopMockAudio();
  
  if (currentDashboardView === 'comparative') {
    renderCompareChart();
  } else {
    renderLineChart(song);
  }
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
    plugins: [
      {
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
      },
      {
        id: 'playbackSyncLine',
        afterDraw: (chart) => {
          // Purple vertical line showing interactive music progress sync
          if (playbackTime >= 0) {
            const chartCtx = chart.ctx;
            const xAxis = chart.scales.x;
            const yAxis = chart.scales.y;
            const xPixel = xAxis.getPixelForValue(playbackTime + 1);
            
            if (xPixel !== undefined && xPixel >= xAxis.left && xPixel <= xAxis.right) {
              chartCtx.save();
              chartCtx.beginPath();
              chartCtx.strokeStyle = '#bf5af2'; // Purple progress line
              chartCtx.lineWidth = 2.5;
              if (isPlaying) {
                chartCtx.setLineDash([]); // Solid line when playing
              } else {
                chartCtx.setLineDash([4, 4]); // Dashed when paused
              }
              chartCtx.moveTo(xPixel, yAxis.top);
              chartCtx.lineTo(xPixel, yAxis.bottom);
              chartCtx.stroke();
              
              // Draw marker pointer
              chartCtx.fillStyle = '#bf5af2';
              chartCtx.beginPath();
              chartCtx.arc(xPixel, yAxis.top, 5, 0, 2 * Math.PI);
              chartCtx.fill();
              
              chartCtx.font = 'bold 10px monospace';
              chartCtx.fillStyle = '#bf5af2';
              chartCtx.textAlign = 'center';
              chartCtx.fillText(`Sync ${playbackTime}s`, xPixel, yAxis.top - 8);
              chartCtx.restore();
            }
          }
        }
      }
    ]
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
    // Inject the amazing MBTI profiles first!
    htmlContent += `
      <div style="width: 100%; margin-bottom: 3.5rem; text-align: center; animation: fadeIn 0.6s ease-out;">
        <h2 style="font-size: 1.6rem; color: #fff; margin-bottom: 0.5rem; display: flex; align-items: center; justify-content: center; gap: 8px;">
          🔮 재미로 보는 조원들의 뇌파 MBTI (Neuro-MBTI)
        </h2>
        <p style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 1.8rem;">평균 뇌파 데이터를 분석하여 도출한 AI 기반의 성향 프로필</p>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1.5rem; width: 100%;">
          
          <!-- 문경수 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(10, 132, 255, 0.15); background: rgba(10, 132, 255, 0.02); transition: all 0.3s ease;">
            <img src="/profile1.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #0a84ff; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(10, 132, 255, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">문경수 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">휴먼AI공학전공 | 21학번</span>
            <div style="background: rgba(10, 132, 255, 0.15); color: #64d2ff; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(10, 132, 255, 0.3);">
              🧠 초집중 AI 마스터 (A.I.F.C)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              인공지능 연구원다운 탄탄한 뇌지컬! 강렬하고 빠른 곡에서도 뇌파의 흔들림 없이 극도의 몰입도를 유지하는 '인간 AI'의 면모를 보여줍니다. 몰입 최고조 시 뇌 효율성 극대화!
            </p>
          </div>
          
          <!-- 김한주 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(48, 209, 88, 0.15); background: rgba(48, 209, 88, 0.02); transition: all 0.3s ease;">
            <img src="/profile2.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #30d158; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(48, 209, 88, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">김한주 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">음악학과 작곡과 | 21학번</span>
            <div style="background: rgba(48, 209, 88, 0.15); color: #30d158; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(48, 209, 88, 0.3);">
              🧘‍♂️ 만사태평 감성 음유시인 (Z.E.N.S)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              음악학과 대표 작곡가답게 어떤 복잡한 멜로디가 나와도 스트레스 반응 최소화! 음악 속에서 가장 편안하고 스트레스 지수가 낮아지는 완벽한 이완도 끝판왕, 걸어 다니는 명상 숲 수준입니다.
            </p>
          </div>
          
          <!-- 김용석 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(255, 159, 10, 0.15); background: rgba(255, 159, 10, 0.02); transition: all 0.3s ease;">
            <img src="/profile3.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #ff9f0a; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(255, 159, 10, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">김용석 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">자유전공학부 인문 | 26학번</span>
            <div style="background: rgba(255, 159, 10, 0.15); color: #ff9f0a; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(255, 159, 10, 0.3);">
              🔥 아드레날린 활성 폭발러 (E.X.C.I)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              지적인 인문학도 뒤에 감춰진 폭발적인 내적 댄스 본능! 웅장한 록 멜로디가 시작되면 뇌파 활성도가 성층권을 뚫고 날아갑니다. 열정적이며 창의적인 에너지를 내면에 품고 있습니다.
            </p>
          </div>
          
          <!-- 홍수민 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(191, 90, 242, 0.15); background: rgba(191, 90, 242, 0.02); transition: all 0.3s ease;">
            <img src="/profile4.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #bf5af2; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(191, 90, 242, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">홍수민 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">조형예술학과 | 25학번</span>
            <div style="background: rgba(191, 90, 242, 0.15); color: #bf5af2; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(191, 90, 242, 0.3);">
              👀 예술적 호기심 탐험가 (I.N.T.R)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              시각 예술가다운 고도의 관찰력과 호기심! 신선한 멜로디나 특이한 음역대가 노출될 때마다 흥미도가 수직으로 상승합니다. 사소한 사운드의 디테일까지 기가 막히게 캐치해 냅니다.
            </p>
          </div>
          
        </div>
      </div>
    `;

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

// Celebration Confetti effect when entering awards
function triggerConfetti() {
  const canvas = document.createElement('canvas');
  canvas.style.position = 'fixed';
  canvas.style.top = '0';
  canvas.style.left = '0';
  canvas.style.width = '100vw';
  canvas.style.height = '100vh';
  canvas.style.pointerEvents = 'none';
  canvas.style.zIndex = '9999';
  document.body.appendChild(canvas);
  
  const ctx = canvas.getContext('2d');
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;
  
  const colors = ['#0a84ff', '#ff9f0a', '#30d158', '#ff375f', '#bf5af2', '#ff453a'];
  const confettiCount = 120;
  const particles = [];
  
  for (let i = 0; i < confettiCount; i++) {
    particles.push({
      x: Math.random() * canvas.width,
      y: Math.random() * canvas.height - canvas.height,
      r: Math.random() * 5 + 3,
      d: Math.random() * confettiCount,
      color: colors[Math.floor(Math.random() * colors.length)],
      tilt: Math.random() * 8 - 4,
      tiltAngleIncremental: Math.random() * 0.05 + 0.02,
      tiltAngle: 0
    });
  }
  
  let animationId;
  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    let active = false;
    
    particles.forEach((p) => {
      p.tiltAngle += p.tiltAngleIncremental;
      p.y += (Math.cos(p.d) + 2.5 + p.r / 2) / 2;
      p.x += Math.sin(p.tiltAngle);
      p.tilt = Math.sin(p.tiltAngle - p.r / 2) * 12;
      
      if (p.y <= canvas.height) {
        active = true;
      }
      
      ctx.beginPath();
      ctx.lineWidth = p.r;
      ctx.strokeStyle = p.color;
      ctx.moveTo(p.x + p.tilt + p.r / 2, p.y);
      ctx.lineTo(p.x + p.tilt, p.y + p.tilt + p.r / 2);
      ctx.stroke();
    });
    
    if (active) {
      animationId = requestAnimationFrame(draw);
    } else {
      document.body.removeChild(canvas);
    }
  }
  
  draw();
}

// ==========================================
// 💡 Interactive Mock Audio Player & Sync Logic
// ==========================================
function initMockPlayer() {
  if (!playerPlayBtn || !playerStopBtn || !playerProgressContainer) return;

  playerPlayBtn.addEventListener('click', () => {
    if (isPlaying) {
      pauseMockAudio();
    } else {
      playMockAudio();
    }
  });

  playerStopBtn.addEventListener('click', () => {
    stopMockAudio();
  });

  playerProgressContainer.addEventListener('click', (e) => {
    const rect = playerProgressContainer.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const width = rect.width;
    const percentage = clickX / width;
    
    playbackTime = Math.floor(percentage * playbackMax);
    updatePlayerUI();
    
    // Sync chart lines instantly
    if (lineChartInstance) lineChartInstance.update('none');
    if (compareChartInstance) compareChartInstance.update('none');
  });
}

function playMockAudio() {
  isPlaying = true;
  playerPlayBtn.textContent = '⏸ Pause';
  albumArt.classList.add('playing');
  playerEq.classList.add('active');

  playbackInterval = setInterval(() => {
    playbackTime++;
    if (playbackTime > playbackMax) {
      stopMockAudio();
    } else {
      updatePlayerUI();
      // Tick line charts to redraw the timeline guide line in real-time
      if (lineChartInstance) lineChartInstance.update('none');
      if (compareChartInstance) compareChartInstance.update('none');
    }
  }, 1000);
}

function pauseMockAudio() {
  isPlaying = false;
  playerPlayBtn.textContent = '▶ Play';
  albumArt.classList.remove('playing');
  playerEq.classList.remove('active');
  if (playbackInterval) clearInterval(playbackInterval);
}

function stopMockAudio() {
  isPlaying = false;
  playbackTime = 0;
  playerPlayBtn.textContent = '▶ Play';
  albumArt.classList.remove('playing');
  playerEq.classList.remove('active');
  if (playbackInterval) clearInterval(playbackInterval);
  updatePlayerUI();
  
  if (lineChartInstance) lineChartInstance.update('none');
  if (compareChartInstance) compareChartInstance.update('none');
}

function updatePlayerUI() {
  const percent = (playbackTime / playbackMax) * 100;
  playerProgressBar.style.width = `${percent}%`;
  
  const secStr = String(playbackTime).padStart(2, '0');
  playerTime.textContent = `00:${secStr}`;
}

// ==========================================
// 💡 Multi-user Comparison Toggle & Chart Renderer
// ==========================================
function initDashboardTabs() {
  if (!btnViewPersonal || !btnViewComparative) return;

  btnViewPersonal.addEventListener('click', () => {
    btnViewPersonal.classList.add('active');
    btnViewComparative.classList.remove('active');
    personalRadarPanel.style.display = 'flex';
    comparativePanel.style.display = 'none';
    currentDashboardView = 'personal';
    
    // Clean player on toggle
    stopMockAudio();
  });

  btnViewComparative.addEventListener('click', () => {
    btnViewComparative.classList.add('active');
    btnViewPersonal.classList.remove('active');
    personalRadarPanel.style.display = 'none';
    comparativePanel.style.display = 'flex';
    currentDashboardView = 'comparative';
    
    stopMockAudio();
    renderCompareChart();
  });

  compareMetricSelect.addEventListener('change', () => {
    renderCompareChart();
  });
}

function renderCompareChart() {
  const ctx = document.getElementById('compareChart').getContext('2d');
  if (compareChartInstance) compareChartInstance.destroy();

  const metric = compareMetricSelect.value;
  const currentSongNum = activeSongIndex + 1;
  const meta = SONG_METADATA[activeSongIndex] || { title: `Song ${currentSongNum}` };

  const memberKeys = ['member3', 'member1', 'member2', 'member4'];
  const colors = {
    member3: '#0a84ff', // 문경수 (Blue)
    member1: '#30d158', // 김한주 (Green)
    member2: '#ff9f0a', // 김용석 (Orange)
    member4: '#bf5af2'  // 홍수민 (Purple)
  };

  const datasets = [];
  let labels = [];

  memberKeys.forEach(mKey => {
    const saved = localStorage.getItem(`brainwaveData_${mKey}`);
    if (saved) {
      const parsed = JSON.parse(saved);
      const songData = parsed.find(s => s.songNumber === currentSongNum);
      if (songData && songData.rawData && songData.rawData[metric]) {
        const rawVals = songData.rawData[metric];
        
        // Use the longest length for labels
        if (rawVals.length > labels.length) {
          labels = rawVals.map((_, i) => i + 1);
        }

        datasets.push({
          label: MEMBERS[mKey],
          data: rawVals,
          borderColor: colors[mKey],
          backgroundColor: 'transparent',
          borderWidth: 2.5,
          pointRadius: 0,
          tension: 0.35
        });
      }
    }
  });

  // Adjust maximum playback limit to match this song's raw timeline length
  playbackMax = labels.length > 0 ? labels.length - 1 : 30;

  compareChartInstance = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top',
          labels: { color: '#f8fafc', font: { family: '-apple-system', size: 11, weight: 'bold' } }
        },
        tooltip: { enabled: true }
      },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.03)' }, ticks: { color: '#94a3b8' } },
        y: { min: 0, max: 1, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } }
      }
    },
    plugins: [{
      id: 'playbackSyncLine',
      afterDraw: (chart) => {
        // Red vertical line showing interactive music progress sync
        if (playbackTime >= 0) {
          const chartCtx = chart.ctx;
          const xAxis = chart.scales.x;
          const yAxis = chart.scales.y;
          const xPixel = xAxis.getPixelForValue(playbackTime + 1);
          
          if (xPixel !== undefined && xPixel >= xAxis.left && xPixel <= xAxis.right) {
            chartCtx.save();
            chartCtx.beginPath();
            chartCtx.strokeStyle = '#ff375f'; // Neon Red
            chartCtx.lineWidth = 2;
            if (isPlaying) {
              chartCtx.setLineDash([]); // Solid line when playing
            } else {
              chartCtx.setLineDash([4, 4]); // Dashed when paused
            }
            chartCtx.moveTo(xPixel, yAxis.top);
            chartCtx.lineTo(xPixel, yAxis.bottom);
            chartCtx.stroke();
            
            // Draw marker pointer
            chartCtx.fillStyle = '#ff375f';
            chartCtx.beginPath();
            chartCtx.arc(xPixel, yAxis.top, 5, 0, 2 * Math.PI);
            chartCtx.fill();
            
            chartCtx.font = 'bold 10px monospace';
            chartCtx.fillStyle = '#ff375f';
            chartCtx.textAlign = 'center';
            chartCtx.fillText(`Sync ${playbackTime}s`, xPixel, yAxis.top - 8);
            chartCtx.restore();
          }
        }
      }
    }]
  });
}

// Global State for Song Analysis
let currentSongAnalysisIndex = 0;
let currentSongAnalysisChartType = 'radar';
let songAnalysisChartInstance = null;
let isSongAnalysisInitialized = false;

function loadSongAnalysis() {
  const songCardList = document.getElementById('songCardList');
  if (!songCardList) return;
  
  songCardList.innerHTML = '';
  SONG_METADATA.forEach((song, idx) => {
    const activeClass = idx === currentSongAnalysisIndex ? 'active' : '';
    const div = document.createElement('div');
    div.className = `song-card ${activeClass}`;
    div.innerHTML = `
      <span class="song-card-index">${String(idx + 1).padStart(2, '0')}</span>
      <div class="song-card-cover">${song.cover || '🎵'}</div>
      <div class="song-card-info">
        <div class="song-card-title">${song.title}</div>
        <div class="song-card-artist">${song.artist}</div>
      </div>
    `;
    div.addEventListener('click', () => selectSongForAnalysis(idx));
    songCardList.appendChild(div);
  });

  // Setup control toggles once
  if (!isSongAnalysisInitialized) {
    const btnRadar = document.getElementById('btnSongChartRadar');
    const btnBar = document.getElementById('btnSongChartBar');
    
    if (btnRadar && btnBar) {
      btnRadar.addEventListener('click', () => {
        btnRadar.classList.add('active');
        btnBar.classList.remove('active');
        currentSongAnalysisChartType = 'radar';
        renderSongCompareChart(currentSongAnalysisIndex, 'radar');
      });
      
      btnBar.addEventListener('click', () => {
        btnBar.classList.add('active');
        btnRadar.classList.remove('active');
        currentSongAnalysisChartType = 'bar';
        renderSongCompareChart(currentSongAnalysisIndex, 'bar');
      });
    }
    isSongAnalysisInitialized = true;
  }

  selectSongForAnalysis(currentSongAnalysisIndex);
}

function selectSongForAnalysis(index) {
  currentSongAnalysisIndex = index;
  
  // Highlight active song card
  const cards = document.querySelectorAll('#songCardList .song-card');
  cards.forEach((card, idx) => {
    if (idx === index) {
      card.classList.add('active');
    } else {
      card.classList.remove('active');
    }
  });

  const songMeta = SONG_METADATA[index];
  const titleEl = document.getElementById('songAnalysisTitle');
  const artistEl = document.getElementById('songAnalysisArtist');
  if (titleEl) titleEl.textContent = songMeta.title;
  if (artistEl) artistEl.textContent = songMeta.artist;

  // Render comparative chart
  renderSongCompareChart(index, currentSongAnalysisChartType);

  // Compute AI insights
  renderAIEmotionInsights(index);
}

function renderSongCompareChart(songIndex, chartType) {
  const canvas = document.getElementById('songAnalysisChart');
  if (!canvas) return;

  if (songAnalysisChartInstance) {
    songAnalysisChartInstance.destroy();
  }

  const ctx = canvas.getContext('2d');
  
  // Build datasets for 4 team members
  const memberList = ['member3', 'member1', 'member2', 'member4'];
  const datasets = memberList.map(mId => {
    const savedData = localStorage.getItem(`brainwaveData_${mId}`);
    const memberData = savedData ? JSON.parse(savedData) : [];
    // The data is mapped to target columns averages
    const songData = memberData[songIndex];
    const dataVals = TARGET_COLUMNS.map(col => songData ? songData.averages[col] : 0);

    const colorInfo = {
      member3: { border: '#0a84ff', bgRadar: 'rgba(10, 132, 255, 0.15)', bgBar: 'rgba(10, 132, 255, 0.7)' }, // 문경수 - Blue
      member1: { border: '#30d158', bgRadar: 'rgba(48, 209, 88, 0.15)', bgBar: 'rgba(48, 209, 88, 0.7)' },   // 김한주 - Green
      member2: { border: '#ff9f0a', bgRadar: 'rgba(255, 159, 10, 0.15)', bgBar: 'rgba(255, 159, 10, 0.7)' },  // 김용석 - Orange
      member4: { border: '#bf5af2', bgRadar: 'rgba(191, 90, 242, 0.15)', bgBar: 'rgba(191, 90, 242, 0.7)' }  // 홍수민 - Purple
    }[mId];

    if (chartType === 'radar') {
      return {
        label: MEMBERS[mId],
        data: dataVals,
        borderColor: colorInfo.border,
        backgroundColor: colorInfo.bgRadar,
        borderWidth: 2,
        pointBackgroundColor: colorInfo.border,
        pointBorderColor: '#fff',
        pointHoverBackgroundColor: '#fff',
        pointHoverBorderColor: colorInfo.border
      };
    } else {
      return {
        label: MEMBERS[mId],
        data: dataVals,
        borderColor: colorInfo.border,
        backgroundColor: colorInfo.bgBar,
        borderWidth: 1,
        borderRadius: 4
      };
    }
  });

  const chartOptions = chartType === 'radar' ? {
    responsive: true,
    maintainAspectRatio: false,
    scales: {
      r: {
        min: 0,
        max: 1.0,
        angleLines: { color: 'rgba(255,255,255,0.06)' },
        grid: { color: 'rgba(255,255,255,0.08)' },
        pointLabels: { color: '#f8fafc', font: { family: '-apple-system', size: 12, weight: '600' } },
        ticks: { display: false, stepSize: 0.2 }
      }
    },
    plugins: {
      legend: {
        position: 'top',
        labels: { color: '#f8fafc', font: { family: '-apple-system', size: 11, weight: 'bold' } }
      },
      tooltip: { enabled: true }
    }
  } : {
    responsive: true,
    maintainAspectRatio: false,
    scales: {
      x: { grid: { color: 'rgba(255,255,255,0.03)' }, ticks: { color: '#f8fafc', font: { family: '-apple-system', size: 11, weight: '600' } } },
      y: { min: 0, max: 1.0, grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#94a3b8' } }
    },
    plugins: {
      legend: {
        position: 'top',
        labels: { color: '#f8fafc', font: { family: '-apple-system', size: 11, weight: 'bold' } }
      },
      tooltip: { enabled: true }
    }
  };

  songAnalysisChartInstance = new Chart(ctx, {
    type: chartType,
    data: {
      labels: LABELS_KO,
      datasets: datasets
    },
    options: chartOptions
  });
}

function renderAIEmotionInsights(songIndex) {
  const insightGrid = document.getElementById('insightGrid');
  if (!insightGrid) return;
  
  insightGrid.innerHTML = '';
  
  TARGET_COLUMNS.forEach((col, cIdx) => {
    const colNameKo = LABELS_KO[cIdx];
    let maxVal = -1;
    let winnerId = '';

    ['member3', 'member1', 'member2', 'member4'].forEach(mId => {
      const savedData = localStorage.getItem(`brainwaveData_${mId}`);
      if (savedData) {
        const memberData = JSON.parse(savedData);
        const songData = memberData[songIndex];
        if (songData) {
          const val = songData.averages[col];
          if (val > maxVal) {
            maxVal = val;
            winnerId = mId;
          }
        }
      }
    });

    const winnerName = MEMBERS[winnerId] || 'Unknown';
    const percentage = Math.round(maxVal * 100);

    let badgeClass = 'badge-blue';
    if (winnerId === 'member1') badgeClass = 'badge-green';
    if (winnerId === 'member2') badgeClass = 'badge-orange';
    if (winnerId === 'member3') badgeClass = 'badge-blue';
    if (winnerId === 'member4') badgeClass = 'badge-purple';

    const card = document.createElement('div');
    card.className = 'insight-card';
    card.innerHTML = `
      <div class="insight-metric-name">${colNameKo}</div>
      <div class="insight-value-container">
        <span class="insight-winner-name">${winnerName} 👑</span>
        <span class="insight-badge ${badgeClass}">${percentage}%</span>
      </div>
      <div class="insight-score-bar-container">
        <div class="insight-score-bar" style="width: ${percentage}%; background-color: var(--neon-${badgeClass.split('-')[1]}); box-shadow: 0 0 8px var(--neon-${badgeClass.split('-')[1]});"></div>
      </div>
    `;
    insightGrid.appendChild(card);
  });
}

// Initialize on script load
initMockPlayer();
initDashboardTabs();

// Initially show home screen
switchScreen('homeScreen');
