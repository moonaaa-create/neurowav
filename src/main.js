import './style.css';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import html2pdf from 'html2pdf.js';
import Chart from 'chart.js/auto';
import preloadedData from './preloadedData.json';

// Initialize data instantly
Object.keys(preloadedData).forEach(key => {
  localStorage.setItem(`brainwaveData_${key}`, JSON.stringify(preloadedData[key]));
});

// Global State & DOM Element References
let currentUserId = 'member3'; // Default to 문경수 (AI 집중 마스터)
let activeSongIndex = 0;
let userData = [];
let lineChartInstance = null;
let compareChartInstance = null;
let currentDashboardView = 'personal';
let currentComparisonMode = 'average';
let showRadarValues = true;

// DOM Element References
let userSelect, btnViewPersonal, btnViewComparative, personalRadarPanel, comparativePanel, radarGrid, albumArt, songTitle, songArtist, lineChartTitle, compareMetricSelect, compareSongModeSelect;

// Hardcoded Team Members
const MEMBERS = {
  member1: "김한주",
  member2: "김용석",
  member3: "문경수",
  member4: "홍수민"
};

// Mock Metadata based on user image with high-quality SoundHelix progressive electronic tracks for perfect web audio sync
const SONG_METADATA = [
  { title: "Syren", artist: "Anyma", cover: "🎵", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3" },
  { title: "Playing with Fire", artist: "BLACKPINK", cover: "🔥", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-2.mp3" },
  { title: "Call Me Maybe", artist: "Carly Rae Jepsen", cover: "📱", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3" },
  { title: "Sie ergibt sich nicht", artist: "Chang Eun Ah", cover: "🎭", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-4.mp3" },
  { title: "Lemon Tree", artist: "Fools Garden", cover: "🍋", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3" },
  { title: "Spain", artist: "Jesus Molina", cover: "🇪🇸", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-6.mp3" },
  { title: "Jane Doe", artist: "Kenshi Yonezu", cover: "👤", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-7.mp3" },
  { title: "Peligrosa", artist: "Ojos", cover: "⚠️", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-8.mp3" },
  { title: "Lost Chapter", artist: "Pentakill, Jorn", cover: "📖", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-9.mp3" },
  { title: "Attack on Titan", artist: "Sawano Hiroyuki", cover: "⚔️", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-10.mp3" },
  { title: "6 Moments musicaux, Op. 16 : No. 4 in E", artist: "Sergei Rachmaninoff", cover: "🎹", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-11.mp3" },
  { title: "Shoreditch", artist: "Vard", cover: "🎸", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-12.mp3" },
  { title: "Look at Me!", artist: "XXTENTACION", cover: "💥", url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-13.mp3" }
];

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];
const LABELS_KO = ['몰입도(Eng)', '흥미도(Int)', '활성도(Exc)', '스트레스(Str)', '이완도(Rel)'];
const SHORT_LABELS = ['En', 'In', 'Ex', 'St', 'Re'];

const app = document.querySelector('#app');

app.innerHTML = `
  <!-- [STAGE 1] INTRO SCREEN -->
  <div id="introScreen" class="screen active" style="position: relative; min-height: 80vh; justify-content: center;">
    <div class="floating-bg">
      <div class="floating-emoji" style="left: 10%; top: 20%;">👾</div>
      <div class="floating-emoji" style="left: 30%; top: 60%; font-size: 4rem;">🔮</div>
      <div class="floating-emoji" style="left: 50%; top: 30%;">🎮</div>
      <div class="floating-emoji" style="left: 70%; top: 70%; font-size: 5rem;">🧠</div>
      <div class="floating-emoji" style="left: 85%; top: 25%;">⚔️</div>
    </div>
    
    <div class="hero-section" style="text-align: center; padding: 3rem 1rem; animation: fadeIn 0.8s ease-out; z-index: 10; position: relative;">
      <div style="font-family: monospace; font-size: 1.6rem; color: #ffd700; text-shadow: 0 0 10px rgba(255,215,0,0.5); margin-bottom: 1rem; letter-spacing: 2px;">STAGE 1: THE BEGINNING</div>
      <h1 style="font-size: 5.5rem; margin-bottom: 2rem; background: linear-gradient(to right, #bf5af2, #0a84ff, #30d158); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-shadow: 0 0 35px rgba(10, 132, 255, 0.35); font-weight: 800; letter-spacing: -2px;">NeuroWav Quest</h1>
      <p class="subtitle" style="font-size: 1.7rem !important; max-width: 900px !important; margin: 0 auto 4rem auto !important; line-height: 1.8 !important; color: rgba(235, 235, 245, 0.85);">
        <strong>"서로 다른 4명이 음악을 듣고 검출된 뇌파 데이터를 통해 서로를 알아가는 N-BTI"</strong><br>
        자료시각화 5조의 레트로 게임 컨셉 뇌파 탐험! 4명의 도트 캐릭터가 각 Stage를 클리어하며<br>
        자신조차 몰랐던 뇌파 속 진짜 취향과 최종 뇌비티아이(N-BTI)를 파헤칩니다.
      </p>
      
      <div style="display: flex; flex-direction: column; gap: 3rem; align-items: center;">
        <div>
          <h3 style="color: var(--text-secondary); margin-bottom: 2rem; font-weight: 700; letter-spacing: 1px; font-size: 2rem;">🎮 플레이어 캐릭터 선택 (Player Select)</h3>
          <div style="display: flex; gap: 2rem; flex-wrap: wrap; justify-content: center;">
            <!-- 김용석 -->
            <div id="intro-btn-member2" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(255, 69, 58, 0.05); padding: 2rem; border-radius: 24px; border: 2.5px solid rgba(255, 69, 58, 0.4); width: 210px; transition: all 0.3s; box-shadow: 0 0 25px rgba(255, 69, 58, 0.15);">
              <img src="/profile3.png" style="width: 140px; height: 140px; border-radius: 20px; object-fit: cover; margin-bottom: 1.2rem; border: 3px solid #ff453a;">
              <span style="font-size: 1.4rem; font-weight: 800; color: #fff;">김용석 님</span>
              <span style="font-size: 1.05rem; color: #ff453a; font-weight: 700; margin-top: 0.4rem;">🔥 아드레날린 락커</span>
              <span style="font-size: 0.95rem; color: var(--text-secondary); margin-top: 0.6rem; text-align: center; line-height: 1.4;">자유전공 인문<br>26학번</span>
            </div>
            <!-- 김한주 -->
            <div id="intro-btn-member1" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(10, 132, 255, 0.05); padding: 2rem; border-radius: 24px; border: 2.5px solid rgba(10, 132, 255, 0.4); width: 210px; transition: all 0.3s; box-shadow: 0 0 25px rgba(10, 132, 255, 0.15);">
              <img src="/profile2.png" style="width: 140px; height: 140px; border-radius: 20px; object-fit: cover; margin-bottom: 1.2rem; border: 3px solid #0a84ff;">
              <span style="font-size: 1.4rem; font-weight: 800; color: #fff;">김한주 님</span>
              <span style="font-size: 1.05rem; color: #0a84ff; font-weight: 700; margin-top: 0.4rem;">🧘‍♂️ 감성 음유시인</span>
              <span style="font-size: 0.95rem; color: var(--text-secondary); margin-top: 0.6rem; text-align: center; line-height: 1.4;">음악학과 작곡과<br>21학번</span>
            </div>
            <!-- 문경수 -->
            <div id="intro-btn-member3" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(191, 90, 242, 0.05); padding: 2rem; border-radius: 24px; border: 2.5px solid rgba(191, 90, 242, 0.4); width: 210px; transition: all 0.3s; box-shadow: 0 0 25px rgba(191, 90, 242, 0.15);">
              <img src="/profile1.png" style="width: 140px; height: 140px; border-radius: 20px; object-fit: cover; margin-bottom: 1.2rem; border: 3px solid #bf5af2;">
              <span style="font-size: 1.4rem; font-weight: 800; color: #fff;">문경수 님</span>
              <span style="font-size: 1.05rem; color: #bf5af2; font-weight: 700; margin-top: 0.4rem;">🧠 AI 집중 마스터</span>
              <span style="font-size: 0.95rem; color: var(--text-secondary); margin-top: 0.6rem; text-align: center; line-height: 1.4;">휴먼AI공학전공<br>21학번</span>
            </div>
            <!-- 홍수민 -->
            <div id="intro-btn-member4" class="profile-card" style="cursor: pointer; display: flex; flex-direction: column; align-items: center; background: rgba(48, 209, 88, 0.05); padding: 2rem; border-radius: 24px; border: 2.5px solid rgba(48, 209, 88, 0.4); width: 210px; transition: all 0.3s; box-shadow: 0 0 25px rgba(48, 209, 88, 0.15);">
              <img src="/profile4.png" style="width: 140px; height: 140px; border-radius: 20px; object-fit: cover; margin-bottom: 1.2rem; border: 3px solid #30d158;">
              <span style="font-size: 1.4rem; font-weight: 800; color: #fff;">홍수민 님</span>
              <span style="font-size: 1.05rem; color: #30d158; font-weight: 700; margin-top: 0.4rem;">🎨 호기심 미술가</span>
              <span style="font-size: 0.95rem; color: var(--text-secondary); margin-top: 0.6rem; text-align: center; line-height: 1.4;">조형예술학과<br>25학번</span>
            </div>
          </div>
        </div>
        
        <div style="margin-top: 3rem; display: flex; gap: 1.5rem; flex-wrap: wrap;">
          <button id="intro-start-btn" class="hero-btn" style="background: linear-gradient(135deg, #0a84ff, #bf5af2); font-size: 1.8rem; padding: 1.2rem 4rem; border-radius: 16px; font-weight: 800; box-shadow: 0 0 35px rgba(10, 132, 255, 0.55); border: 2px solid rgba(255,255,255,0.25);">⚔️ QUEST START (탐험 시작)</button>
        </div>
      </div>
    </div>
  </div>

  <!-- [STAGE 2] HOW WE DID SCREEN -->
  <div id="howWeDidScreen" class="screen">
    <header class="team-header" style="text-align: center; margin-bottom: 3rem;">
      <div style="font-family: monospace; font-size: 1.5rem; color: #30d158; margin-bottom: 0.8rem; letter-spacing: 2px; font-weight: 700;">STAGE 2: THE DATA JOURNEY</div>
      <h1 style="color: #30d158; font-size: 4rem; font-weight: 800; margin-bottom: 1rem;">🛠️ How We DID</h1>
      <p class="subtitle" style="font-size: 1.7rem !important; color: var(--text-secondary);">프로젝트 준비물부터 데이터 검출, 정제 및 클리닝 과정까지의 개발 과정</p>
    </header>
    
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 2rem; width: 100%; margin-bottom: 2rem;">
      <!-- Card 1: Equipment -->
      <div class="panel" style="padding: 2.5rem; text-align: left; gap: 1.5rem; border-color: rgba(48, 209, 88, 0.3); border-width: 1.5px;">
        <div style="font-size: 3.5rem;">🎧</div>
        <h3 style="font-size: 1.8rem; font-weight: 800; color: #fff; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem;">Quest 1: 준비물 및 검사 방법</h3>
        <p style="font-size: 1.2rem; color: var(--text-secondary); line-height: 1.7; text-align: justify;">
          실제 EEG 뇌파를 정밀하게 획득하기 위해 **Emotiv 뇌파 기기**를 머리에 장착한 후, 서로 다른 장르의 대표 음악 **13곡**을 선정하여 각 30초 이상 연속 청취하며 뇌파 피드백 데이터를 스트리밍 및 기록하였습니다.
        </p>
      </div>
      <!-- Card 2: Archiving -->
      <div class="panel" style="padding: 2.5rem; text-align: left; gap: 1.5rem; border-color: rgba(10, 132, 255, 0.3); border-width: 1.5px;">
        <div style="font-size: 3.5rem;">🧠</div>
        <h3 style="font-size: 1.8rem; font-weight: 800; color: #fff; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem;">Quest 2: 5대 감정지표 수집</h3>
        <p style="font-size: 1.2rem; color: var(--text-secondary); line-height: 1.7; text-align: justify;">
          EmotivApp을 통해 실시간 검출된 **Engagement(몰입), Excitement(활성), Interest(흥미), Stress(스트레스), Relaxation(이완)**의 5가지 감정 데이터를 시계열 로그 형태인 CSV/XLSX 원본 파일로 안전하게 아카이빙하였습니다.
        </p>
      </div>
      <!-- Card 3: Archiving -->
      <div class="panel" style="padding: 2.5rem; text-align: left; gap: 1.5rem; border-color: rgba(191, 90, 242, 0.3); border-width: 1.5px;">
        <div style="font-size: 3.5rem;">🧹</div>
        <h3 style="font-size: 1.8rem; font-weight: 800; color: #fff; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem;">Quest 3: AI 기반 데이터 클리닝</h3>
        <p style="font-size: 1.2rem; color: var(--text-secondary); line-height: 1.7; text-align: justify;">
          기존 데이터가 수동 검사 환경의 편차로 인해 노이즈가 많았기에, **바이브 코딩 및 AI 데이터 정제 솔루션**을 사용해 결측치 처리, 데이터 스케일링(0 ~ 1 범위 정규화), 헤더 컬럼 표준화 등의 **클리닝 작업**을 사전 완료하였습니다.
        </p>
      </div>
    </div>
    
    <!-- Big flow graphic -->
    <div class="panel" style="padding: 3rem; text-align: center; gap: 2rem; border-color: rgba(255,255,255,0.1);">
      <h3 style="font-size: 1.8rem; font-weight: 800; color: #fff;">⚙️ NeuroWav Data Flow Pipeline</h3>
      <div class="pipeline-container" style="display: flex; justify-content: space-around; align-items: center; flex-wrap: wrap; gap: 2rem; margin-top: 1.5rem;">
        <div style="background: rgba(255,255,255,0.03); border: 1.5px solid rgba(255,255,255,0.1); padding: 1.5rem 2rem; border-radius: 16px; min-width: 200px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
          <span style="font-size: 2.2rem; display: block; margin-bottom: 0.8rem;">🎧</span>
          <strong style="color: #64d2ff; font-size: 1.35rem; font-weight: 700;">Emotiv EPOC</strong>
          <span style="font-size: 1rem; color: var(--text-secondary); display: block; margin-top: 0.4rem;">Raw EEG Stream</span>
        </div>
        <div class="pipeline-arrow" style="color: var(--text-secondary); font-size: 2.5rem; font-weight: bold;">➡️</div>
        <div style="background: rgba(255,255,255,0.03); border: 1.5px solid rgba(255,255,255,0.1); padding: 1.5rem 2rem; border-radius: 16px; min-width: 200px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
          <span style="font-size: 2.2rem; display: block; margin-bottom: 0.8rem;">📊</span>
          <strong style="color: #ff9f0a; font-size: 1.35rem; font-weight: 700;">XLSX / CSV Logs</strong>
          <span style="font-size: 1rem; color: var(--text-secondary); display: block; margin-top: 0.4rem;">5 Emotions (30s)</span>
        </div>
        <div class="pipeline-arrow" style="color: var(--text-secondary); font-size: 2.5rem; font-weight: bold;">➡️</div>
        <div style="background: rgba(255,255,255,0.03); border: 1.5px solid rgba(255,255,255,0.1); padding: 1.5rem 2rem; border-radius: 16px; min-width: 200px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
          <span style="font-size: 2.2rem; display: block; margin-bottom: 0.8rem;">🤖</span>
          <strong style="color: #30d158; font-size: 1.35rem; font-weight: 700;">AI & Node Engine</strong>
          <span style="font-size: 1rem; color: var(--text-secondary); display: block; margin-top: 0.4rem;">Data Cleaning & Scale</span>
        </div>
        <div class="pipeline-arrow" style="color: var(--text-secondary); font-size: 2.5rem; font-weight: bold;">➡️</div>
        <div style="background: rgba(10,132,255,0.15); border: 2px solid var(--accent-color); padding: 1.5rem 2rem; border-radius: 16px; min-width: 200px; box-shadow: 0 0 25px rgba(10,132,255,0.35);">
          <span style="font-size: 2.2rem; display: block; margin-bottom: 0.8rem;">🔮</span>
          <strong style="color: #fff; font-size: 1.35rem; font-weight: 700;">NeuroWav SPA</strong>
          <span style="font-size: 1rem; color: #ffd700; display: block; margin-top: 0.4rem; font-weight: 600;">Real-time Dash uploader</span>
        </div>
      </div>
    </div>
  </div>

  <!-- [STAGE 3] SEE MUSIC SCREEN -->
  <div id="seeMusicScreen" class="screen">
    <header class="team-header">
      <div style="font-family: monospace; font-size: 0.9rem; color: #bf5af2; margin-bottom: 0.5rem; letter-spacing: 1px;">STAGE 3: TRACK ANALYSIS</div>
      <h1 style="color: #bf5af2;">🎵 Let's See the Music</h1>
      <p class="subtitle">13개 트랙별 장르 분석 해설 및 4인 조원 평균 감정 반응 랭킹</p>
    </header>
    
    <div class="song-analysis-layout">
      <!-- Left Panel: 13 Songs Selection Card List -->
      <div class="panel song-list-sidebar-panel" style="display: flex; flex-direction: column; gap: 1rem; padding: 1.2rem; height: 100%; min-height: 980px;">
        <h3 style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.2rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem; display: flex; align-items: center; gap: 0.5rem;">
          💿 Tracks <span style="font-size: 0.8rem; font-weight: 400; color: var(--text-secondary);">(13 total)</span>
        </h3>
        <div class="song-card-list" id="songCardList" style="max-height: 900px; flex-grow: 1; overflow-y: auto;">
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
          
          <div style="display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;">
            <label class="neon-switch-container" style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer; user-select: none;">
              <input type="checkbox" id="chkShowTeamAverage" checked>
              <span class="neon-switch-slider"></span>
              <span style="font-size: 0.85rem; font-weight: 600; color: #ffd700; text-shadow: 0 0 5px rgba(255, 215, 0, 0.4); display: flex; align-items: center; gap: 0.3rem;">
                👥 Show Team Average
              </span>
            </label>

            <div class="mini-toggle-container">
              <button id="btnSongChartRadar" class="mini-toggle-btn active">🕸️ Radar Overlap</button>
              <button id="btnSongChartBar" class="mini-toggle-btn">📊 Grouped Bar</button>
            </div>
            
            <button id="btnExportCurrentSongRaw" class="neo-select" style="background: rgba(48, 209, 88, 0.15); border: 1px solid rgba(48, 209, 88, 0.4); color: #30d158; border-radius: 8px; padding: 0.4rem 0.8rem; font-size: 0.82rem; font-weight: 700; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.2s;">📥 이 곡 원본 엑셀</button>
          </div>
        </div>

        <!-- Composer Genre Note Card (Special Addition) -->
        <div class="panel" style="padding: 1.5rem; background: rgba(191, 90, 242, 0.05); border-color: rgba(191, 90, 242, 0.2); gap: 0.5rem;">
          <h4 style="font-size: 0.95rem; color: #bf5af2; font-weight: 700; display: flex; align-items: center; gap: 6px;">
            🎼 작곡 전공 김한주의 장르 해설 노트
          </h4>
          <p id="genreComposerNote" style="font-size: 0.88rem; color: #d1c4e9; line-height: 1.6; font-style: italic;">
            -
          </p>
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

    <!-- 13-Track Team Averages Ranking Panel -->
    <div class="panel team-average-compare-panel" style="padding: 1.5rem; display: flex; flex-direction: column; gap: 1rem; width: 100%; margin-top: 2rem;">
      <div class="compare-header-bar" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.75rem;">
        <h3 style="font-size: 1.1rem; font-weight: 600; display: flex; align-items: center; gap: 0.5rem; margin: 0;">
          📊 13-Track Team Averages & Leaderboard
        </h3>
        <div style="display: flex; align-items: center; gap: 1.2rem; flex-wrap: wrap;">
          <button id="btnExportAllAverages" class="neo-select" style="background: rgba(255, 215, 0, 0.15); border: 1px solid rgba(255, 215, 0, 0.4); color: #ffd700; border-radius: 8px; padding: 0.4rem 0.8rem; font-size: 0.82rem; font-weight: 700; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.2s;">📊 13곡 평균 엑셀 다운</button>
          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <span style="font-size: 0.8rem; color: var(--text-secondary); font-weight: 500;">분석 기준:</span>
            <select id="compareSongModeSelect" class="neo-select" style="background: rgba(18, 18, 24, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: #ffd700; padding: 0.4rem 0.8rem; font-size: 0.85rem; outline: none; cursor: pointer; box-shadow: 0 0 10px rgba(0,0,0,0.5); font-weight: 700;">
              <option value="average">📊 팀 평균 감정 수치</option>
              <option value="peak">⚡ 크리티컬 최고 자극치</option>
              <option value="volatility">🎢 감정 롤러코스터 변동폭</option>
            </select>
          </div>
          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <span style="font-size: 0.8rem; color: var(--text-secondary); font-weight: 500;">대상 감정:</span>
            <select id="compareSongMetricSelect" class="neo-select" style="background: rgba(18, 18, 24, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: white; padding: 0.4rem 0.8rem; font-size: 0.85rem; outline: none; cursor: pointer; box-shadow: 0 0 10px rgba(0,0,0,0.5);">
              <option value="engagement">몰입도 (Engagement)</option>
              <option value="interest">흥미도 (Interest)</option>
              <option value="excitement">활성도 (Excitement)</option>
              <option value="stress">스트레스 (Stress)</option>
              <option value="relaxation">이완도 (Relaxation)</option>
            </select>
          </div>
        </div>
      </div>
      <div class="ranking-flex-container" style="display: grid; grid-template-columns: 1.2fr 1fr; gap: 1.5rem; width: 100%; align-items: start; flex-wrap: wrap;">
        <div style="position: relative; height: 420px; width: 100%; min-width: 300px;">
          <canvas id="songTeamAverageChart"></canvas>
        </div>
        <div style="background: rgba(0,0,0,0.25); border: 1px solid rgba(255,255,255,0.06); border-radius: 12px; padding: 1.2rem; display: flex; flex-direction: column; gap: 0.6rem; max-height: 420px; overflow-y: auto; min-width: 300px;">
          <h4 style="font-size: 0.95rem; font-weight: 700; color: #ffd700; display: flex; align-items: center; gap: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 0.5rem; margin-bottom: 0.25rem;">
            🏆 뇌파 감정 반응 전체 순위 (1위 ~ 13위)
          </h4>
          <div id="songTeamAverageLeaderboard" style="display: flex; flex-direction: column; gap: 0.4rem;">
            <!-- Leaderboard rows injected here via JS -->
          </div>
        </div>
      </div>
    </div>
  </div>
  </div>

  <!-- [STAGE 5] SEE PERSON SCREEN (HIGH INTERACTIONS) -->
  <div id="dashboardScreen" class="screen">
    <header class="team-header" style="margin-bottom: 0.5rem;">
      <div style="font-family: monospace; font-size: 0.9rem; color: #0a84ff; margin-bottom: 0.5rem; letter-spacing: 1px;">STAGE 5: THE NEURAL PROFILE</div>
      <h1 style="color: #0a84ff;">👤 Let's See a Person</h1>
      <p class="subtitle">런치패드 기반 동적 로딩을 통한 조원 정밀 분석 및 뇌파 피드백 동기화</p>
    </header>
    
    <!-- 🎛️ MIDI Launchpad control board (4x5 Grid) -->
    <div class="panel" style="padding: 1.5rem; background: rgba(10, 132, 255, 0.03); border: 2px solid rgba(10, 132, 255, 0.2); border-radius: 20px;">
      <h3 style="font-size: 1.1rem; font-weight: 700; color: #ffd700; margin-bottom: 1rem; display: flex; align-items: center; gap: 8px;">
        🎛️ 뇌파 인터랙티브 런치패드 (4 x 5 Neural Launchpad)
        <span style="font-size: 0.75rem; color: var(--text-secondary); font-weight: normal;">(격자 패드를 눌러 조원과 감정을 즉시 믹싱/시각화해 보세요!)</span>
      </h3>
      <div class="launchpad-grid" id="neuralLaunchpad" style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 0.75rem; max-width: 900px; margin: 0 auto; width: 100%;">
        <!-- Interactive 4x5 launchpad buttons generated via JS -->
      </div>
    </div>

    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem; border-bottom: 0.5px solid var(--border-color); padding-bottom: 1rem; flex-wrap: wrap; gap: 1rem; margin-top: 1rem;">
      <header class="dash-header" style="display: flex; align-items: center; gap: 1rem; border-bottom: none; padding-bottom: 0; margin-bottom: 0;">
        <img id="headerProfileImg" src="/profile3.png" style="width: 54px; height: 54px; border-radius: 50%; border: 2.5px solid var(--accent-color); object-fit: cover; box-shadow: 0 0 15px rgba(10, 132, 255, 0.3);">
        <div class="user-info">
          <select id="userSelect" style="font-size: 1.6rem; font-weight: 700; border: none; background: transparent; color: var(--accent-color); cursor: pointer; outline: none; margin-bottom: 0.15rem;">
            <option value="member2">김용석's Profile</option>
            <option value="member1">김한주's Profile</option>
            <option value="member3">문경수's Profile</option>
            <option value="member4">홍수민's Profile</option>
          </select>
          <div id="headerMemberInfo" style="font-size: 0.85rem; color: var(--text-secondary); padding-left: 0.25rem;">자유전공학부 인문 | 26학번</div>
        </div>
      </header>
      
      <div class="toggle-switch-container" style="display: flex; gap: 0.5rem; align-items: center;">
        <button id="btnToggleRadarValues" class="toggle-btn" style="background: rgba(255, 215, 0, 0.15); border: 1px solid rgba(255, 215, 0, 0.4); color: #ffd700;">🏷️ 수치 ON</button>
        <button id="btnViewPersonal" class="toggle-btn active">👤 개인 분석</button>
        <button id="btnViewComparative" class="toggle-btn">👥 조원 비교</button>
      </div>
    </div>

    <!-- Interactive Track Selector Panel -->
    <div style="display: flex; align-items: center; gap: 0.75rem; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05); border-radius: var(--radius-md); padding: 0.8rem 1.2rem; flex-wrap: wrap;">
      <span style="font-size: 0.85rem; font-weight: 600; color: #ffd700;">🎵 현재 오디오 탐색 곡 선택:</span>
      <select id="dashboardSongSelect" class="neo-select" style="background: rgba(18, 18, 24, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: white; padding: 0.4rem 0.8rem; font-size: 0.85rem; outline: none; cursor: pointer; flex-grow: 1;">
        <!-- Filled by JS -->
      </select>
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
              </div>
            </div>
          </div>
          <!-- Radar items injected here -->
        </div>
      </div>

      <!-- Left Panel (Comparative Mode): Multi-user Line Chart -->
      <div class="panel comparison-panel" id="comparativePanel" style="display: none; grid-column: 1 / 2; grid-row: 1 / 2;">
        <div class="comparison-selector-panel" style="display: flex; gap: 1rem; align-items: center; flex-wrap: wrap;">
          <span style="font-weight: 600; font-size: 0.95rem;">👥 조원 전체 감정 비교</span>
          <select id="compareMetricSelect">
            <option value="engagement">🧠 몰입도 (Engagement)</option>
            <option value="interest">👀 흥미도 (Interest)</option>
            <option value="excitement">🔥 활성도 (Excitement)</option>
            <option value="stress">🤯 스트레스 (Stress)</option>
            <option value="relaxation">🧘‍♂️ 이완도 (Relaxation)</option>
          </select>
          <select id="comparePanelSongSelect" style="background: rgba(18, 18, 24, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: white; padding: 0.4rem 0.8rem; font-size: 0.85rem; outline: none; cursor: pointer;">
          </select>
          <button id="btnExportComparativeRaw" class="neo-select" style="background: rgba(10, 132, 255, 0.15); border: 1px solid rgba(10, 132, 255, 0.4); color: #0a84ff; border-radius: 8px; padding: 0.4rem 0.8rem; font-size: 0.82rem; font-weight: 700; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.2s;">📥 비교 데이터 엑셀</button>
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

  <!-- [STAGE 4] N-BTI RESULT SCREEN -->
  <div id="nbtiScreen" class="screen">
    <header class="team-header">
      <div style="font-family: monospace; font-size: 0.9rem; color: #ffd700; margin-bottom: 0.5rem; letter-spacing: 1px;">STAGE 4: PERSONALITY DIAGNOSIS</div>
      <h1 style="color: #ffd700;">🔮 Let's Diagnose N-BTI !</h1>
      <p class="subtitle">실시간 뇌파 비율 척도를 통한 개인별 최종 뇌비티아이 진단</p>
    </header>
    
    <div style="display: flex; gap: 1rem; justify-content: center; margin-bottom: 1.5rem; flex-wrap: wrap;">
      <button id="nbti-tab-member2" class="hero-btn active" style="border-color: #ff453a; background: rgba(255,69,58,0.15);">김용석 님</button>
      <button id="nbti-tab-member1" class="hero-btn" style="border-color: #0a84ff;">김한주 님</button>
      <button id="nbti-tab-member3" class="hero-btn" style="border-color: #bf5af2;">문경수 님</button>
      <button id="nbti-tab-member4" class="hero-btn" style="border-color: #30d158;">홍수민 님</button>
    </div>

    <div class="nbti-container-block" id="nbtiReportBlock" style="width: 100%; min-height: 500px;">
      <!-- Dynamic N-BTI Card & MBTI Bar Chart generated here -->
    </div>
  </div>

  <!-- [STAGE 6] AWARDS / ACHIEVEMENTS SCREEN -->
  <div id="teamScreen" class="screen">
    <header class="team-header">
      <div style="font-family: monospace; font-size: 0.9rem; color: #ff375f; margin-bottom: 0.5rem; letter-spacing: 1px;">STAGE 6: HALL OF FAME</div>
      <h1 id="awardsHeaderTitle" style="color: #ff375f;">🏆 NeuroWav Quest Achievements</h1>
      <p id="awardsHeaderSubtitle" class="subtitle">조원별 뇌파 특성과 결합된 고유 전설적 업적 및 획득 무기 시각화</p>
    </header>
    
    <div style="display: flex; gap: 1rem; justify-content: center; margin-bottom: 2rem; flex-wrap: wrap;">
      <button id="ach-tab-member2" class="hero-btn active" style="border-color: #ff453a; background: rgba(255, 69, 58, 0.15);">김용석's Achievements</button>
      <button id="ach-tab-member1" class="hero-btn" style="border-color: #0a84ff;">김한주's Achievements</button>
      <button id="ach-tab-member3" class="hero-btn" style="border-color: #bf5af2;">문경수's Achievements</button>
      <button id="ach-tab-member4" class="hero-btn" style="border-color: #30d158;">홍수민's Achievements</button>
    </div>

    <div class="team-awards-container" id="teamGrid" style="display: flex; flex-direction: column; gap: 2rem;">
      <!-- Custom Game achievements layout injected here -->
    </div>
  </div>

  <!-- DATA UPLOADER SCREEN -->
  <div id="uploaderScreen" class="screen">
    <header class="team-header">
      <h1 style="color: var(--accent-color);">📤 Brainwave Data Uploader</h1>
      <p class="subtitle">Excel (.xlsx) 또는 CSV 파일을 업로드하여 조원의 뇌파 데이터를 실시간으로 업데이트하세요.</p>
    </header>
    
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 2rem; max-width: 900px; margin: 0 auto; width: 100%;">
      <!-- Left Panel: Uploader Form -->
      <div class="panel" style="padding: 2rem; display: flex; flex-direction: column; gap: 1.5rem;">
        <h3 style="font-size: 1.2rem; font-weight: 600; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px;">
          👤 대상 멤버 선택
        </h3>
        
        <div style="display: flex; flex-direction: column; gap: 0.5rem;">
          <label style="font-size: 0.85rem; color: var(--text-secondary); font-weight: 500;">데이터를 업데이트할 조원:</label>
          <select id="uploadMemberSelect" class="neo-select" style="background: rgba(18, 18, 24, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; color: white; padding: 0.6rem 1rem; font-size: 1rem; outline: none; cursor: pointer; width: 100%; box-shadow: 0 0 10px rgba(0,0,0,0.5);">
            <option value="member2">김용석 (자유전공학부 인문)</option>
            <option value="member1">김한주 (음악학과 작곡과)</option>
            <option value="member3">문경수 (휴먼AI공학전공)</option>
            <option value="member4">홍수민 (조형예술학과)</option>
          </select>
        </div>
        
        <div style="width: 100%; height: 1px; background: rgba(255,255,255,0.08); margin: 0.5rem 0;"></div>
        
        <h3 style="font-size: 1.2rem; font-weight: 600; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px;">
          📂 파일 드롭 또는 선택
        </h3>
        
        <div class="upload-box" id="dropZone" style="border: 2px dashed rgba(255,255,255,0.2); border-radius: 16px; padding: 3rem 1.5rem; text-align: center; cursor: pointer; transition: all 0.3s; background: rgba(255,255,255,0.01);">
          <div class="upload-icon" style="font-size: 3rem; margin-bottom: 1rem;">📁</div>
          <div class="upload-text" style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.5rem;">Drag & drop files here</div>
          <div class="upload-subtext" style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 1.5rem;">or click to select files (.xlsx, .csv)</div>
          <input type="file" id="fileInput" multiple accept=".xlsx,.xls,.csv" style="display: none;">
          <span style="font-size: 0.75rem; color: #ff9f0a; background: rgba(255, 159, 10, 0.1); padding: 0.3rem 0.8rem; border-radius: 20px;">노래 번호 자동 매칭 (예: 1.xlsx, 2_Syren.csv 등)</span>
        </div>
      </div>
      
      <!-- Right Panel: Status Log -->
      <div class="panel" style="padding: 2rem; display: flex; flex-direction: column; gap: 1rem; max-height: 450px;">
        <h3 style="font-size: 1.2rem; font-weight: 600; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px; justify-content: space-between;">
          <span>📈 업로드 및 파싱 로그</span>
          <button id="clearLogBtn" style="font-size: 0.75rem; padding: 0.3rem 0.6rem; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 4px; color: var(--text-secondary);">Clear</button>
        </h3>
        
        <div id="uploadLog" style="flex-grow: 1; background: rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; padding: 1rem; font-family: monospace; font-size: 0.8rem; overflow-y: auto; color: #94a3b8; display: flex; flex-direction: column; gap: 0.5rem; height: 300px;">
          <div style="color: #64d2ff;">&gt; 업로더가 준비되었습니다. 파일을 드롭해 주세요.</div>
        </div>
      </div>
  </div>
`;

// Initialize DOM References
userSelect = document.getElementById('userSelect');
btnViewPersonal = document.getElementById('btnViewPersonal');
btnViewComparative = document.getElementById('btnViewComparative');
personalRadarPanel = document.getElementById('personalRadarPanel');
comparativePanel = document.getElementById('comparativePanel');
radarGrid = document.getElementById('radarGrid');
albumArt = document.getElementById('albumArt');
songTitle = document.getElementById('songTitle');
songArtist = document.getElementById('songArtist');
lineChartTitle = document.getElementById('lineChartTitle');
compareMetricSelect = document.getElementById('compareMetricSelect');
compareSongModeSelect = document.getElementById('compareSongModeSelect');

if (userSelect) {
  userSelect.addEventListener('change', (e) => {
    currentUserId = e.target.value;
    loadDashboard();
  });
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

  // Sync the userSelect dropdown value to currentUserId
  const userSelectEl = document.getElementById('userSelect');
  if (userSelectEl) {
    userSelectEl.value = currentUserId;
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
    
    const radarPlugins = [];
    if (showRadarValues) {
      radarPlugins.push({
        id: 'showValues',
        afterDatasetsDraw: (chart) => {
          const { ctx, scales: { r } } = chart;
          ctx.save();
          ctx.font = 'bold 9px sans-serif';
          ctx.fillStyle = '#ffffff';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';

          chart.data.datasets.forEach((dataset, datasetIndex) => {
            const meta = chart.getDatasetMeta(datasetIndex);
            meta.data.forEach((point, index) => {
              const val = dataset.data[index];
              if (val === undefined || val === null) return;
              
              const center = { x: r.xCenter, y: r.yCenter };
              const dx = point.x - center.x;
              const dy = point.y - center.y;
              const dist = Math.sqrt(dx * dx + dy * dy);
              
              let xText = point.x;
              let yText = point.y;
              if (dist > 0) {
                const offset = 12; // offset in pixels
                xText += (dx / dist) * offset;
                yText += (dy / dist) * offset;
              }
              
              // Draw a small background for readability
              ctx.fillStyle = 'rgba(0, 0, 0, 0.6)';
              const textWidth = ctx.measureText(val.toFixed(2)).width;
              ctx.fillRect(xText - textWidth/2 - 2, yText - 6, textWidth + 4, 12);
              
              // Draw the text
              ctx.fillStyle = '#ffd700'; // gold color
              ctx.fillText(val.toFixed(2), xText, yText);
            });
          });
          ctx.restore();
        }
      });
    }

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
      },
      plugins: radarPlugins
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
  
  const dsSelect = document.getElementById('dashboardSongSelect');
  if (dsSelect) dsSelect.value = index;
  
  const cpSelect = document.getElementById('comparePanelSongSelect');
  if (cpSelect) cpSelect.value = index;
  
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
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(191, 90, 242, 0.15); background: rgba(191, 90, 242, 0.02); transition: all 0.3s ease;">
            <img src="/profile1.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #bf5af2; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(191, 90, 242, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">문경수 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">휴먼AI공학전공 | 21학번</span>
            <div style="background: rgba(191, 90, 242, 0.15); color: #bf5af2; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(191, 90, 242, 0.3);">
              🧠 도파민 광전사 (A.I.F.C)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              전체 평균 흥미도(0.530)와 활성도(0.422) 압도적 팀 내 1위! 음악을 들을 때 뇌가 고도의 인지 활성과 아드레날린 분출 반응을 보이는 정열적이고 폭발적인 리액션의 대명사입니다.
            </p>
          </div>
          
          <!-- 김한주 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(10, 132, 255, 0.15); background: rgba(10, 132, 255, 0.02); transition: all 0.3s ease;">
            <img src="/profile2.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #0a84ff; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(10, 132, 255, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">김한주 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">음악학과 작곡과 | 21학번</span>
            <div style="background: rgba(10, 132, 255, 0.15); color: #0a84ff; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(10, 132, 255, 0.3);">
              🧘‍♂️ 개념 치유사 (Z.E.N.S)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              전체 평균 스트레스 최저(0.355) 및 이완도 최고(0.426)! 음악 감상 시 뇌가 가장 극단적인 평온과 안정적 심신 상태를 단단히 유도해 내며, 선율 속에서 자극 없이 명상적 치유를 이룹니다.
            </p>
          </div>
          
          <!-- 김용석 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(255, 69, 58, 0.15); background: rgba(255, 69, 58, 0.02); transition: all 0.3s ease;">
            <img src="/profile3.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #ff453a; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(255, 69, 58, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">김용석 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">자유전공학부 인문 | 26학번</span>
            <div style="background: rgba(255, 69, 58, 0.15); color: #ff453a; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(255, 69, 58, 0.3);">
              🔥 예민한 분석가 (E.X.C.I)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              전체 평균 스트레스 최고(0.404)이면서도 큰 곡별 반응 편차! 매 멜로디의 사소한 신호와 악기 배음 하나까지도 뇌가 기민하고 섬세하게 감지하여 다방면으로 지각하는 예민한 탐색형입니다.
            </p>
          </div>
          
          <!-- 홍수민 -->
          <div class="king-card" style="padding: 1.5rem; display: flex; flex-direction: column; align-items: center; text-align: center; border: 1px solid rgba(48, 209, 88, 0.15); background: rgba(48, 209, 88, 0.02); transition: all 0.3s ease;">
            <img src="/profile4.png" style="width: 72px; height: 72px; border-radius: 50%; object-fit: cover; border: 2.5px solid #30d158; margin-bottom: 0.8rem; box-shadow: 0 0 15px rgba(48, 209, 88, 0.3);">
            <h3 style="font-size: 1.1rem; color: #fff; margin-bottom: 0.2rem;">홍수민 님</h3>
            <span style="font-size: 0.72rem; color: var(--text-secondary); margin-bottom: 0.6rem;">조형예술학과 | 25학번</span>
            <div style="background: rgba(48, 209, 88, 0.15); color: #30d158; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.8rem; font-weight: 700; margin-bottom: 0.8rem; border: 1px solid rgba(48, 209, 88, 0.3);">
              👀 강철의 포커페이스 (I.N.T.R)
            </div>
            <p style="font-size: 0.8rem; color: #b0c4de; line-height: 1.5; text-align: justify; text-justify: inter-character; margin: 0;">
              전체 평균 활성도 극저치(0.222) 수호자! 웬만한 사운드 자극이나 소음 변화에도 뇌가 결코 동요하지 않으며, 단단한 멘탈 통제력을 통해 명상에 수렴하는 부동의 포커페이스를 보여줍니다.
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

// 💡 Multi-user Comparison Toggle & Chart Renderer
// ==========================================
function initDashboardTabs() {
  if (!btnViewPersonal || !btnViewComparative) return;

  const btnToggleRadarValues = document.getElementById('btnToggleRadarValues');
  if (btnToggleRadarValues) {
    btnToggleRadarValues.addEventListener('click', () => {
      showRadarValues = !showRadarValues;
      btnToggleRadarValues.textContent = showRadarValues ? '🏷️ 수치 ON' : '🏷️ 수치 OFF';
      if (showRadarValues) {
        btnToggleRadarValues.style.background = 'rgba(255, 215, 0, 0.15)';
        btnToggleRadarValues.style.borderColor = 'rgba(255, 215, 0, 0.4)';
        btnToggleRadarValues.style.color = '#ffd700';
      } else {
        btnToggleRadarValues.style.background = 'rgba(255, 255, 255, 0.05)';
        btnToggleRadarValues.style.borderColor = 'rgba(255, 255, 255, 0.1)';
        btnToggleRadarValues.style.color = 'var(--text-secondary)';
      }
      renderRadarGrid();
    });
  }

  btnViewPersonal.addEventListener('click', () => {
    btnViewPersonal.classList.add('active');
    btnViewComparative.classList.remove('active');
    personalRadarPanel.style.display = 'flex';
    comparativePanel.style.display = 'none';
    currentDashboardView = 'personal';
    

  });

  btnViewComparative.addEventListener('click', () => {
    btnViewComparative.classList.add('active');
    btnViewPersonal.classList.remove('active');
    personalRadarPanel.style.display = 'none';
    comparativePanel.style.display = 'flex';
    currentDashboardView = 'comparative';
    

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

  const memberKeys = ['member2', 'member1', 'member3', 'member4'];
  const colors = {
    member3: '#bf5af2', // 문경수 (Purple)
    member1: '#0a84ff', // 김한주 (Blue)
    member2: '#ff453a', // 김용석 (Red)
    member4: '#30d158'  // 홍수민 (Green)
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
    plugins: []
  });
}

// Global State for Song Analysis
let currentSongAnalysisIndex = 0;
let currentSongAnalysisChartType = 'radar';
let songAnalysisChartInstance = null;
let songTeamAverageChartInstance = null;
let showTeamAverage = true;
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
    const chkTeamAvg = document.getElementById('chkShowTeamAverage');
    const selectMetric = document.getElementById('compareSongMetricSelect');
    
    if (btnRadar && btnBar) {
      btnRadar.addEventListener('click', () => {
        btnRadar.classList.add('active');
        btnBar.classList.remove('active');
        currentSongAnalysisChartType = 'radar';
        renderSongCompareChart(currentSongAnalysisIndex, currentSongAnalysisChartType);
      });
      
      btnBar.addEventListener('click', () => {
        btnBar.classList.add('active');
        btnRadar.classList.remove('active');
        currentSongAnalysisChartType = 'bar';
        renderSongCompareChart(currentSongAnalysisIndex, currentSongAnalysisChartType);
      });
    }

    if (chkTeamAvg) {
      // Set the initial checkbox state from our global variable
      chkTeamAvg.checked = showTeamAverage;
      chkTeamAvg.addEventListener('change', (e) => {
        showTeamAverage = e.target.checked;
        renderSongCompareChart(currentSongAnalysisIndex, currentSongAnalysisChartType);
      });
    }

    if (selectMetric) {
      selectMetric.addEventListener('change', (e) => {
        render13TracksTeamAverageChart(e.target.value);
      });
    }

    const selectMode = document.getElementById('compareSongModeSelect');
    if (selectMode) {
      selectMode.addEventListener('change', (e) => {
        currentComparisonMode = e.target.value;
        const currentMetric = selectMetric ? selectMetric.value : 'engagement';
        render13TracksTeamAverageChart(currentMetric);
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

  const genreComposerNote = document.getElementById('genreComposerNote');
  if (genreComposerNote) {
    const COMPOSER_NOTES = {
      0: "멜로딕 테크노. 반복적이고 몽환적인 신스 베이스가 뇌를 서서히 자극하여 조원 전체의 몰입도(Engagement)를 고조시킵니다. (Anyma - Syren 🎵)",
      1: "K-POP 댄스. 강렬한 비트와 친숙한 훅송 멜로디가 활성도(Excitement)와 흥미도(Interest)를 즉각적으로 끌어올리는 효과가 있습니다. (BLACKPINK - 불장난 🔥)",
      2: "버블검 팝. 밝고 신나는 리듬감 덕분에 이완도(Relaxation)와 흥미도가 높게 유지되며, 스트레스가 급격히 하락합니다. (Carly Rae Jepsen - Call Me Maybe 📱)",
      3: "뮤지컬 락. 드라마틱하고 웅장한 가창이 활성도를 크게 흔들어 놓으며 감정적 기복(스트레스/흥미 혼재)을 자극합니다. (장은아 - Sie ergibt sich nicht 🎭)",
      4: "어쿠스틱 인디 팝. 경쾌하고 따뜻한 멜로디라인이 뇌파를 극도로 편안하게 만들어 김한주 님을 비롯한 조원들의 스트레스를 지우고 이완도를 최고치로 올립니다. (이완도 폭발 명약 Lemon Tree 🍋)",
      5: "라틴 재즈 피아노. 복잡하고 빠른 재즈 스케일과 리듬 변화가 일어날 때 뇌파의 흥미도와 활성도가 어지럽게 요동치는 패턴을 보여줍니다. (Jesus Molina - Spain 🇪🇸)",
      6: "J-POP. J-POP 특유의 속도감 있는 세션 구성과 독보적인 리듬감이 뇌의 흥미도와 고도의 몰입을 동시에 잡습니다. (Kenshi Yonezu - Jane Doe 👤)",
      7: "프렌치 인디 팝. 이국적이고 시크한 톤앤매너로, 고도의 예술적 자극을 유도하여 미술 전공 홍수민 님의 예술적 호기심을 극대화합니다. (Ojos - Peligrosa ⚠️)",
      8: "헤비 메탈. 거친 메탈 기타 리프가 뇌를 각성시켜 김용석 님의 아드레날린(Excitement)을 최대치로 분출하지만, 동시에 스트레스 반응도 동반 상승하는 패턴을 보입니다. (Pentakill - Lost Chapter 📖)",
      9: "오케스트라 에픽. 영화 같은 웅장한 사운드 트랙이 조원들의 몰입(Engagement)을 끌어올리며 성전의 한복판에 선 듯한 긴장감을 자극합니다. (진격의 거인 OST ⚔️)",
      10: "정통 클래식 피아노. 한순간도 쉴 틈 없는 건반의 속주가 클래식 선율 특유의 우아함과 현란함으로 작곡과 김한주 님의 뇌파를 완전히 동화시켜 클래식 거장다운 깊은 교감을 이끌어냅니다. (Rachmaninoff - Moment Musicaux 🎹)",
      11: "인디 기타 인스트루멘탈. 빈티지한 기타 사운드가 조용히 마음을 어루만져 조원 전체적으로 높은 수준의 이완도와 스트레스 해소 작용을 합니다. (Vard - Shoreditch 🎸)",
      12: "디스토션 트랩/힙합. 의도적으로 깨진 거친 베이스 사운드가 조원들의 스트레스 반응을 즉시 폭발시키며, 뇌파를 가장 강렬하게 흔들어놓는 파괴적인 충격을 줍니다. (XXTENTACION - Look at Me! 💥)"
    };
    genreComposerNote.innerHTML = COMPOSER_NOTES[index] || "장르 해설이 준비되지 않았습니다.";
  }

  // Render comparative chart
  renderSongCompareChart(index, currentSongAnalysisChartType);

  // Compute AI insights
  renderAIEmotionInsights(index);

  // Render 13-track Team Averages ranking chart based on current metric selection
  const selectMetric = document.getElementById('compareSongMetricSelect');
  const currentMetric = selectMetric ? selectMetric.value : 'engagement';
  render13TracksTeamAverageChart(currentMetric);
}

function renderSongCompareChart(songIndex, chartType) {
  const canvas = document.getElementById('songAnalysisChart');
  if (!canvas) return;

  if (songAnalysisChartInstance) {
    songAnalysisChartInstance.destroy();
  }

  const ctx = canvas.getContext('2d');
  
  // Build datasets for 4 team members
  const memberList = ['member2', 'member1', 'member3', 'member4'];
  const datasets = memberList.map(mId => {
    const savedData = localStorage.getItem(`brainwaveData_${mId}`);
    const memberData = savedData ? JSON.parse(savedData) : [];
    // The data is mapped to target columns averages
    const songData = memberData[songIndex];
    const dataVals = TARGET_COLUMNS.map(col => songData ? songData.averages[col] : 0);

    const colorInfo = {
      member3: { border: '#bf5af2', bgRadar: 'rgba(191, 90, 242, 0.15)', bgBar: 'rgba(191, 90, 242, 0.7)' }, // 문경수 - Purple
      member1: { border: '#0a84ff', bgRadar: 'rgba(10, 132, 255, 0.15)', bgBar: 'rgba(10, 132, 255, 0.7)' },   // 김한주 - Blue
      member2: { border: '#ff453a', bgRadar: 'rgba(255, 69, 58, 0.15)', bgBar: 'rgba(255, 69, 58, 0.7)' },    // 김용석 - Red
      member4: { border: '#30d158', bgRadar: 'rgba(48, 209, 88, 0.15)', bgBar: 'rgba(48, 209, 88, 0.7)' }     // 홍수민 - Green
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

  // Calculate & Add Team Average if toggled active
  if (showTeamAverage) {
    const teamAvgVals = TARGET_COLUMNS.map(col => {
      let sum = 0;
      let count = 0;
      memberList.forEach(mId => {
        const savedData = localStorage.getItem(`brainwaveData_${mId}`);
        if (savedData) {
          const memberData = JSON.parse(savedData);
          const songData = memberData[songIndex];
          if (songData) {
            sum += songData.averages[col];
            count++;
          }
        }
      });
      return count > 0 ? (sum / count) : 0;
    });

    if (chartType === 'radar') {
      datasets.push({
        label: "👥 조원 평균 (Team Avg)",
        data: teamAvgVals,
        borderColor: '#ffd700',
        backgroundColor: 'rgba(255, 215, 0, 0.15)',
        borderWidth: 3.5, // Thicker border for prominence
        pointBackgroundColor: '#ffd700',
        pointBorderColor: '#fff',
        pointHoverBackgroundColor: '#fff',
        pointHoverBorderColor: '#ffd700',
        order: -1 // Bring to front
      });
    } else {
      datasets.push({
        label: "👥 조원 평균 (Team Avg)",
        data: teamAvgVals,
        borderColor: '#ffd700',
        backgroundColor: 'rgba(255, 215, 0, 0.75)',
        borderWidth: 2,
        borderRadius: 4,
        order: -1
      });
    }
  }

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

function render13TracksTeamAverageChart(metric) {
  const canvas = document.getElementById('songTeamAverageChart');
  if (!canvas) return;

  if (songTeamAverageChartInstance) {
    songTeamAverageChartInstance.destroy();
  }

  const ctx = canvas.getContext('2d');
  const memberList = ['member2', 'member1', 'member3', 'member4'];

  // Calculate team averages/peaks/volatility for all 13 songs
  const songAverages = SONG_METADATA.map((song, songIdx) => {
    let sum = 0;
    let count = 0;

    memberList.forEach(mId => {
      const savedData = localStorage.getItem(`brainwaveData_${mId}`);
      if (savedData) {
        const memberData = JSON.parse(savedData);
        const songData = memberData[songIdx];
        if (songData && songData.rawData && songData.rawData[metric]) {
          const rawVals = songData.rawData[metric];
          if (rawVals.length > 0) {
            if (currentComparisonMode === 'peak') {
              // Peak Critical Value: the maximum stimulus achieved during the track
              sum += Math.max(...rawVals);
            } else if (currentComparisonMode === 'volatility') {
              // Volatility / Rollercoaster index: Peak minus Valley (range)
              sum += (Math.max(...rawVals) - Math.min(...rawVals));
            } else {
              // Standard Average
              sum += songData.averages[metric];
            }
            count++;
          }
        }
      }
    });

    const val = count > 0 ? (sum / count) : 0;
    return {
      index: songIdx,
      title: song.title,
      cover: song.cover || '🎵',
      value: val
    };
  });

  // Sort descending by value (Ranking style)
  songAverages.sort((a, b) => b.value - a.value);

  const labels = songAverages.map(item => `${String(item.index + 1).padStart(2, '0')}. ${item.cover} ${item.title}`);
  const dataVals = songAverages.map(item => item.value);

  // We will highlight the currently active song in the rankings with a golden glow!
  const bgColors = songAverages.map(item => {
    if (item.index === currentSongAnalysisIndex) {
      return 'rgba(255, 215, 0, 0.85)'; // Active song: glowing gold
    }
    return 'rgba(191, 90, 242, 0.35)'; // Non-active songs: cool violet
  });

  const borderColors = songAverages.map(item => {
    if (item.index === currentSongAnalysisIndex) {
      return '#ffd700'; // Gold border
    }
    return '#bf5af2'; // Purple border
  });

  const metricKo = LABELS_KO[TARGET_COLUMNS.indexOf(metric)];

  let modeLabel = '평균치';
  let tooltipPrefix = '평균';
  if (currentComparisonMode === 'peak') {
    modeLabel = '크리티컬 최고 자극치';
    tooltipPrefix = '최고치';
  } else if (currentComparisonMode === 'volatility') {
    modeLabel = '롤러코스터 변동폭';
    tooltipPrefix = '변동폭';
  }

  songTeamAverageChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: `👥 4인 ${modeLabel} (${metricKo})`,
        data: dataVals,
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: 1.5,
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y', // Horizontal Bar Chart
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          min: 0,
          max: 1.0,
          grid: { color: 'rgba(255,255,255,0.04)' },
          ticks: { color: '#94a3b8' }
        },
        y: {
          grid: { display: false },
          ticks: {
            color: (context) => {
              const index = context.index;
              if (songAverages[index] && songAverages[index].index === currentSongAnalysisIndex) {
                return '#ffd700';
              }
              return '#f8fafc';
            },
            font: (context) => {
              const index = context.index;
              const isActive = songAverages[index] && songAverages[index].index === currentSongAnalysisIndex;
              return {
                family: '-apple-system',
                size: 11,
                weight: isActive ? '700' : '500'
              };
            }
          }
        }
      },
      plugins: {
        legend: {
          display: false
        },
        tooltip: {
          enabled: true,
          callbacks: {
            label: function(context) {
              const val = context.raw;
              return ` 👥 조원 ${tooltipPrefix} ${metricKo}: ${(val * 100).toFixed(1)}%`;
            }
          }
        }
      }
    }
  });

  // Populate the text leaderboard list
  const leaderboardEl = document.getElementById('songTeamAverageLeaderboard');
  if (leaderboardEl) {
    leaderboardEl.innerHTML = '';
    songAverages.forEach((item, idx) => {
      const rank = idx + 1;
      let badge = '';
      if (rank === 1) badge = '🥇 ';
      else if (rank === 2) badge = '🥈 ';
      else if (rank === 3) badge = '🥉 ';
      else badge = `<span style="display: inline-block; width: 20px; height: 20px; border-radius: 50%; background: rgba(255,255,255,0.08); color: var(--text-secondary); text-align: center; font-size: 0.75rem; line-height: 20px; font-weight: bold; margin-right: 4px;">${rank}</span>`;

      const isActive = item.index === currentSongAnalysisIndex;
      const rowBg = isActive ? 'rgba(255, 215, 0, 0.12)' : 'rgba(255,255,255,0.02)';
      const rowBorder = isActive ? '1px solid #ffd700' : '1px solid rgba(255,255,255,0.04)';
      const textColor = isActive ? '#ffd700' : '#fff';
      
      const rowDiv = document.createElement('div');
      rowDiv.style.display = 'flex';
      rowDiv.style.justifyContent = 'space-between';
      rowDiv.style.alignItems = 'center';
      rowDiv.style.padding = '0.5rem 0.75rem';
      rowDiv.style.background = rowBg;
      rowDiv.style.border = rowBorder;
      rowDiv.style.borderRadius = '8px';
      rowDiv.style.fontSize = '0.85rem';
      rowDiv.style.cursor = 'pointer';
      rowDiv.style.transition = 'all 0.2s';
      rowDiv.innerHTML = `
        <div style="display: flex; align-items: center; gap: 0.5rem; min-width: 0;">
          ${badge}
          <span style="font-weight: 700; color: ${textColor}; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${String(item.index + 1).padStart(2, '0')}. ${item.cover} ${item.title}</span>
        </div>
        <strong style="color: ${isActive ? '#ffd700' : 'var(--accent-color)'}; font-family: monospace; font-size: 0.9rem;">${(item.value * 100).toFixed(1)}%</strong>
      `;
      rowDiv.addEventListener('click', () => {
        selectSongForAnalysis(item.index);
      });
      leaderboardEl.appendChild(rowDiv);
    });
  }
}

function renderAIEmotionInsights(songIndex) {
  const insightGrid = document.getElementById('insightGrid');
  if (!insightGrid) return;
  
  insightGrid.innerHTML = '';
  
  TARGET_COLUMNS.forEach((col, cIdx) => {
    const colNameKo = LABELS_KO[cIdx];
    let maxVal = -1;
    let winnerId = '';

    ['member2', 'member1', 'member3', 'member4'].forEach(mId => {
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

function initUploaderListeners() {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const uploadMemberSelect = document.getElementById('uploadMemberSelect');
  const uploadLog = document.getElementById('uploadLog');
  const clearLogBtn = document.getElementById('clearLogBtn');

  if (!dropZone || !fileInput || !uploadLog) return;

  // Clear log
  clearLogBtn.addEventListener('click', () => {
    uploadLog.innerHTML = `<div style="color: #64d2ff;">&gt; 업로드 로그가 초기화되었습니다.</div>`;
  });

  // Drag over
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.style.borderColor = 'var(--accent-color)';
    dropZone.style.background = 'rgba(10, 132, 255, 0.05)';
  });

  // Drag leave
  dropZone.addEventListener('dragleave', () => {
    dropZone.style.borderColor = 'rgba(255,255,255,0.2)';
    dropZone.style.background = 'rgba(255,255,255,0.01)';
  });

  // Click on dropzone
  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  // Change input file
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      processFiles(e.target.files);
    }
  });

  // Drop
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.style.borderColor = 'rgba(255,255,255,0.2)';
    dropZone.style.background = 'rgba(255,255,255,0.01)';
    
    if (e.dataTransfer.files.length > 0) {
      processFiles(e.dataTransfer.files);
    }
  });

  function logMessage(text, type = 'info') {
    const div = document.createElement('div');
    if (type === 'error') {
      div.style.color = '#ff453a';
    } else if (type === 'success') {
      div.style.color = '#30d158';
    } else if (type === 'warning') {
      div.style.color = '#ff9f0a';
    } else {
      div.style.color = '#94a3b8';
    }
    div.innerHTML = `&gt; ${text}`;
    uploadLog.appendChild(div);
    uploadLog.scrollTop = uploadLog.scrollHeight;
  }

  function extractSongNumber(filename) {
    const match = filename.match(/\d+/);
    return match ? parseInt(match[0], 10) : null;
  }

  function processFiles(files) {
    const memberKey = uploadMemberSelect.value;
    const memberName = MEMBERS[memberKey];
    logMessage(`[${memberName}] 데이터 업로드 시작... (${files.length}개 파일)`, 'info');

    // Load existing data for this member if present, or create empty list
    let memberData = [];
    const saved = localStorage.getItem(`brainwaveData_${memberKey}`);
    if (saved) {
      try {
        memberData = JSON.parse(saved);
      } catch (err) {
        memberData = [];
      }
    }

    let filesProcessed = 0;
    const validFiles = Array.from(files).filter(f => f.name.endsWith('.csv') || f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));

    if (validFiles.length === 0) {
      logMessage(`[에러] 업로드할 수 있는 유효한 파일(.csv, .xlsx, .xls)이 없습니다.`, 'error');
      return;
    }
    
    validFiles.forEach(file => {
      const songNumber = extractSongNumber(file.name);
      if (songNumber === null || songNumber < 1 || songNumber > 13) {
        logMessage(`[에러] '${file.name}'에서 노래 번호(1~13)를 찾을 수 없습니다. 건너뜁니다.`, 'error');
        filesProcessed++;
        return;
      }

      logMessage(`파일 처리 중: ${file.name} -> 노래 ${songNumber}번으로 매칭`, 'info');

      const reader = new FileReader();

      if (file.name.endsWith('.csv')) {
        reader.onload = function(e) {
          const text = e.target.result;
          Papa.parse(text, {
            header: false,
            skipEmptyLines: true,
            complete: function(results) {
              const rows = results.data;
              const stats = calculateStats(rows, file.name, songNumber);
              if (stats) {
                // Remove existing song if present
                memberData = memberData.filter(s => s.songNumber !== songNumber);
                memberData.push(stats);
                logMessage(`[성공] '${file.name}' 분석 성공! (평균 몰입도: ${(stats.averages.engagement * 100).toFixed(1)}%)`, 'success');
              } else {
                logMessage(`[에러] '${file.name}' 데이터 파싱에 실패했습니다. (뇌파 컬럼 누락)`, 'error');
              }
              checkCompletion();
            }
          });
        };
        reader.readAsText(file);
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        reader.onload = function(e) {
          const data = new Uint8Array(e.target.result);
          try {
            const workbook = XLSX.read(data, { type: 'array' });
            let allRows = [];
            workbook.SheetNames.forEach(sheetName => {
              const worksheet = workbook.Sheets[sheetName];
              const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
              if (rows && rows.length > 0) {
                allRows = allRows.concat(rows);
              }
            });

            const stats = calculateStats(allRows, file.name, songNumber);
            if (stats) {
              memberData = memberData.filter(s => s.songNumber !== songNumber);
              memberData.push(stats);
              logMessage(`[성공] '${file.name}' 분석 성공! (평균 몰입도: ${(stats.averages.engagement * 100).toFixed(1)}%)`, 'success');
            } else {
              logMessage(`[에러] '${file.name}' 데이터 파싱에 실패했습니다. (뇌파 컬럼 누락)`, 'error');
            }
          } catch (err) {
            logMessage(`[에러] '${file.name}'을 읽는 도중 오류가 발생했습니다: ${err.message}`, 'error');
          }
          checkCompletion();
        };
        reader.readAsArrayBuffer(file);
      }
    });

    function checkCompletion() {
      filesProcessed++;
      if (filesProcessed >= validFiles.length) {
        // Sort member data by songNumber ascending
        memberData.sort((a, b) => a.songNumber - b.songNumber);
        
        // Save to localStorage
        localStorage.setItem(`brainwaveData_${memberKey}`, JSON.stringify(memberData));
        logMessage(`🏁 모든 파일 처리 완료! ${memberName} 님의 대시보드가 실시간 업데이트되었습니다.`, 'success');
        
        // Trigger confetti for the uploader success!
        triggerConfetti();
      }
    }
  }
}

// ==========================================
// 💡 Presentation Navigation & Storyboards
// ==========================================

const navIntro = document.getElementById('navIntro');
const navHowWeDid = document.getElementById('navHowWeDid');
const navSeeMusic = document.getElementById('navSeeMusic');
const navSeePerson = document.getElementById('navSeePerson');
const navNbti = document.getElementById('navNbti');
const navAchievements = document.getElementById('navAchievements');
const navUploader = document.getElementById('navUploader');

const introScreen = document.getElementById('introScreen');
const howWeDidScreen = document.getElementById('howWeDidScreen');
const seeMusicScreen = document.getElementById('seeMusicScreen');
const dashboardScreen = document.getElementById('dashboardScreen');
const nbtiScreen = document.getElementById('nbtiScreen');
const teamScreen = document.getElementById('teamScreen');
const uploaderScreen = document.getElementById('uploaderScreen');

function switchScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-links button').forEach(b => b.classList.remove('active'));
  
  if (screenId === 'introScreen') {
    navIntro.classList.add('active');
    introScreen.classList.add('active');
  }
  if (screenId === 'howWeDid') {
    navHowWeDid.classList.add('active');
    howWeDidScreen.classList.add('active');
  }
  if (screenId === 'seeMusic') {
    navSeeMusic.classList.add('active');
    seeMusicScreen.classList.add('active');
    loadSongAnalysis();
  }
  if (screenId === 'seePerson') {
    navSeePerson.classList.add('active');
    dashboardScreen.classList.add('active');
    currentUserId = 'member3'; // Set default to 문경수 when entering personal dashboard
    loadDashboard();
  }
  if (screenId === 'nbti') {
    navNbti.classList.add('active');
    nbtiScreen.classList.add('active');
    loadNbtiScreen('member2'); // Default to Kim Yong Seok
    triggerConfetti();
  }
  if (screenId === 'achievements') {
    navAchievements.classList.add('active');
    teamScreen.classList.add('active');
    loadAchievementsScreen('member2'); // Default to Kim Yong Seok
    triggerConfetti();
  }
  if (screenId === 'uploader') {
    navUploader.classList.add('active');
    uploaderScreen.classList.add('active');
  }
}

// Nav items click bindings
navIntro.addEventListener('click', () => switchScreen('introScreen'));
navHowWeDid.addEventListener('click', () => switchScreen('howWeDid'));
navSeeMusic.addEventListener('click', () => switchScreen('seeMusic'));
navSeePerson.addEventListener('click', () => switchScreen('seePerson'));
navNbti.addEventListener('click', () => switchScreen('nbti'));
navAchievements.addEventListener('click', () => switchScreen('achievements'));
navUploader.addEventListener('click', () => switchScreen('uploader'));

// Nav Logo click to home
document.querySelector('.nav-logo').style.cursor = 'pointer';
document.querySelector('.nav-logo').addEventListener('click', () => switchScreen('introScreen'));

// Intro page buttons click handlers
['member2', 'member1', 'member3', 'member4'].forEach(mId => {
  const btn = document.getElementById(`intro-btn-${mId}`);
  if (btn) {
    btn.addEventListener('click', () => {
      currentUserId = mId;
      document.getElementById('userSelect').value = mId;
      switchScreen('seePerson');
    });
  }
});

const introStartBtn = document.getElementById('intro-start-btn');
if (introStartBtn) {
  introStartBtn.addEventListener('click', () => {
    switchScreen('howWeDid');
  });
}

// Stage 5 N-BTI member selector tabs
['member2', 'member1', 'member3', 'member4'].forEach(mId => {
  const tabBtn = document.getElementById(`nbti-tab-${mId}`);
  if (tabBtn) {
    tabBtn.addEventListener('click', () => {
      document.querySelectorAll('#nbtiScreen .hero-btn').forEach(b => {
        b.classList.remove('active');
        b.style.background = 'transparent';
      });
      tabBtn.classList.add('active');
      const colors = { member3: '#bf5af2', member1: '#0a84ff', member2: '#ff453a', member4: '#30d158' };
      tabBtn.style.background = `rgba(${hexToRgb(colors[mId])}, 0.15)`;
      loadNbtiScreen(mId);
    });
  }
});

// Stage 6 Achievements member selector tabs
['member2', 'member1', 'member3', 'member4'].forEach(mId => {
  const tabBtn = document.getElementById(`ach-tab-${mId}`);
  if (tabBtn) {
    tabBtn.addEventListener('click', () => {
      document.querySelectorAll('#teamScreen .hero-btn').forEach(b => {
        b.classList.remove('active');
        b.style.background = 'transparent';
      });
      tabBtn.classList.add('active');
      const colors = { member3: '#bf5af2', member1: '#0a84ff', member2: '#ff453a', member4: '#30d158' };
      tabBtn.style.background = `rgba(${hexToRgb(colors[mId])}, 0.15)`;
      loadAchievementsScreen(mId);
    });
  }
});

function hexToRgb(hex) {
  const bigint = parseInt(hex.replace('#', ''), 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `${r}, ${g}, ${b}`;
}

// ==========================================
// 🎛️ Interactive 4x5 Neural Launchpad
// ==========================================
function initLaunchpad() {
  const launchpad = document.getElementById('neuralLaunchpad');
  if (!launchpad) return;

  launchpad.innerHTML = '';
  
  const subjects = ['member2', 'member1', 'member3', 'member4'];
  const metrics = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];
  const metricsKo = ['몰입', '흥미', '활성', '스트레스', '이완'];
  const colors = {
    member3: 'rgba(191, 90, 242, 0.45)', // 문경수 - Purple
    member1: 'rgba(10, 132, 255, 0.45)', // 김한주 - Blue
    member2: 'rgba(255, 69, 58, 0.45)',  // 김용석 - Red
    member4: 'rgba(48, 209, 88, 0.45)'   // 홍수민 - Green
  };

  subjects.forEach(mId => {
    metrics.forEach((metric, mIdx) => {
      const btn = document.createElement('button');
      btn.className = 'launchpad-btn';
      btn.style.padding = '0.75rem 0.5rem';
      btn.style.fontSize = '0.8rem';
      btn.style.fontFamily = 'monospace';
      btn.style.fontWeight = '700';
      btn.style.color = '#fff';
      btn.style.border = '1px solid rgba(255,255,255,0.08)';
      btn.style.borderRadius = '8px';
      btn.style.cursor = 'pointer';
      btn.style.transition = 'all 0.2s';
      btn.style.background = 'rgba(255, 255, 255, 0.02)';
      btn.innerHTML = `${MEMBERS[mId].substring(0, 3)}<br><span style="font-size: 0.65rem; color: var(--text-secondary);">${metricsKo[mIdx]}</span>`;

      // Highlight active cell
      if (currentUserId === mId && compareMetricSelect.value === metric) {
        btn.classList.add('active');
        btn.style.background = colors[mId];
        btn.style.borderColor = '#fff';
        btn.style.boxShadow = `0 0 15px ${colors[mId].replace('0.45', '0.8')}`;
      }

      btn.addEventListener('click', () => {
        // Remove active class from all launchpad buttons
        document.querySelectorAll('.launchpad-btn').forEach(b => {
          b.style.background = 'rgba(255, 255, 255, 0.02)';
          b.style.borderColor = 'rgba(255,255,255,0.08)';
          b.style.boxShadow = 'none';
        });

        // Set active state
        btn.style.background = colors[mId];
        btn.style.borderColor = '#fff';
        btn.style.boxShadow = `0 0 15px ${colors[mId].replace('0.45', '0.8')}`;

        // Dynamic State updates
        currentUserId = mId;
        document.getElementById('userSelect').value = mId;
        compareMetricSelect.value = metric;
        
        // Load the personal dashboard
        loadDashboard();
      });

      launchpad.appendChild(btn);
    });
  });
}

// ==========================================
// 🔮 Stage 5: N-BTI Personality Diagnosis
// ==========================================
function loadNbtiScreen(mId) {
  const block = document.getElementById('nbtiReportBlock');
  if (!block) return;

  const saved = localStorage.getItem(`brainwaveData_${mId}`);
  if (!saved) {
    block.innerHTML = `<p style="text-align: center; color: red;">이 조원의 뇌파 데이터가 존재하지 않습니다. Uploader 탭에서 데이터를 업로드해 주세요.</p>`;
    return;
  }

  const parsed = JSON.parse(saved);
  
  // Calculate average stats over all songs
  let totalStats = { engagement: 0, interest: 0, excitement: 0, stress: 0, relaxation: 0 };
  let count = 0;
  
  parsed.forEach(song => {
    TARGET_COLUMNS.forEach(col => {
      totalStats[col] += song.averages[col] || 0;
    });
    count++;
  });
  
  if (count > 0) {
    TARGET_COLUMNS.forEach(col => {
      totalStats[col] = totalStats[col] / count;
    });
  }

  // Define N-BTI personality profile texts
  const PROFILE_DB = {
    member1: {
      nbti: "Z.E.N.S (Zen & Emotional Serene)",
      title: "🧘‍♂️ 개념 치유사 (The Composer)",
      bg: "rgba(10, 132, 255, 0.03)",
      border: "#0a84ff",
      glow: "rgba(10, 132, 255, 0.3)",
      description: "음악학과 작곡 전공자다운 평정심의 수호자입니다. 전체 평균 스트레스가 0.355로 팀 내에서 가장 낮고, 이완도는 0.426으로 가장 높아 음악을 들을 때 뇌가 가장 편안하고 안정적인 상태를 유지합니다. 어떤 주파수 변화도 평안하게 수용하는 힐러와 같은 신경 아키텍처를 가졌습니다.",
      quote: "전체 평균 스트레스 최저(0.355), 이완도 최고(0.426). 특히 5번 곡에서 몰입도 평균 0.596, 이완도 최고 0.857을 기록하며 음악과 완전한 물아일체에 도달했습니다. 🎧"
    },
    member3: {
      nbti: "A.I.F.C (AI Focused chipset)",
      title: "🧠 도파민 광전사 (Dopamine Berserker)",
      bg: "rgba(191, 90, 242, 0.03)",
      border: "#bf5af2",
      glow: "rgba(191, 90, 242, 0.3)",
      description: "음악을 들을 때 뇌파가 격렬하게 춤을 추는 뇌 리액션의 끝판왕입니다! 전체 평균 흥미도(0.530)와 활성도(0.422)가 압도적인 팀 내 1위를 차지하여, 고도의 인지 연산 능력과 아드레날린 분출이 결합된 파괴적이고 다이내믹한 뇌지컬 집중 반응을 여과 없이 보입니다.",
      quote: "전체 평균 흥미도(0.530) 및 활성도(0.422) 압도적 팀 내 1위! 특히 13번 곡에서 활성도 평균 0.673, 스트레스 최고 0.918, 이완도 최고 0.977이라는 파괴적인 수치를 폭발시켰습니다. ⚡"
    },
    member2: {
      nbti: "E.X.C.I (Excitement Catalyst Intellect)",
      title: "🔥 예민한 분석가 (The Analyzer)",
      bg: "rgba(255, 69, 58, 0.03)",
      border: "#ff453a",
      glow: "rgba(255, 69, 58, 0.3)",
      description: "자극에 대해 매우 섬세하고 직관적으로 반응하는 지성적 탐구자입니다. 전체 평균 스트레스가 0.404로 가장 높으면서도 곡에 따라 반응 편차가 큽니다. 매 사운드트랙마다 미세한 신호마저 놓치지 않고 예민하게 받아들여 고도의 분석적 흐름을 타는 양상을 보입니다.",
      quote: "전체 평균 스트레스 최고(0.404) 및 높은 곡별 편차. 6번 곡에서 활성도(0.447)와 스트레스(0.451)가 동시에 솟구치는 등 매 음악을 극도로 기민하고 디테일하게 지각했습니다. 🎸"
    },
    member4: {
      nbti: "I.N.T.R (Interactive Neuro receptor)",
      title: "👀 강철의 포커페이스 (Iron Wall)",
      bg: "rgba(48, 209, 88, 0.03)",
      border: "#30d158",
      glow: "rgba(48, 209, 88, 0.3)",
      description: "조형예술 전공자다운 흔들림 없는 부동의 미학적 마인드를 지녔습니다. 전체 평균 활성도가 0.222로 극도로 낮아, 웬만큼 강한 음악 자극이나 소음 유발에도 뇌파가 전혀 동요하지 않고 명상에 가까운 깊고 단단한 심신 안정을 드러냅니다.",
      quote: "전체 평균 활성도 극저치(0.222). 웬만한 음악적 자극에는 뇌가 미동도 하지 않으며, 특히 3번 곡에서는 활성도 최고치마저 0.188에 그치며 명상 수준의 평안함을 지켜냈습니다. 🎨"
    }
  };

  const profile = PROFILE_DB[mId] || PROFILE_DB.member3;

  // Compute direct emotion scores (Left-aligned) and their math complements (Right-aligned)
  const scoreEng = Math.round(totalStats.engagement * 100);
  const scoreDiv = 100 - scoreEng;

  const scoreInt = Math.round(totalStats.interest * 100);
  const scoreInd = 100 - scoreInt;

  const scoreExc = Math.round(totalStats.excitement * 100);
  const scoreCalm = 100 - scoreExc;

  const scoreStr = Math.round(totalStats.stress * 100);
  const scoreStab = 100 - scoreStr;

  const scoreRel = Math.round(totalStats.relaxation * 100);
  const scoreTen = 100 - scoreRel;

  const avatarMap = {
    member1: '/profile2.png',
    member2: '/profile3.png',
    member3: '/profile1.png',
    member4: '/profile4.png'
  };
  const avatarUrl = avatarMap[mId] || '/profile1.png';
  
  const spriteMap = {
    member1: '/sprite2.png',
    member2: '/sprite3.png',
    member3: '/sprite1.png',
    member4: '/sprite4.png'
  };
  const spriteUrl = spriteMap[mId] || '/sprite1.png';

  // Inner helper function to render beautiful horizontal MBTI dimension bars
  const renderMbtiBar = (leftName, leftScore, rightName, rightScore, labelTitle) => {
    const isLeft = leftScore >= rightScore;
    const activeColor = profile.border;
    const glowColor = profile.glow;
    return `
      <div style="display: flex; flex-direction: column; gap: 0.3rem; margin-bottom: 0.4rem;">
        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">
          <span style="color: var(--text-secondary); font-size: 0.72rem;">${labelTitle}</span>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 0.8rem; font-weight: 700; color: #fff;">
          <span style="${isLeft ? `color: ${activeColor}; text-shadow: 0 0 8px ${glowColor};` : 'color: rgba(255,255,255,0.45);'}">${leftName} ${leftScore}%</span>
          <span style="${!isLeft ? `color: ${activeColor}; text-shadow: 0 0 8px ${glowColor};` : 'color: rgba(255,255,255,0.45);'}">${rightScore}% ${rightName}</span>
        </div>
        <div style="width: 100%; height: 10px; background: rgba(255,255,255,0.04); border-radius: 5px; overflow: hidden; display: flex; position: relative; border: 0.5px solid rgba(255,255,255,0.05);">
          ${isLeft 
            ? `<div style="width: ${leftScore}%; background: linear-gradient(90deg, rgba(${hexToRgb(activeColor)}, 0.4), ${activeColor}); height: 100%; border-radius: 5px 0 0 5px;"></div>`
            : `<div style="width: ${leftScore}%; background: transparent; height: 100%;"></div>
               <div style="width: ${rightScore}%; background: linear-gradient(90deg, ${activeColor}, rgba(${hexToRgb(activeColor)}, 0.4)); height: 100%; border-radius: 0 5px 5px 0;"></div>`
          }
        </div>
      </div>
    `;
  };

  block.innerHTML = `
    <div id="pdf-report-container" style="background: ${profile.bg}; border: 2.5px solid ${profile.border}; border-radius: 24px; padding: 2rem; box-shadow: 0 0 30px ${profile.glow}; animation: fadeIn 0.5s ease-out; display: flex; flex-direction: column; gap: 2rem;">
      
      <!-- Top header -->
      <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 1rem; border-bottom: 1px solid rgba(255,255,255,0.08); padding-bottom: 1rem;">
        <div>
          <div style="font-family: monospace; font-size: 1rem; color: #ffd700; font-weight: 700; margin-bottom: 0.2rem; letter-spacing: 1px;">FINAL BRAINWAVE DIAGNOSIS</div>
          <h2 style="font-size: 1.8rem; color: #fff; font-weight: 800; margin-bottom: 0.2rem;">${MEMBERS[mId]} 님의 N-BTI 성향 보고서</h2>
          <span style="font-size: 1.25rem; font-weight: 700; color: ${profile.border}; text-shadow: 0 0 10px ${profile.glow}; font-family: monospace;">${profile.nbti}</span>
        </div>
        <button id="downloadPdfBtn" style="background: linear-gradient(135deg, ${profile.border}, #050508); border: 1.5px solid ${profile.border}; box-shadow: 0 0 15px ${profile.glow}; font-size: 0.9rem; font-weight: 700; padding: 0.6rem 1.2rem; border-radius: 12px; display: flex; align-items: center; gap: 8px;">
          📄 리포트 PDF 저장
        </button>
      </div>

      <!-- [TOP BLOCK] 16Personalities style MBTI Illustration Card -->
      <div class="nbti-grid-1" style="display: grid; grid-template-columns: 1fr 1.3fr; gap: 2rem; align-items: stretch;">
        
        <!-- Left Panel: Character illustration card -->
        <div style="background: linear-gradient(135deg, rgba(${hexToRgb(profile.border)}, 0.15) 0%, rgba(18, 18, 24, 0.75) 100%); border-radius: 20px; padding: 2rem; border: 1.5px solid rgba(${hexToRgb(profile.border)}, 0.35); text-align: center; position: relative; overflow: hidden; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 1rem; min-height: 340px; box-shadow: inset 0 0 20px rgba(${hexToRgb(profile.border)}, 0.15);">
          <!-- Decorative wave pattern backdrop -->
          <div style="position: absolute; bottom: 0; left: 0; right: 0; height: 160px; background: linear-gradient(180deg, transparent 0%, rgba(${hexToRgb(profile.border)}, 0.18) 100%); clip-path: polygon(0 40%, 100% 0, 100% 100%, 0 100%); z-index: 1;"></div>
          
          <span style="font-size: 0.85rem; font-weight: 700; color: rgba(255,255,255,0.45); text-transform: uppercase; letter-spacing: 2px; z-index: 2;">성격 유형</span>
          <h3 style="font-size: 1.7rem; color: #fff; font-weight: 800; text-shadow: 0 0 15px rgba(${hexToRgb(profile.border)}, 0.4); z-index: 2; margin: 0; line-height: 1.2;">
            ${profile.title.split('(')[0].replace(/[\u2300-\u23FF\u2600-\u27BF]/g, '').trim()}
          </h3>
          
          <!-- Full-Body Pixel Character Sprite (16Personalities style) -->
          <img src="${spriteUrl}" class="pixelated" style="height: 140px; object-fit: contain; filter: drop-shadow(0 0 15px rgba(${hexToRgb(profile.border)}, 0.5)); z-index: 2; transform: translateY(-3px); animation: characterFloat 3s ease-in-out infinite;">
          
          <span style="font-size: 1.25rem; font-weight: 800; color: ${profile.border}; letter-spacing: 2px; text-shadow: 0 0 8px rgba(${hexToRgb(profile.border)}, 0.3); z-index: 2; margin: 0; font-family: monospace;">
            ${profile.nbti.split(' ')[0]}-T
          </span>
        </div>

        <!-- Right Panel: 5 Dimensions Scales -->
        <div style="background: rgba(10, 10, 15, 0.4); border-radius: 20px; padding: 1.8rem; border: 1px solid rgba(255,255,255,0.06); display: flex; flex-direction: column; gap: 0.9rem; justify-content: center; min-height: 340px; box-shadow: 0 4px 20px rgba(0,0,0,0.25);">
          ${renderMbtiBar('몰입형 (Engagement)', scoreEng, '분산형 (Diversion)', scoreDiv, '정신 (Mind) - 몰입 vs 분산')}
          ${renderMbtiBar('흥미형 (Interest)', scoreInt, '무관심형 (Indifference)', scoreInd, '본성 (Nature) - 흥미 vs 무관심')}
          ${renderMbtiBar('활성형 (Excitement)', scoreExc, '평온형 (Calmness)', scoreCalm, '에너지 (Energy) - 활성 vs 평온')}
          ${renderMbtiBar('스트레스형 (Stress)', scoreStr, '안정형 (Stability)', scoreStab, '자아 (Identity) - 스트레스 vs 안정')}
          ${renderMbtiBar('이완형 (Relaxation)', scoreRel, '긴장형 (Tension)', scoreTen, '전술 (Tactics) - 이완 vs 긴장')}
        </div>

      </div>

      <!-- [BOTTOM BLOCK] Detailed Brainwave Analysis & canvas radar chart -->
      <div class="nbti-grid-2" style="display: grid; grid-template-columns: 1.2fr 1fr; gap: 2rem; align-items: stretch;">
        
        <!-- Left Panel: Narrative interpretation & Quote -->
        <div style="background: rgba(18, 18, 24, 0.5); border-radius: 20px; padding: 1.8rem; border: 1.5px solid rgba(255,255,255,0.05); display: flex; flex-direction: column; gap: 1.2rem; justify-content: space-between; box-shadow: 0 4px 20px rgba(0,0,0,0.2);">
          <div>
            <h3 style="font-size: 1.15rem; color: #fff; font-weight: 700; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px;">
              💎 뇌파 성향 정밀 분석
            </h3>
            <p style="font-size: 0.95rem; color: #cfd8dc; line-height: 1.7; text-align: justify; margin: 0;">
              ${profile.description}
            </p>
          </div>
          <div style="background: rgba(${hexToRgb(profile.border)}, 0.04); border: 1px dashed ${profile.border}; padding: 0.8rem 1rem; border-radius: 12px; font-weight: 500; font-size: 0.85rem; color: #ffd700; line-height: 1.4; box-shadow: inset 0 0 10px rgba(${hexToRgb(profile.border)}, 0.05);">
            💡 <strong>시그니처 분석 근거:</strong> ${profile.quote}
          </div>
        </div>

        <!-- Right Panel: Emotional Brainwave Radar Chart -->
        <div style="background: rgba(10, 10, 15, 0.4); border-radius: 20px; padding: 1.5rem; border: 1px solid rgba(255,255,255,0.06); display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 250px; position: relative; box-shadow: 0 4px 20px rgba(0,0,0,0.25);">
          <h3 style="font-size: 0.9rem; color: #fff; font-weight: 700; margin-bottom: 0.5rem; text-transform: uppercase; letter-spacing: 1px; width: 100%; text-align: center;">
            🧠 5대 감정 뇌파 지표 분석 (EEG)
          </h3>
          <div style="position: relative; width: 100%; height: 180px;">
            <canvas id="nbtiEmotionChart" style="max-height: 180px; width: 100%;"></canvas>
          </div>
        </div>

      </div>

    </div>
  `;

  // Dynamically initialize the Chart.js radar chart in the bottom-right corner
  setTimeout(() => {
    const canvas = document.getElementById('nbtiEmotionChart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const labels = ['몰입', '흥미', '활성', '스트레스', '이완'];
    const dataVals = [
      totalStats.engagement,
      totalStats.interest,
      totalStats.excitement,
      totalStats.stress,
      totalStats.relaxation
    ];

    new Chart(ctx, {
      type: 'radar',
      data: {
        labels: labels,
        datasets: [{
          label: '평균 뇌파 지표',
          data: dataVals,
          borderColor: profile.border,
          backgroundColor: `rgba(${hexToRgb(profile.border)}, 0.22)`,
          borderWidth: 2,
          pointBackgroundColor: profile.border,
          pointBorderColor: '#fff',
          pointHoverBackgroundColor: '#fff',
          pointHoverBorderColor: profile.border
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          r: {
            angleLines: { color: 'rgba(255,255,255,0.08)' },
            grid: { color: 'rgba(255,255,255,0.08)' },
            pointLabels: { color: 'rgba(255,255,255,0.5)', font: { size: 10, weight: 'bold' } },
            ticks: { display: false, maxTicksLimit: 5 },
            suggestedMin: 0,
            suggestedMax: 0.8
          }
        },
        plugins: {
          legend: { display: false }
        }
      }
    });
  }, 30);

  // Attach PDF Generation logic
  const downloadPdfBtn = document.getElementById('downloadPdfBtn');
  if (downloadPdfBtn) {
    downloadPdfBtn.addEventListener('click', () => {
      generatePDFReport(mId, profile);
    });
  }
}

function generatePDFReport(mId, profile) {
  const element = document.getElementById('pdf-report-container');
  if (!element) return;
  
  const opt = {
    margin:       0.5,
    filename:     `NeuroWav_N-BTI_${MEMBERS[mId]}.pdf`,
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true, backgroundColor: '#050508' },
    jsPDF:        { unit: 'in', format: 'letter', orientation: 'landscape' }
  };
  
  // Load loading overlay
  const loading = document.getElementById('loadingOverlay');
  if (loading) {
    loading.querySelector('.loading-text').textContent = 'Generating PDF Report...';
    loading.style.display = 'flex';
  }
  
  html2pdf().set(opt).from(element).save().then(() => {
    if (loading) loading.style.display = 'none';
  }).catch(err => {
    console.error("PDF generation failed:", err);
    if (loading) loading.style.display = 'none';
  });
}

// ==========================================
// 🏆 Stage 6: Retro Achievements Hall of Fame
// ==========================================
function loadAchievementsScreen(mId) {
  const grid = document.getElementById('teamGrid');
  if (!grid) return;

  const saved = localStorage.getItem(`brainwaveData_${mId}`);
  if (!saved) {
    grid.innerHTML = `<p style="text-align: center; color: red;">이 조원의 뇌파 데이터가 존재하지 않습니다. Uploader 탭에서 데이터를 업로드해 주세요.</p>`;
    return;
  }

  const parsed = JSON.parse(saved);

  // Compute highest and lowest stress songs
  let maxStress = -Infinity;
  let maxStressSong = "";
  let minStress = Infinity;
  let minStressSong = "";

  parsed.forEach(song => {
    const sIdx = song.songNumber;
    const meta = SONG_METADATA[sIdx - 1] || { title: `Song ${sIdx}` };
    if (song.averages.stress > maxStress) {
      maxStress = song.averages.stress;
      maxStressSong = `${sIdx}. ${meta.title}`;
    }
    if (song.averages.stress < minStress) {
      minStress = song.averages.stress;
      minStressSong = `${sIdx}. ${meta.title}`;
    }
  });

  const ACHIEVEMENTS_DB = {
    member1: {
      avatar: "/profile2.png",
      weapon: "🎹 Golden Tuning Fork (황금 소리굽쇠)",
      weaponIcon: "🎹",
      border: "#0a84ff",
      colorName: "blue",
      list: [
        { title: "🛡️ 이완의 천재 (Relaxation Deity)", desc: "어떤 변박 사운드나 웅장한 가창곡이 귓가에 들이닥쳐도, 심적 자극 없이 극락의 심리적 이완도 지수를 끝까지 유지하여 숲속 명상 힐러임을 입증함! (이완도 팀원 최고 0.426)" },
        { title: "🧊 스트레스 최저 브레이커 (Cold-Blooded)", desc: "팀원들이 가장 경악하고 스트레스를 폭발시켰던 디스토션 트랩 'Look at Me!' 구간에서도 팀내 최저 스트레스 반응을 단단히 수호해냄!" },
        { title: "🎹 물아일체 사운드 융합 (Lyrical Fusion)", desc: "Lemon Tree 감상 중 몰입도 0.596 돌파와 최고 이완도 0.857을 달성하며 선율과 신경세포가 완전히 한 덩어리로 어우러지는 기적을 보여줌." }
      ]
    },
    member3: {
      avatar: "/profile1.png",
      weapon: "⚡ Neural Beam Cannon (신경 광선 포)",
      weaponIcon: "⚡",
      border: "#bf5af2",
      colorName: "purple",
      list: [
        { title: "🧠 초집중 인공 신경망 (Neural Processor)", desc: "외부의 감정적 동요(Stress/Relaxation)에 아랑곳하지 않고, 단 1초의 딜레이도 없이 신경계를 초고속 집중 모드로 기동하는 팀 최고의 인지력 소유자!" },
        { title: "🚀 뇌파 오버클럭 캡틴 (Synapse Accelerator)", desc: "테크노 및 빠른 오토튠 음악 청취 시 시냅스의 연산 반응이 실시간으로 급증하며 고속 비트에 초정밀 응답하는 성향을 보여줌." },
        { title: "⚔️ 진격의 거인 타이탄 크러셔 (Titan Slayer)", desc: "Sawano Hiroyuki의 웅장한 진격의 거인 에픽 사운드가 귓가에 도달할 때, 뇌의 전 영역이 압도적인 전투 몰입도로 똘똘 뭉침." }
      ]
    },
    member2: {
      avatar: "/profile3.png",
      weapon: "🎸 Overdrive Distortion Guitar (지옥의 일렉기타)",
      weaponIcon: "🎸",
      border: "#ff453a",
      colorName: "red",
      list: [
        { title: "🔥 아드레날린 제트 엔진 (Adrenaline Overdrive)", desc: "메탈 기타 리프나 강렬한 록 사운드를 접할 때, 뇌가 즉각적으로 성층권을 뚫는 강력한 초활성(Excitement)과 내적 분출 상태를 감행함!" },
        { title: "🤘 헤비메탈 콘서트 라이더 (Metal Core Headbanger)", desc: "Pentakill의 묵직한 디스토션 멜로디가 전개될 때, 내면에 잠자던 전사의 투지와 내적 헤드뱅잉 본능이 뇌파에 폭발적으로 매핑됨." },
        { title: "💥 거침없는 데이터 탐색가 (Fearless Explorer)", desc: "어려운 변박 재즈 음악 속에서도 뇌가 당황(Stress)하지 않고 끊임없이 소리를 탐색하고 흥미를 느끼는 진취성을 입증함." }
      ]
    },
    member4: {
      avatar: "/profile4.png",
      weapon: "🎨 Prismatic Spectrum Staff (예술의 프리즘 요술봉)",
      weaponIcon: "🎨",
      border: "#30d158",
      colorName: "green",
      list: [
        { title: "👀 예술적 미장센 디텍터 (Aesthetic Receptor)", desc: "악기의 주파수 질감 변화 and 처음 들어보는 이색적 프렌치 인디 팝에 흥미도 지표가 실시간 번개처럼 예민하게 반응하여 최고 미학 감각 증명!" },
        { title: "🎭 극적 뇌파 큐레이션 (Dramatic Mind)", desc: "Sie ergibt sich nicht와 같이 드라마틱한 가창과 가사가 풍부한 뮤지컬 락 장르에서 뇌파 감정이 한 편의 서사시처럼 기민하게 춤을 춤." },
        { title: "✨ 호기심의 찬란한 팔레트 (Curious Palette)", desc: "곡의 인트로 구간이 재생될 때마다 '이것은 무슨 악기인가?'라는 경이로운 호기심 감수성을 팀내에서 가장 크게 터트림." }
      ]
    }
  };

  const ach = ACHIEVEMENTS_DB[mId] || ACHIEVEMENTS_DB.member3;
  
  const spriteMap = {
    member1: '/sprite2.png',
    member2: '/sprite3.png',
    member3: '/sprite1.png',
    member4: '/sprite4.png'
  };
  const spriteUrl = spriteMap[mId] || '/sprite1.png';

  let listHtml = '';
  ach.list.forEach(item => {
    listHtml += `
      <div style="background: rgba(0, 0, 0, 0.35); border-left: 5px solid ${ach.border}; border-radius: 12px; padding: 1.5rem; text-align: left; display: flex; flex-direction: column; gap: 0.5rem; border: 1px solid rgba(255,255,255,0.04); border-left: 5px solid ${ach.border}; transition: transform 0.2s;">
        <h4 style="font-size: 1.15rem; font-weight: 700; color: #fff;">${item.title}</h4>
        <p style="font-size: 0.88rem; color: var(--text-secondary); line-height: 1.6;">${item.desc}</p>
      </div>
    `;
  });

  grid.innerHTML = `
    <div style="background: rgba(28,28,30,0.75); border: 2.5px solid ${ach.border}; border-radius: var(--radius-lg); padding: 2.5rem; display: flex; flex-direction: column; gap: 2.5rem; text-align: center; box-shadow: 0 0 25px rgba(${hexToRgb(ach.border)}, 0.15); animation: fadeIn 0.5s ease-out;">
      <div style="display: flex; flex-direction: column; align-items: center; gap: 1rem;">
        <div style="display: flex; gap: 1.5rem; align-items: center; justify-content: center; flex-wrap: wrap;">
          <img src="${ach.avatar}" style="width: 100px; height: 100px; border-radius: 50%; object-fit: cover; border: 3px solid ${ach.border}; box-shadow: 0 0 15px rgba(${hexToRgb(ach.border)}, 0.3);">
          <img src="${spriteUrl}" class="pixelated" style="height: 140px; object-fit: contain; filter: drop-shadow(0 0 15px rgba(${hexToRgb(ach.border)}, 0.5)); animation: characterFloat 3s ease-in-out infinite;">
        </div>
        <h2 style="font-size: 2.2rem; font-weight: 800; color: #fff; margin: 0;">${MEMBERS[mId]} 님의 최종 전설적 업적</h2>
        <div style="display: flex; align-items: center; gap: 8px; background: rgba(${hexToRgb(ach.border)}, 0.15); color: #fff; padding: 0.4rem 1.2rem; border-radius: 30px; font-size: 0.95rem; font-weight: 700; border: 1px solid ${ach.border};">
          <span>${ach.weaponIcon} 장착 무기: <strong>${ach.weapon}</strong></span>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1.5rem; width: 100%; margin-top: 1rem;">
        ${listHtml}
      </div>

      <!-- Stress Track Summary -->
      <div class="stress-summary-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-top: 1rem; flex-wrap: wrap;">
        <div style="background: rgba(255, 69, 58, 0.05); border: 1px dashed rgba(255, 69, 58, 0.3); border-radius: 16px; padding: 1.5rem; text-align: left;">
          <h4 style="font-size: 1rem; color: #ff453a; font-weight: 700; margin-bottom: 0.5rem;">🤯 뇌에 부담을 준 기피 트랙 (Max Stress)</h4>
          <span style="font-size: 1.15rem; color: #fff; font-weight: 700;">${maxStressSong}</span>
          <p style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 0.4rem; line-height: 1.4;">
            이 트랙 청취 시 스트레스 지표가 최고조인 <strong>${(maxStress * 100).toFixed(1)}%</strong>에 수렴하여 뇌가 본능적으로 기피하는 반응을 나타냈습니다.
          </p>
        </div>
        <div style="background: rgba(48, 209, 88, 0.05); border: 1px dashed rgba(48, 209, 88, 0.3); border-radius: 16px; padding: 1.5rem; text-align: left;">
          <h4 style="font-size: 1rem; color: #30d158; font-weight: 700; margin-bottom: 0.5rem;">🧘‍♂️ 뇌를 평온하게 한 치유 트랙 (Min Stress)</h4>
          <span style="font-size: 1.15rem; color: #fff; font-weight: 700;">${minStressSong}</span>
          <p style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 0.4rem; line-height: 1.4;">
            이 트랙 청취 시 스트레스 지표가 최저치인 <strong>${(minStress * 100).toFixed(1)}%</strong>로 감소하여 뇌 세포가 극도로 평안하고 안정적인 상태에 이완되었습니다.
          </p>
        </div>
      </div>
    </div>
  `;
}

// ==========================================
// 💡 Real-time Track Selector Syncing
// ==========================================
function initDashboardSongSelect() {
  const songSelect = document.getElementById('dashboardSongSelect');
  const compareSelect = document.getElementById('comparePanelSongSelect');
  if (!songSelect) return;

  songSelect.innerHTML = '';
  if (compareSelect) compareSelect.innerHTML = '';
  
  SONG_METADATA.forEach((song, idx) => {
    const opt = document.createElement('option');
    opt.value = idx;
    opt.textContent = `${idx + 1}. ${song.title} - ${song.artist} (${song.cover})`;
    songSelect.appendChild(opt);
    
    if (compareSelect) {
      const opt2 = document.createElement('option');
      opt2.value = idx;
      opt2.textContent = `${idx + 1}. ${song.title} - ${song.artist} (${song.cover})`;
      compareSelect.appendChild(opt2);
    }
  });

  songSelect.addEventListener('change', (e) => {
    selectSong(parseInt(e.target.value));
  });
  
  if (compareSelect) {
    compareSelect.addEventListener('change', (e) => {
      selectSong(parseInt(e.target.value));
    });
  }
}

// Overwrite loadDashboard to add Launchpad & Song Select initializers
const oldLoadDashboard = loadDashboard;
loadDashboard = function() {
  oldLoadDashboard();
  initLaunchpad();
  initDashboardSongSelect();
  
  // Sync the dropdown value
  const songSelect = document.getElementById('dashboardSongSelect');
  if (songSelect) {
    songSelect.value = activeSongIndex;
  }
  const compareSelect = document.getElementById('comparePanelSongSelect');
  if (compareSelect) {
    compareSelect.value = activeSongIndex;
  }
};

// Initialize on script load
initMockPlayer();
initDashboardTabs();
initUploaderListeners();
initExportButtons();

// Initially show Stage 1 Intro
switchScreen('introScreen');

// ==========================================
// 📥 Client-Side Excel Export for Tableau
// ==========================================
function initExportButtons() {
  const btnExportCurrentSongRaw = document.getElementById('btnExportCurrentSongRaw');
  if (btnExportCurrentSongRaw) {
    btnExportCurrentSongRaw.addEventListener('click', () => {
      const idx = currentSongAnalysisIndex;
      const songMeta = SONG_METADATA[idx];
      const currentSongNum = idx + 1;
      
      const rows = [];
      const memberKeys = ['member2', 'member1', 'member3', 'member4'];
      const genres = ["멜로딕 테크노", "K-POP 댄스", "버블검 팝", "뮤지컬 락", "어쿠스틱 인디 팝", "라틴 재즈 피아노", "J-POP", "프렌치 인디 팝", "헤비 메탈", "오케스트라 에픽", "정통 클래식 피아노", "인디 기타 인스트루멘탈", "디스토션 트랩/힙합"];
      const songGenre = genres[idx] || "기타";
      
      memberKeys.forEach(mKey => {
        const saved = localStorage.getItem(`brainwaveData_${mKey}`);
        if (saved) {
          const parsed = JSON.parse(saved);
          const songData = parsed.find(s => s.songNumber === currentSongNum);
          if (songData && songData.rawData) {
            const raw = songData.rawData;
            const maxLen = Math.max(
              raw.engagement ? raw.engagement.length : 0,
              raw.interest ? raw.interest.length : 0,
              raw.excitement ? raw.excitement.length : 0,
              raw.stress ? raw.stress.length : 0,
              raw.relaxation ? raw.relaxation.length : 0
            );
            
            for (let t = 0; t < maxLen; t++) {
              rows.push({
                "시간(초)": t + 1,
                "조원 이름": MEMBERS[mKey],
                "곡 제목": songMeta.title,
                "장르": songGenre,
                "몰입도 (Engagement)": raw.engagement && raw.engagement[t] !== undefined ? parseFloat(raw.engagement[t].toFixed(4)) : null,
                "흥미도 (Interest)": raw.interest && raw.interest[t] !== undefined ? parseFloat(raw.interest[t].toFixed(4)) : null,
                "활성도 (Excitement)": raw.excitement && raw.excitement[t] !== undefined ? parseFloat(raw.excitement[t].toFixed(4)) : null,
                "스트레스 (Stress)": raw.stress && raw.stress[t] !== undefined ? parseFloat(raw.stress[t].toFixed(4)) : null,
                "이완도 (Relaxation)": raw.relaxation && raw.relaxation[t] !== undefined ? parseFloat(raw.relaxation[t].toFixed(4)) : null
              });
            }
          }
        }
      });
      
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, "Raw_Time_Series");
      XLSX.writeFile(wb, `neurowav_track_${currentSongNum}_raw_data.xlsx`);
    });
  }

  const btnExportAllAverages = document.getElementById('btnExportAllAverages');
  if (btnExportAllAverages) {
    btnExportAllAverages.addEventListener('click', () => {
      const averagesData = [];
      const memberKeys = ['member2', 'member1', 'member3', 'member4'];
      const memberDetails = {
        member1: { nbti: "Z.E.N.S (Zen & Emotional Serene)", title: "개념 치유사 (The Composer)" },
        member2: { nbti: "E.X.C.I (Excitement Catalyst Intellect)", title: "아드레날린 락커 (The Rocker)" },
        member3: { nbti: "A.I.F.C (AI Focused chipset)", title: "AI 집중 마스터 (The AI Expert)" },
        member4: { nbti: "I.N.T.R (Interactive Neuro receptor)", title: "호기심 미술가 (The Artist)" }
      };
      const genres = ["멜로딕 테크노", "K-POP 댄스", "버블검 팝", "뮤지컬 락", "어쿠스틱 인디 팝", "라틴 재즈 피아노", "J-POP", "프렌치 인디 팝", "헤비 메탈", "오케스트라 에픽", "정통 클래식 피아노", "인디 기타 인스트루멘탈", "디스토션 트랩/힙합"];

      memberKeys.forEach(mKey => {
        const saved = localStorage.getItem(`brainwaveData_${mKey}`);
        if (saved) {
          const parsed = JSON.parse(saved);
          parsed.forEach((song, idx) => {
            const meta = SONG_METADATA[idx];
            averagesData.push({
              "조원 이름": MEMBERS[mKey],
              "N-BTI 유형": memberDetails[mKey].nbti,
              "고유 타이틀": memberDetails[mKey].title,
              "트랙 번호": song.songNumber,
              "곡 제목": meta.title,
              "아티스트": meta.artist,
              "장르": genres[idx] || "기타",
              "몰입도 (Engagement)": parseFloat(song.averages.engagement.toFixed(4)),
              "흥미도 (Interest)": parseFloat(song.averages.interest.toFixed(4)),
              "활성도 (Excitement)": parseFloat(song.averages.excitement.toFixed(4)),
              "스트레스 (Stress)": parseFloat(song.averages.stress.toFixed(4)),
              "이완도 (Relaxation)": parseFloat(song.averages.relaxation.toFixed(4))
            });
          });
        }
      });
      
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(averagesData);
      XLSX.utils.book_append_sheet(wb, ws, "Song_Averages");
      XLSX.writeFile(wb, "neurowav_13_tracks_averages.xlsx");
    });
  }

  const btnExportComparativeRaw = document.getElementById('btnExportComparativeRaw');
  if (btnExportComparativeRaw) {
    btnExportComparativeRaw.addEventListener('click', () => {
      const idx = activeSongIndex;
      const songMeta = SONG_METADATA[idx];
      const currentSongNum = idx + 1;
      const metric = compareMetricSelect.value;
      const metricLabel = {
        engagement: '몰입도(Engagement)',
        interest: '흥미도(Interest)',
        excitement: '활성도(Excitement)',
        stress: '스트레스(Stress)',
        relaxation: '이완도(Relaxation)'
      }[metric];

      const rows = [];
      const memberKeys = ['member2', 'member1', 'member3', 'member4'];

      memberKeys.forEach(mKey => {
        const saved = localStorage.getItem(`brainwaveData_${mKey}`);
        if (saved) {
          const parsed = JSON.parse(saved);
          const songData = parsed.find(s => s.songNumber === currentSongNum);
          if (songData && songData.rawData && songData.rawData[metric]) {
            const rawVals = songData.rawData[metric];
            rawVals.forEach((val, t) => {
              rows.push({
                "시간(초)": t + 1,
                "조원 이름": MEMBERS[mKey],
                "곡 제목": songMeta.title,
                "비교 대상 감정": metricLabel,
                "감정 지표 수치": parseFloat(val.toFixed(4))
              });
            });
          }
        }
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, "Comparative_Data");
      XLSX.writeFile(wb, `neurowav_track_${currentSongNum}_${metric}_comparative.xlsx`);
    });
  }
}
