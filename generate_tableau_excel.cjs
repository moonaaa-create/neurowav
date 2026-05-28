const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 1. Load data
const preloadedDataPath = path.join(__dirname, 'src', 'preloadedData.json');
if (!fs.existsSync(preloadedDataPath)) {
  console.error("preloadedData.json not found!");
  process.exit(1);
}
const preloadedData = JSON.parse(fs.readFileSync(preloadedDataPath, 'utf8'));

const SONG_METADATA = [
  { title: "Syren", artist: "Anyma", genre: "멜로딕 테크노" },
  { title: "Playing with Fire", artist: "BLACKPINK", genre: "K-POP 댄스" },
  { title: "Call Me Maybe", artist: "Carly Rae Jepsen", genre: "버블검 팝" },
  { title: "Sie ergibt sich nicht", artist: "Chang Eun Ah", genre: "뮤지컬 락" },
  { title: "Lemon Tree", artist: "Fools Garden", genre: "어쿠스틱 인디 팝" },
  { title: "Spain", artist: "Jesus Molina", genre: "라틴 재즈 피아노" },
  { title: "Jane Doe", artist: "Kenshi Yonezu", genre: "J-POP" },
  { title: "Peligrosa", artist: "Ojos", genre: "프렌치 인디 팝" },
  { title: "Lost Chapter", artist: "Pentakill, Jorn", genre: "헤비 메탈" },
  { title: "Attack on Titan", artist: "Sawano Hiroyuki", genre: "오케스트라 에픽" },
  { title: "6 Moments musicaux, Op. 16 : No. 4 in E", artist: "Sergei Rachmaninoff", genre: "정통 클래식 피아노" },
  { title: "Shoreditch", artist: "Vard", genre: "인디 기타 인스트루멘탈" },
  { title: "Look at Me!", artist: "XXTENTACION", genre: "디스토션 트랩/힙합" }
];

const MEMBERS = {
  member1: { name: "김한주", nbti: "Z.E.N.S (Zen & Emotional Serene)", title: "개념 치유사 (The Composer)" },
  member2: { name: "김용석", nbti: "E.X.C.I (Excitement Catalyst Intellect)", title: "예민한 분석가 (The Analyzer)" },
  member3: { name: "문경수", nbti: "A.I.F.C (AI Focused chipset)", title: "도파민 광전사 (Dopamine Berserker)" },
  member4: { name: "홍수민", nbti: "I.N.T.R (Interactive Neuro receptor)", title: "강철의 포커페이스 (Iron Wall)" }
};

const EMOTION_MAP = {
  engagement: "몰입도 (Engagement)",
  interest: "흥미도 (Interest)",
  excitement: "활성도 (Excitement)",
  stress: "스트레스 (Stress)",
  relaxation: "이완도 (Relaxation)"
};

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];

// ----------------------------------------------------
// [시트 1] EEG_Averages
// ----------------------------------------------------
console.log("Generating EEG_Averages Sheet rows...");
const eegAveragesRows = [];

Object.keys(preloadedData).forEach(mId => {
  const member = MEMBERS[mId];
  const songList = preloadedData[mId];
  
  songList.forEach((song, idx) => {
    const meta = SONG_METADATA[idx];
    
    TARGET_COLUMNS.forEach(col => {
      const avgVal = parseFloat(song.averages[col].toFixed(4));
      
      // Calculate max value from rawData if it exists and has items
      let maxVal = avgVal;
      if (song.rawData && song.rawData[col] && song.rawData[col].length > 0) {
        maxVal = parseFloat(Math.max(...song.rawData[col]).toFixed(4));
      }
      
      eegAveragesRows.push({
        "Person": member.name,
        "Track_No": song.songNumber,
        "Title": meta.title,
        "Genre": meta.genre,
        "Emotion": EMOTION_MAP[col],
        "Avg_Value": avgVal,
        "Max_Value": maxVal
      });
    });
  });
});

// ----------------------------------------------------
// [시트 2] Genre_Metadata
// ----------------------------------------------------
console.log("Generating Genre_Metadata Sheet rows...");
const uniqueGenres = [...new Set(SONG_METADATA.map(meta => meta.genre))];
const genreRows = uniqueGenres.map(g => ({
  "Genre": g,
  "Genre_Description": ""
}));

// ----------------------------------------------------
// [시트 3] N_BTI_Results
// ----------------------------------------------------
console.log("Generating N_BTI_Results Sheet rows...");
const nbtiRows = [
  {
    "Person": "김한주",
    "N_BTI_Title": "개념 치유사 (The Composer)",
    "Achievement_1": "최고의 스트레스 돌파!",
    "Achievement_2": "이완의 천재, 냉혈인",
    "Achievement_3": "몰입천재",
    "Description": "전체 평균 스트레스가 0.355로 팀 내 최저, 이완도는 0.426으로 최고임. 특히 5번 곡에서 몰입도 평균 0.596, 이완도 최고 0.857을 기록하며 음악과 물아일체됨.",
    "Avg_Stress": 0.3558,
    "Avg_Relaxation": 0.4267
  },
  {
    "Person": "김용석",
    "N_BTI_Title": "예민한 분석가 (The Analyzer)",
    "Achievement_1": "아드레날린 제트 엔진",
    "Achievement_2": "헤비메탈 콘서트 라이더",
    "Achievement_3": "거침없는 데이터 탐색가",
    "Description": "전체 평균 스트레스 최고(0.404) 및 높은 곡별 편차. 6번 곡에서 활성도(0.447)와 스트레스(0.451)가 동시에 솟구치는 등 매 음악을 극도로 기민하고 디테일하게 지각했습니다.",
    "Avg_Stress": 0.4043,
    "Avg_Relaxation": 0.3837
  },
  {
    "Person": "문경수",
    "N_BTI_Title": "도파민 광전사 (Dopamine Berserker)",
    "Achievement_1": "초집중 인공 신경망",
    "Achievement_2": "뇌파 오버클럭 캡틴",
    "Achievement_3": "진격의 거인 타이탄 크러셔",
    "Description": "외부의 감정적 동요(Stress/Relaxation)에 아랑곳하지 않고, 단 1초의 딜레이도 없이 신경계를 초고속 집중 모드로 기동하는 팀 최고의 인지력 소유자!",
    "Avg_Stress": 0.4074,
    "Avg_Relaxation": 0.4078
  },
  {
    "Person": "홍수민",
    "N_BTI_Title": "강철의 포커페이스 (Iron Wall)",
    "Achievement_1": "예술적 미장센 디텍터",
    "Achievement_2": "극적 뇌파 큐레이션",
    "Achievement_3": "호기심의 찬란한 팔레트",
    "Description": "조형예술 전공자다운 흔들림 없는 부동의 미학적 마인드를 지녔습니다. 전체 평균 활성도가 0.222로 극도로 낮아, 뇌파가 전혀 동요하지 않고 명상에 가까운 깊고 단단한 안정을 드러냅니다.",
    "Avg_Stress": 0.3561,
    "Avg_Relaxation": 0.4081
  }
];

// Create workbook
console.log("Writing XLSX sheets...");
const wb = XLSX.utils.book_new();

const wsAverages = XLSX.utils.json_to_sheet(eegAveragesRows);
XLSX.utils.book_append_sheet(wb, wsAverages, "EEG_Averages");

const wsGenre = XLSX.utils.json_to_sheet(genreRows);
XLSX.utils.book_append_sheet(wb, wsGenre, "Genre_Metadata");

const wsNbti = XLSX.utils.json_to_sheet(nbtiRows);
XLSX.utils.book_append_sheet(wb, wsNbti, "N_BTI_Results");

const outputFilePath = '/Users/a111/Downloads/neurowav_tableau_dashboard_data.xlsx';
XLSX.writeFile(wb, outputFilePath);
console.log(`Successfully generated and saved Tableau Dashboard data to: ${outputFilePath}`);
