const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Load raw data
const preloadedDataPath = path.join(__dirname, 'src', 'preloadedData.json');
if (!fs.existsSync(preloadedDataPath)) {
  console.error("preloadedData.json not found!");
  process.exit(1);
}

const preloadedData = JSON.parse(fs.readFileSync(preloadedDataPath, 'utf8'));

const SONG_METADATA = [
  { title: "Syren", artist: "Anyma", cover: "🎵", genre: "멜로딕 테크노" },
  { title: "Playing with Fire", artist: "BLACKPINK", cover: "🔥", genre: "K-POP 댄스" },
  { title: "Call Me Maybe", artist: "Carly Rae Jepsen", cover: "📱", genre: "버블검 팝" },
  { title: "Sie ergibt sich nicht", artist: "Chang Eun Ah", cover: "🎭", genre: "뮤지컬 락" },
  { title: "Lemon Tree", artist: "Fools Garden", cover: "🍋", genre: "어쿠스틱 인디 팝" },
  { title: "Spain", artist: "Jesus Molina", cover: "🇪🇸", genre: "라틴 재즈 피아노" },
  { title: "Jane Doe", artist: "Kenshi Yonezu", cover: "👤", genre: "J-POP" },
  { title: "Peligrosa", artist: "Ojos", cover: "⚠️", genre: "프렌치 인디 팝" },
  { title: "Lost Chapter", artist: "Pentakill, Jorn", cover: "📖", genre: "헤비 메탈" },
  { title: "Attack on Titan", artist: "Sawano Hiroyuki", cover: "⚔️", genre: "오케스트라 에픽" },
  { title: "6 Moments musicaux, Op. 16 : No. 4 in E", artist: "Sergei Rachmaninoff", cover: "🎹", genre: "정통 클래식 피아노" },
  { title: "Shoreditch", artist: "Vard", cover: "🎸", genre: "인디 기타 인스트루멘탈" },
  { title: "Look at Me!", artist: "XXTENTACION", cover: "💥", genre: "디스토션 트랩/힙합" }
];

const MEMBERS = {
  member1: { name: "김한주", nbti: "Z.E.N.S (Zen & Emotional Serene)", title: "개념 치유사 (The Composer)" },
  member2: { name: "김용석", nbti: "E.X.C.I (Excitement Catalyst Intellect)", title: "아드레날린 락커 (The Rocker)" },
  member3: { name: "문경수", nbti: "A.I.F.C (AI Focused chipset)", title: "AI 집중 마스터 (The AI Expert)" },
  member4: { name: "홍수민", nbti: "I.N.T.R (Interactive Neuro receptor)", title: "호기심 미술가 (The Artist)" }
};

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];

// ----------------------------------------------------
// 1. Export Averages Dataset
// ----------------------------------------------------
console.log("Preparing Averages dataset...");
const averagesData = [];

Object.keys(preloadedData).forEach(mId => {
  const member = MEMBERS[mId];
  const songList = preloadedData[mId];
  
  songList.forEach((song, idx) => {
    const meta = SONG_METADATA[idx];
    const row = {
      "조원 ID": mId,
      "조원 이름": member.name,
      "N-BTI 유형": member.nbti,
      "고유 타이틀": member.title,
      "트랙 번호": song.songNumber,
      "곡 제목": meta.title,
      "아티스트": meta.artist,
      "장르": meta.genre,
      "몰입도 (Engagement)": parseFloat(song.averages.engagement.toFixed(4)),
      "흥미도 (Interest)": parseFloat(song.averages.interest.toFixed(4)),
      "활성도 (Excitement)": parseFloat(song.averages.excitement.toFixed(4)),
      "스트레스 (Stress)": parseFloat(song.averages.stress.toFixed(4)),
      "이완도 (Relaxation)": parseFloat(song.averages.relaxation.toFixed(4))
    };
    averagesData.push(row);
  });
});

const wbAverages = XLSX.utils.book_new();
const wsAverages = XLSX.utils.json_to_sheet(averagesData);
XLSX.utils.book_append_sheet(wbAverages, wsAverages, "Song_Averages");
const averagesFile = path.join(__dirname, "neurowav_song_averages.xlsx");
XLSX.writeFile(wbAverages, averagesFile);
console.log(`Saved song averages to: ${averagesFile}`);

// ----------------------------------------------------
// 2. Export Raw Time-Series Dataset
// ----------------------------------------------------
console.log("Preparing Raw Time-Series dataset (this may take a few seconds)...");
const timeSeriesData = [];

Object.keys(preloadedData).forEach(mId => {
  const member = MEMBERS[mId];
  const songList = preloadedData[mId];
  
  songList.forEach((song, idx) => {
    const meta = SONG_METADATA[idx];
    const raw = song.rawData;
    
    // Find maximum length of raw values across columns to iterate
    const maxLen = Math.max(
      raw.engagement ? raw.engagement.length : 0,
      raw.interest ? raw.interest.length : 0,
      raw.excitement ? raw.excitement.length : 0,
      raw.stress ? raw.stress.length : 0,
      raw.relaxation ? raw.relaxation.length : 0
    );
    
    for (let t = 0; t < maxLen; t++) {
      const row = {
        "시간(초)": t + 1,
        "조원 이름": member.name,
        "트랙 번호": song.songNumber,
        "곡 제목": meta.title,
        "장르": meta.genre,
        "몰입도 (Engagement)": raw.engagement && raw.engagement[t] !== undefined ? parseFloat(raw.engagement[t].toFixed(4)) : null,
        "흥미도 (Interest)": raw.interest && raw.interest[t] !== undefined ? parseFloat(raw.interest[t].toFixed(4)) : null,
        "활성도 (Excitement)": raw.excitement && raw.excitement[t] !== undefined ? parseFloat(raw.excitement[t].toFixed(4)) : null,
        "스트레스 (Stress)": raw.stress && raw.stress[t] !== undefined ? parseFloat(raw.stress[t].toFixed(4)) : null,
        "이완도 (Relaxation)": raw.relaxation && raw.relaxation[t] !== undefined ? parseFloat(raw.relaxation[t].toFixed(4)) : null
      };
      timeSeriesData.push(row);
    }
  });
});

const wbRaw = XLSX.utils.book_new();
const wsRaw = XLSX.utils.json_to_sheet(timeSeriesData);
XLSX.utils.book_append_sheet(wbRaw, wsRaw, "Raw_Time_Series");
const rawFile = path.join(__dirname, "neurowav_raw_time_series.xlsx");
XLSX.writeFile(wbRaw, rawFile);
console.log(`Saved raw time series to: ${rawFile}`);

console.log("All data exported successfully for Tableau!");
