const fs = require('fs');
const path = require('path');

const preloadedDataPath = path.join(__dirname, 'src', 'preloadedData.json');
const preloadedData = JSON.parse(fs.readFileSync(preloadedDataPath, 'utf8'));

const MEMBERS = {
  member1: "김한주",
  member2: "김용석",
  member3: "문경수",
  member4: "홍수민"
};

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];

Object.keys(preloadedData).forEach(mId => {
  const songList = preloadedData[mId];
  const sums = { engagement: 0, interest: 0, excitement: 0, stress: 0, relaxation: 0 };
  let count = 0;
  
  songList.forEach(song => {
    TARGET_COLUMNS.forEach(col => {
      sums[col] += song.averages[col] || 0;
    });
    count++;
  });
  
  console.log(`${MEMBERS[mId]} (${mId}):`);
  TARGET_COLUMNS.forEach(col => {
    const avg = sums[col] / count;
    console.log(`  ${col}: ${(avg * 100).toFixed(2)}%`);
  });
});
