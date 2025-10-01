const API_KEY = '';

const TOURNAMENT_URL = '';
const INPUT_SPREADSHEET_ID = '';
const OUTPUT_SPREADSHEET_ID = '';
const rounds = 8;
const players = 24;

function refreshPairings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Matches");

  const lastRow = sheet.getLastRow();
  const formulas = sheet.getRange(2, 6, 1+rounds*players/2, 4).getFormulas();
   //sheet.getRange("A2:K" + lastRow).clearContent();
  sheet.getRange("A14:E" + lastRow).clearContent(); //keep R1, and scores will show blank after input sheet is cleared.

  const { idToName } = internal.fetchParticipants();
  const matches = internal.fetchMatches();
  const currentRound = internal.getCurrentRound(matches);

  if (!currentRound) {
    sheet.appendRow(['No open matches.']);
    if (formulas.length) {
      sheet.getRange(2, 6, formulas.length, 4).setFormulas(formulas);
    }
  return;
  }

  const rows = matches.map(m => {
    const p1 = idToName[m.player1_id] || 'TBD';
    const p2 = idToName[m.player2_id] || 'TBD';

    return [
      m.id,
      m.round,
      p1,
      p2,
      idToName[m.winner_id] || '',
      '', '', '', '',  // Columns F-I for formulas
      m.scores_csv || '',
      m.state === 'complete' ? 'Reported' : 'Pending Report'
    ];
  });

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    if (formulas.length) {
      sheet.getRange(2, 6, formulas.length, 4).setFormulas(formulas);
    }
  }

  sheet.getRange("I:J").setNumberFormat("@");
}

function updateInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allMatchesSheet = ss.getSheetByName("All Matches");
  const data = allMatchesSheet.getDataRange().getValues();
  const headers = data.shift();

  const roundIdx = headers.indexOf("Round");
  const p1Idx = headers.indexOf("Player 1");
  const p2Idx = headers.indexOf("Player 2");
  const matchIdIdx = headers.indexOf("Match ID");

  if ([roundIdx, p1Idx, p2Idx, matchIdIdx].includes(-1)) {
    throw new Error("Missing required columns");
  }

  // Get latest round number
  const rounds = [...new Set(data.map(r => r[roundIdx]).filter(n => typeof n === "number"))];
  const latestRound = Math.max(...rounds);

  // Filter rows for latest round in the order they appear
  const roundRows = data.filter(r => r[roundIdx] === latestRound);

  const { nameToId } = internal.fetchParticipants();
  const matches = roundRows.map(r => {
    const p1 = r[p1Idx], p2 = r[p2Idx];
    return [
      "",
      p1,
      internal.calculatePointsBeforeRound(data, p1, latestRound),
      "",
      p2,
      internal.calculatePointsBeforeRound(data, p2, latestRound),
      "", "", "", "", "" // Game 1–3, Score, Winner
    ];
  });

  const targetSS = SpreadsheetApp.openById(INPUT_SPREADSHEET_ID);
  let targetSheet = targetSS.getSheetByName("R" + latestRound) || targetSS.insertSheet("R" + latestRound);

  // Clear only columns B–L (11 columns), leaving col A intact
  const lastRow = targetSheet.getLastRow();
  if (lastRow > 1) {
    targetSheet.getRange(2, 2, lastRow - 1, 11).clearContent();
  }

  // Write headers into B–L
  const headersBL = ["Player 1 ID", "Player 1", "Points", "Player 2 ID", "Player 2", "Points", "Game 1", "Game 2", "Game 3", "Score", "Winner"];
  targetSheet.getRange(1, 2, 1, headersBL.length).setValues([headersBL]);

  // Write matches starting row 2, col B
  if (matches.length > 0) {
    targetSheet.getRange(2, 2, matches.length, matches[0].length).setValues(matches);
  }

  internal.writePairingsToOutput(latestRound, matches);
  update_droplist(latestRound);
}


function reportMatches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Matches");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const matchIdIdx  = headers.indexOf("Match ID");
  const set1Idx     = headers.indexOf("Set 1"); // Column F
  const set2Idx     = headers.indexOf("Set 2"); // Column G
  const set3Idx     = headers.indexOf("Set 3"); // Column H
  const matchScoreIdx = headers.indexOf("Match Score"); // Column I
  const csvIdx      = headers.indexOf("Score (CSV)"); // Column J
  const winnerIdx   = headers.indexOf("Winner");
  const p1Idx       = headers.indexOf("Player 1");
  const p2Idx       = headers.indexOf("Player 2");
  const statusIdx   = headers.indexOf("Status");

  //if ([matchIdIdx, set1Idx, set2Idx, set3Idx, matchScoreIdx, csvIdx, winnerIdx, p1Idx, p2Idx, statusIdx].includes(-1)) {
    if ([matchIdIdx, set1Idx, set2Idx, set3Idx, matchScoreIdx, p1Idx, p2Idx, statusIdx].includes(-1)) {
    throw new Error("Missing required columns in 'All Matches'");
  }

  const { nameToId } = internal.fetchParticipants();

  data.forEach((row, i) => {
    const matchId    = row[matchIdIdx];
    const set1       = row[set1Idx];
    const set2       = row[set2Idx];
    const set3       = row[set3Idx];
    const matchScore = row[matchScoreIdx];
    const p1         = row[p1Idx];
    const p2         = row[p2Idx];
    const status     = row[statusIdx];

    // Only send if at least one set score exists
    const sets = [set1, set2, set3].filter(s => s && String(s).includes("-"));
    if (sets.length === 0) return;
    if (status=="Reported") return;

    // Build scores_csv
    const scoresCsv = sets.join(",");

    // Determine winner
    const scoreParts = matchScore ? matchScore.split("-").map(Number) : null;
    if (!scoreParts || scoreParts.length !== 2 || scoreParts.some(isNaN)) {
      Logger.log(`Skipping row ${i+2}: Invalid match score "${matchScore}"`);
      return;
    }
    const winnerName = scoreParts[0] > scoreParts[1] ? p1 : p2;
    const winnerId   = nameToId[winnerName.toLowerCase()];
    if (!winnerId) {
      Logger.log(`Skipping row ${i+2}: Could not find winner ID for ${winnerName}`);
      return;
    }

    // Fetch live match from Challonge
    const matchUrl = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/matches/${matchId}.json?api_key=${API_KEY}`;
    const matchRes = UrlFetchApp.fetch(matchUrl);
    const matchObj = JSON.parse(matchRes.getContentText()).match;

    const liveP1Id = matchObj.player1_id;
    const liveP2Id = matchObj.player2_id;

    if (![liveP1Id, liveP2Id].includes(winnerId)) {
      Logger.log(`Skipping row ${i+2}: Winner ID ${winnerId} is not in live match participants`);
      return;
    }

    // Report to Challonge
    internal.reportMatchResult(matchId, scoresCsv, winnerId);

    // Update sheet
    sheet.getRange(i + 2, csvIdx + 1).setValue(scoresCsv); // update CSV column J
    sheet.getRange(i + 2, winnerIdx + 1).setValue(winnerName);
    sheet.getRange(i + 2, statusIdx + 1).setValue("Reported");

    Logger.log(`Reported match ${matchId}: Winner ${winnerName}, Scores CSV ${scoresCsv}`);
  });

  internal.mirrorScoresToOutput();
  checkRoundComplete();
}

function reportAllMatches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Matches");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const matchIdIdx  = headers.indexOf("Match ID");
  const set1Idx     = headers.indexOf("Set 1"); // Column F
  const set2Idx     = headers.indexOf("Set 2"); // Column G
  const set3Idx     = headers.indexOf("Set 3"); // Column H
  const matchScoreIdx = headers.indexOf("Match Score"); // Column I
  const csvIdx      = headers.indexOf("Score (CSV)"); // Column J
  const winnerIdx   = headers.indexOf("Winner");
  const p1Idx       = headers.indexOf("Player 1");
  const p2Idx       = headers.indexOf("Player 2");
  const statusIdx   = headers.indexOf("Status");

  if ([matchIdIdx, set1Idx, set2Idx, set3Idx, matchScoreIdx, csvIdx, winnerIdx, p1Idx, p2Idx, statusIdx].includes(-1)) {
    throw new Error("Missing required columns in 'All Matches'");
  }

  const { nameToId } = internal.fetchParticipants();

  data.forEach((row, i) => {
    const matchId    = row[matchIdIdx];
    const set1       = row[set1Idx];
    const set2       = row[set2Idx];
    const set3       = row[set3Idx];
    const matchScore = row[matchScoreIdx];
    const p1         = row[p1Idx];
    const p2         = row[p2Idx];
    const status     = row[statusIdx];

    // Only send if at least one set score exists
    const sets = [set1, set2, set3].filter(s => s && String(s).includes("-"));
    if (sets.length === 0) return;
    //if (status=="Reported") return;

    // Build scores_csv
    const scoresCsv = sets.join(",");

    // Determine winner
    const scoreParts = matchScore ? matchScore.split("-").map(Number) : null;
    if (!scoreParts || scoreParts.length !== 2 || scoreParts.some(isNaN)) {
      Logger.log(`Skipping row ${i+2}: Invalid match score "${matchScore}"`);
      return;
    }
    const winnerName = scoreParts[0] > scoreParts[1] ? p1 : p2;
    const winnerId   = nameToId[winnerName.toLowerCase()];
    if (!winnerId) {
      Logger.log(`Skipping row ${i+2}: Could not find winner ID for ${winnerName}`);
      return;
    }

    // Fetch live match from Challonge
    const matchUrl = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/matches/${matchId}.json?api_key=${API_KEY}`;
    const matchRes = UrlFetchApp.fetch(matchUrl);
    const matchObj = JSON.parse(matchRes.getContentText()).match;

    const liveP1Id = matchObj.player1_id;
    const liveP2Id = matchObj.player2_id;

    if (![liveP1Id, liveP2Id].includes(winnerId)) {
      Logger.log(`Skipping row ${i+2}: Winner ID ${winnerId} is not in live match participants`);
      return;
    }

    // Report to Challonge
    internal.reportMatchResult(matchId, scoresCsv, winnerId);

    // Update sheet
    sheet.getRange(i + 2, csvIdx + 1).setValue(scoresCsv); // update CSV column J
    sheet.getRange(i + 2, winnerIdx + 1).setValue(winnerName);
    sheet.getRange(i + 2, statusIdx + 1).setValue("Reported");

    Logger.log(`Reported match ${matchId}: Winner ${winnerName}, Scores CSV ${scoresCsv}`);
  });

  internal.mirrorScoresToOutput();
  checkRoundComplete();
}

function generateStandings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matchesSheet = ss.getSheetByName("All Matches");
  const standingsSheet = ss.getSheetByName("Current Standings");
  const historySheet = ss.getSheetByName("All Standings");

  const matchData = matchesSheet.getDataRange().getValues();
  const headers = matchData.shift();

  // const idxP1ID = headers.indexOf("Player1_ID");
  // const idxP2ID = headers.indexOf("Player2_ID");
  const idxP1 = headers.indexOf("Player 1");
  const idxP2 = headers.indexOf("Player 2");
  const idxS1 = headers.indexOf("Set 1");
  const idxS2 = headers.indexOf("Set 2");
  const idxS3 = headers.indexOf("Set 3");
  const idxMatchScore = headers.indexOf("Match Score");

  let players = {};

  // Initialize stats function
  function initPlayer(name) {
    if(!players[name]){
      players[name] = {
        name,
        wins: 0,
        tb1: 0,
        tb2: 0,
        tb3: 0,
        tb4: 0,
        tb5: 0,
        points: 0
      };
    }
  }

  // Process matches
  matchData.forEach(row => {
    const p1name = row[idxP1];
    const p2name = row[idxP2];
    const set1 = row[idxS1];
    const set2 = row[idxS2];
    const set3 = row[idxS3];
    const matchScore = row[idxMatchScore];

    initPlayer(p1name);
    initPlayer(p2name);

    // Helper to get set wins and score diff
    function parseSet(set) {
      const [a, b] = set.split("-").map(n => parseInt(n, 10) || 0);
      return { a, b };
    }

    const sets = [set1, set2, set3].map(parseSet);
    const match = parseSet(matchScore);

    const p1SetWins = sets.filter(s => s.a > s.b).length;
    const p2SetWins = sets.filter(s => s.b > s.a).length;
    const p1ScoreDiff = sets.reduce((sum, s) => sum + (s.a - s.b), 0);
    const p2ScoreDiff = -p1ScoreDiff;

    // Update total set and score diffs (TB4 & TB5)
    players[p1name].tb4 += p1SetWins - p2SetWins;
    players[p2name].tb4 += p2SetWins - p1SetWins;
    players[p1name].tb5 += p1ScoreDiff;
    players[p2name].tb5 += p2ScoreDiff;

    // Determine match winner
    if (match.a > match.b) {
      players[p1name].wins++;
    } else if (match.b > match.a) {
      players[p2name].wins++;
    }
  });

  // Calculate TB1–TB3 (vs tied players)
  const winGroups = {};
  for (let name in players) {
    const wins = players[name].wins;
    if (!winGroups[wins]) winGroups[wins] = [];
    winGroups[wins].push(name);
  }

  matchData.forEach(row => {
    const p1name = row[idxP1];
    const p2name = row[idxP2];
    const set1 = row[idxS1];
    const set2 = row[idxS2];
    const set3 = row[idxS3];
    const matchScore = row[idxMatchScore];

    const sameGroup = players[p1name].wins === players[p2name].wins;
    if (!sameGroup) return;

    function parseSet(set) {
      const [a, b] = set.split("-").map(n => parseInt(n, 10) || 0);
      return { a, b };
    }

    const sets = [set1, set2, set3].map(parseSet);
    const match = parseSet(matchScore);

    const p1SetWins = sets.filter(s => s.a > s.b).length;
    const p2SetWins = sets.filter(s => s.b > s.a).length;
    const p1ScoreDiff = sets.reduce((sum, s) => sum + (s.a - s.b), 0);
    const p2ScoreDiff = -p1ScoreDiff;

    if (match.a > match.b) players[p1name].tb1++;
    if (match.b > match.a) players[p2name].tb1++;

    players[p1name].tb2 += p1SetWins - p2SetWins;
    players[p2name].tb2 += p2SetWins - p1SetWins;

    players[p1name].tb3 += p1ScoreDiff;
    players[p2name].tb3 += p2ScoreDiff;
  });

  Object.values(players).forEach(p => {
    p.points = (p.wins * 100000) +
               (p.tb1 * 10000) +
               (p.tb2 * 1000) +
               (p.tb3 * 100) +
               (p.tb4 * 10) +
               (p.tb5);
  });

  let sorted = Object.values(players).sort((a, b) => b.points - a.points);

  // Write to Standings sheet
  standingsSheet.clear();
  standingsSheet.appendRow(["Rank", "Player", "Wins", "TB1", "TB2", "TB3", "TB4", "TB5", "Points"]);
  sorted.forEach((p, i) => {
    standingsSheet.appendRow([i + 1, p.name, p.wins, p.tb1, p.tb2, p.tb3, p.tb4, p.tb5, p.points]);
  });

  // Archive snapshot
  const timestamp = new Date();
  sorted.forEach((p, i) => {
    historySheet.appendRow([timestamp, i + 1, p.name, p.wins, p.tb1, p.tb2, p.tb3, p.tb4, p.tb5, p.points]);
  });
}

function checkRoundComplete() {
  Logger.log("Checking round...");
  const props = PropertiesService.getScriptProperties();

  const matches = internal.fetchMatches();
  if (matches.length === 0) {
    Logger.log("No matches found — tournament might not have started yet.");
    return;
  }

  // --- First complete check ---
  let allComplete = matches.every(m => m.state === 'complete');
  if (allComplete) {
    Logger.log("All matches currently marked complete — waiting 5 seconds to confirm...");
    Utilities.sleep(5000); // wait 5 seconds

    // Re-fetch to confirm
    const matchesRetry = internal.fetchMatches();
    allComplete = matchesRetry.every(m => m.state === 'complete');

    if (allComplete) {
      Logger.log("Tournament is confirmed complete!");

      if (props.getProperty("tournamentComplete") === "true") {
        Logger.log("Tournament already finalized.");
        return;
      }
      props.setProperty("tournamentComplete", "true");
      generateStandings();
      Logger.log("Final standings generated and report refreshed. Tournament complete.");
      return;
    } else {
      Logger.log("False alarm — new matches detected after wait.");
    }
  }

  // Group matches by round
  const rounds = [...new Set(matches.map(m => m.round))].sort((a, b) => a - b);

  let lastCompleteRound = null;
  for (let r of rounds) {
    const roundMatches = matches.filter(m => m.round === r);
    const roundComplete = roundMatches.every(m => m.state === 'complete');
    if (roundComplete) {
      lastCompleteRound = r;
    } else {
      break;
    }
  }

  if (!lastCompleteRound) {
    Logger.log("No rounds are fully complete yet.");
    return;
  }

  // Prevent re-running multiple times for same round
  const lastProcessed = Number(props.getProperty("lastProcessedRound"));
  if (lastProcessed === lastCompleteRound) {
    Logger.log(`Round ${lastCompleteRound} already processed.`);
    return;
  }

  props.setProperty("lastProcessedRound", String(lastCompleteRound));

  Logger.log(`Round ${lastCompleteRound} is complete — running updates...`);
  generateStandings();
  internal.writeStandingsToOutput(lastCompleteRound);
  refreshPairings();
  updateInput();
}

function startTournament() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/start.json?api_key=${API_KEY}`;

  const options = {
    method: 'post',
    muteHttpExceptions: true
  };
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedRound");
  PropertiesService.getScriptProperties().setProperty("tournamentComplete", "false");

  // const playersSheet = ss.getSheetByName("Player Names");
  
  // const playerNames = playersSheet.getRange(3, 3, playersSheet.getLastRow()).getValues()
  //   .flat()
  //   .filter(name => name); // remove empty cells

  // playerNames.forEach(name => {
  //   const addUrl = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/participants.json?api_key=${API_KEY}`;
  //   const payload = {
  //     "participant[name]": name
  //   };
  //   const addOptions = {
  //     method: 'post',
  //     payload: payload,
  //     contentType: 'application/x-www-form-urlencoded',
  //     muteHttpExceptions: true
  //   };
  //   const res = UrlFetchApp.fetch(addUrl, addOptions);
  //   Logger.log(`Added player '${name}': ${res.getContentText()}`);
  // });

  UrlFetchApp.fetch(url, options);

  const standingsSheet = ss.getSheetByName("Current Standings");
  const historySheet = ss.getSheetByName("All Standings");

  if (standingsSheet) standingsSheet.clearContents();
  if (historySheet) historySheet.clearContents();

  Logger.log("Standings sheets cleared at the start of the tournament.");
  internal.clearInput();
  internal.clearOutput();
  refreshPairings();
  updateInput();
}

function update_droplist(round){
 const form = FormApp.openById("1DhzuG7Ff3cjS0pks6_3T44SpMcdfpbtPl5BHX8cK1s0");
 var matches=parseInt(players/2); //12 matches each round, for section with 24 players;

 //Logger.log(form.getTitle());
 //var round=1
 var tab="R"+round;
 //const sheet = SpreadsheetApp.openById("1HinYpvFedp1EWwEY3R1tDlULEymTztDkGocDsQx2828").getSheetByName(tab);
 const sheet = SpreadsheetApp.openById("1uHcSdtAnLIuu2syecYPbwTWxUsHXmVRUBpWsSza-zFc").getSheetByName(tab);

 var values=sheet.getRange(2,15,matches,1).getValues();
  Logger.log(values);
  var items=form.getItems();
  /*Logger.log(items[0].getId().toString());
  Logger.log(items[1].getId().toString());
  */
  
  
  var item=form.getItemById(items[0].getId().toString());
     
  var item=form.getItemById(items[1].getId().toString());
   //Logger.log(item.asTextItem().getTitle());
   
   Logger.log(item.asListItem().getChoices());
   item.asListItem().setTitle("Round"+round+" match")
   item.asListItem().setChoiceValues(values);
}

const internal = {
  reportMatchResult : function(matchId, scoresCsv, winnerId) {
    const url = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/matches/${matchId}.json?api_key=${API_KEY}`;

    const formData = {
      "match[scores_csv]": scoresCsv,
      "match[winner_id]": String(winnerId)
    };

    const options = {
      method: 'put',
      payload: formData,
      contentType: 'application/x-www-form-urlencoded',
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`Response for match ${matchId}: ${response.getContentText()}`);
  },

  calculatePointsBeforeRound : function(allMatches, playerName, roundNumber) {
    let wins = 0;
    for (let r of allMatches) {
      if (typeof r[1] !== "number" || r[1] >= roundNumber) continue;
      const matchScore = r[8]; 
      if (!matchScore) continue;
      const [s1, s2] = matchScore.toString().split('-').map(Number);
      if (r[2] === playerName && s1 > s2) wins++;
      if (r[3] === playerName && s2 > s1) wins++;
    }
    return wins;
  },

  fetchParticipants : function() {
    const url = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/participants.json?api_key=${API_KEY}`;
    const res = UrlFetchApp.fetch(url);
    const list = JSON.parse(res.getContentText());
    const idToName = {};
    const nameToId = {};
    list.forEach(p => {
      idToName[p.participant.id] = p.participant.name;
      nameToId[p.participant.name.toLowerCase()] = p.participant.id;
    });
    return { idToName, nameToId };
  },

  fetchMatches : function() {
    const url = `https://api.challonge.com/v1/tournaments/${TOURNAMENT_URL}/matches.json?api_key=${API_KEY}`;
    const res = UrlFetchApp.fetch(url);
    return JSON.parse(res.getContentText()).map(m => m.match);
  },

  getCurrentRound : function(matches) {
    const open = matches.filter(m => m.state === "open");
    return open.length ? Math.min(...open.map(m => m.round)) : null;
  },

  clearInput : function() {
    const targetSS = SpreadsheetApp.openById(INPUT_SPREADSHEET_ID); 
    const sheets = targetSS.getSheets(); 
    sheets.forEach(
      sheet => { 
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          // Clear from row 2 down, cols B–L
          sheet.getRange(2, 2, lastRow - 1, 11).clearContent();
        }
      }
    ); 
    Logger.log('All '+sheets.length+' sheets cleared in the input spreadsheet.'); 
  },

  clearOutput : function() {
    const targetSS = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID); 
    const sheets = targetSS.getSheets(); 
    sheets.forEach(
      sheet => { 
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.clearContents();
        }
      }
    ); 
    Logger.log('All '+sheets.length+' sheets cleared in the output spreadsheet.'); 
  },

  writePairingsToOutput : function(latestRound, matches) {
    const targetSS = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID);
    let targetSheet = targetSS.getSheetByName("R" + latestRound) || targetSS.insertSheet("R" + latestRound);
    targetSheet.clearContents();

    // Title row
    targetSheet.getRange(1, 1).setValue("Round " + latestRound + " Pairings");

    // Headers
    targetSheet.getRange(2, 1, 1, 6).setValues([
      ["Table","Player 1","Points","Score","Points","Player 2"]
    ]);

    // Data
    const rows = matches.map(m => ['',m[1], m[2], m[10], m[5], m[4]]);
    if (rows.length) {
      targetSheet.getRange(3, 1, rows.length, rows[0].length).setValues(rows);
    }
  },

  writeStandingsToOutput : function(roundNumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const standingsSheet = ss.getSheetByName("Current Standings");
    const standings = standingsSheet.getDataRange().getValues();
    if (standings.length <= 1) return;

    const targetSS = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID);
    let targetSheet = targetSS.getSheetByName("R" + roundNumber) || targetSS.insertSheet("R" + roundNumber);

    // Clear old standings area (columns H onward)
    const lastRow = targetSheet.getMaxRows();
    const lastCol = targetSheet.getMaxColumns();
    if (lastCol > 7) {
      targetSheet.getRange(1, 8, lastRow, lastCol - 7).clear();
    }

    // Title
    targetSheet.getRange(1, 8).setValue("Standings after Round " + roundNumber);

    // Extract only: Rank, Player, Wins → as Points
    const headerIdx = {};
    standings[0].forEach((h, i) => headerIdx[h] = i);
    const rankIdx = headerIdx["Rank"];
    const playerIdx = headerIdx["Player"];
    const winsIdx = headerIdx["Wins"];

    const slim = [["Rank","Player","Points"]];
    for (let i = 1; i < standings.length; i++) {
      const row = standings[i];
      slim.push([row[rankIdx], row[playerIdx], row[winsIdx]]);
    }

    // Write slim standings starting at col H row 2
    targetSheet.getRange(2, 8, slim.length, slim[0].length).setValues(slim);
  },

  mirrorScoresToOutput : function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allMatchesSheet = ss.getSheetByName("All Matches");
    const data = allMatchesSheet.getDataRange().getValues();
    const headers = data.shift();

    const roundIdx  = headers.indexOf("Round");
    const p1Idx     = headers.indexOf("Player 1");
    const p2Idx     = headers.indexOf("Player 2");
    const scoreIdx  = headers.indexOf("Match Score");
    const winnerIdx = headers.indexOf("Winner");

    const targetSS = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID);

    data.forEach(row => {
      const round = row[roundIdx];
      const p1 = row[p1Idx], p2 = row[p2Idx];
      const score = row[scoreIdx], winner = row[winnerIdx];

      const sheet = targetSS.getSheetByName("R" + round);
      if (!sheet) return;

      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === p1 && values[i][5] === p2) {
          sheet.getRange(i + 1, 4).setValue(score);
        }
      }
    });
  }

}
