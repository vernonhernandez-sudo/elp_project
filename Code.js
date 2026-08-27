const SHEET_ID = '1dPNpl5gPc4avLQ8PNqqCzyxE9L448HzJfr71mnQv_RM';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('ELP').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ========== SIMPLE LOGIN (No 2FA, No Roles) ==========
function loginUser(username, password) {
  if (!username || !password) return { success: false, message: 'Username and password required.' };
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents');
    var data = sheet.getDataRange().getValues(); data.shift();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var storedUsername = row[1] ? row[1].toString().trim().toLowerCase() : '';
      var storedPassword = row[2] ? row[2].toString().trim() : '';
      if (storedUsername === username.toLowerCase() && storedPassword === password) {
        var accountStatus = (row[8] || 'Active').toString().trim().toLowerCase();
        if (accountStatus !== 'active') {
          var statusMessage = accountStatus === 'suspended' ? 'Account is suspended.' : 'Account is inactive.';
          return { success: false, message: statusMessage + ' Please contact an administrator.' };
        }
        var user = { id: row[0], username: row[1], name: row[3], email: row[4], role: row[7] };
        var sessionId = Utilities.getUuid();
        CacheService.getScriptCache().put(sessionId, JSON.stringify(user), 21600);
        return { success: true, sessionId: sessionId, user: { name: user.name, email: user.email, role: user.role } };
      }
    }
    return { success: false, message: 'Invalid username or password.' };
  } catch (error) { return { success: false, message: 'Error: ' + error.toString() }; }
}

function getUserFromCache(sessionId) {
  if (!sessionId) return null;
  try {
    var data = CacheService.getScriptCache().get(sessionId);
    if (!data) return null;
    var user = JSON.parse(data);
    if (!user.role && user.id !== undefined && user.id !== null && user.id !== '') {
      var agents = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues();
      for (var i = 1; i < agents.length; i++) {
        if (String(agents[i][0]) === String(user.id)) {
          user.role = agents[i][7] || '';
          CacheService.getScriptCache().put(sessionId, JSON.stringify(user), 21600);
          break;
        }
      }
    }
    return user;
  } catch (e) { return null; }
}

function logoutUser(sessionId) {
  if (sessionId) try { CacheService.getScriptCache().remove(sessionId); } catch (e) {}
  return { success: true };
}

// ========== ENTRY MANAGEMENT (Simplified) ==========
function addEntry(sessionId, entryData) {
  try {
    var user = getUserFromCache(sessionId);
    if (!user) return { success: false, message: 'Session expired.' };
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries');
    if (!entryData.authorName || !entryData.book || !entryData.isbn) {
      return { success: false, message: 'Author Name, Book, and ISBN are required.' };
    }
    var data = sheet.getDataRange().getValues();
    var email = (entryData.email || '').toString().trim().toLowerCase();
    if (email) {
      for (var i = 1; i < data.length; i++) {
        if ((data[i][3] || '').toString().trim().toLowerCase() === email) {
          return { success: false, message: 'Email already exists.' };
        }
      }
    }
    var newEntryId = data.length;
    sheet.appendRow([
      newEntryId, 
      entryData.authorName, 
      entryData.phones || '', 
      email,
      entryData.book, 
      entryData.isbn, 
      entryData.address || '', 
      '', 
      '', 
      new Date().toISOString(), 
      'No', 
      'No'
    ]);
    SpreadsheetApp.flush();
    logActivity(user.id, user.name, 'Created entry #' + newEntryId + ': ' + entryData.authorName);
    // Auto-distribute to available agent
    distributeNewEntry(newEntryId);
    return { success: true, message: 'Entry added!', entryId: newEntryId };
  } catch (e) { return { success: false, message: 'Error: ' + e.toString() }; }
}

function updateEntry(sessionId, entryId, entryData) {
  try { 
    var u=getUserFromCache(sessionId); 
    if(!u)return{success:false,message:'Session expired.'}; 
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries'); 
    var d=s.getDataRange().getValues(); 
    var email=(entryData.email||'').toString().trim().toLowerCase();
    if(email){
      for(var j=1;j<d.length;j++){
        if(d[j][0]!=entryId&&(d[j][3]||'').toString().trim().toLowerCase()===email){
          return{success:false,message:'Email already exists.'};
        }
      }
    }
    for(var i=1;i<d.length;i++){
      if(d[i][0]==entryId){
        s.getRange(i+1,2).setValue(entryData.authorName);
        s.getRange(i+1,3).setValue(entryData.phones);
        s.getRange(i+1,4).setValue(email);
        s.getRange(i+1,5).setValue(entryData.book);
        s.getRange(i+1,6).setValue(entryData.isbn);
        s.getRange(i+1,7).setValue(entryData.address||'');
        logActivity(u.id,u.name,'Updated entry #'+entryId);
        return{success:true,message:'Updated!'};
      }
    } 
    return{success:false,message:'Not found.'}; 
  } catch(e){return{success:false,message:'Error'};}
}

function deleteEntry(sessionId, entryId) {
  try {
    var u=getUserFromCache(sessionId); 
    if(!u)return{success:false,message:'Session expired.'};
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries'); 
    var d=s.getDataRange().getValues();
    for(var k=d.length-1;k>=1;k--){
      if(d[k][0]==entryId){
        s.deleteRow(k+1);
        logActivity(u.id,u.name,'Deleted entry #'+entryId);
        return{success:true,message:'Entry deleted.'};
      }
    }
    return{success:false,message:'Not found.'};
  } catch(e){return{success:false,message:'Error: '+e.toString()};}
}

function buildEntrySummary(entries) {
  var summary = { totalAssigned: entries.length, pipeCount: 0, soldCount: 0, pendingCount: 0, vmCount: 0, dncCount: 0, exclusiveCount: 0, deleteCount: 0 };
  for (var i = 0; i < entries.length; i++) {
    if (entries[i].status === 'Pipe') summary.pipeCount++;
    if (entries[i].status === 'Sold') summary.soldCount++;
    if (entries[i].status === 'Pending') summary.pendingCount++;
    if (entries[i].status === 'VM') summary.vmCount++;
    if (entries[i].status === 'DNC') summary.dncCount++;
    if (entries[i].status === 'Exclusive') summary.exclusiveCount++;
    if (entries[i].status === 'Delete') summary.deleteCount++;
  }
  return summary;
}

function getEntriesForAgent(sessionId) { try {
    var u = getUserFromCache(sessionId);
    if (!u) { return { success: false, entries: [], summary: { pipeCount: 0, soldCount: 0, totalAssigned: 0 } }; }
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var entriesSheet = ss.getSheetByName('Entries');
    var ed = entriesSheet.getDataRange().getValues();
    if (ed.length <= 1) { return { success: true, entries: [], summary: { pipeCount: 0, soldCount: 0, totalAssigned: 0 } }; }
    ed.shift();
    var agentsSheet = ss.getSheetByName('Agents');
    var ad = agentsSheet.getDataRange().getValues();
    if (ad.length > 0) { ad.shift(); }
    var am = {};
    for (var k = 0; k < ad.length; k++) { var agentId = ad[k][0]; var agentName = ad[k][3]; if (agentId !== '' && agentId != null) { am[String(agentId)] = agentName || ''; } }
    var loggedInUserId = u.id;

    if (loggedInUserId === undefined || loggedInUserId === null || loggedInUserId === '') { return { success: false, entries: [], summary: { pipeCount: 0, soldCount: 0, totalAssigned: 0 } }; }

    loggedInUserId = String(loggedInUserId);
    var entries = [];
    for (var i = 0; i < ed.length; i++) { var assignedAgentId = ed[i][7]; if ( assignedAgentId === '' || assignedAgentId === null || assignedAgentId === undefined ) { continue; } if (String(assignedAgentId) !== loggedInUserId) { continue; } var assignedAgentName = am[String(assignedAgentId)] || 'Unassigned'; entries.push({ id: ed[i][0], authorName: ed[i][1] || '', phones: ed[i][2] || '', email: ed[i][3] || '', book: ed[i][4] || '', isbn: ed[i][5] || '', address: ed[i][6] || '', assignedAgentId: assignedAgentId, assignedAgentName: assignedAgentName, status: ed[i][8] || '', createdAt: ed[i][9] || '' }); }

    var pipeCount = 0;
    var soldCount = 0;

    for (var j = 0; j < entries.length; j++) { if (entries[j].status === 'Pipe') { pipeCount++; } if (entries[j].status === 'Sold') { soldCount++; } }
    return { success: true, entries: entries, summary: buildEntrySummary(entries) };
  } catch (e) { Logger.log('getEntries error: ' + e); return { success: false, entries: [], summary: { pipeCount: 0, soldCount: 0, totalAssigned: 0 }, error: e.toString() }; } }


// Everyone sees all entries (master list)
function getEntries(sessionId) {
  try {
    var u=getUserFromCache(sessionId);
    if(!u)return{success:false,entries:[],summary:{pipeCount:0,soldCount:0,totalAssigned:0}};
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var ed=ss.getSheetByName('Entries').getDataRange().getValues();
    if(ed.length<=1)return{success:true,entries:[],summary:{pipeCount:0,soldCount:0,totalAssigned:0}};
    ed.shift();
    var ad=ss.getSheetByName('Agents').getDataRange().getValues();ad.shift();
    var am={};
    for(var k=0;k<ad.length;k++){am[ad[k][0]]=ad[k][3];}
    var entries=[];
    for(var i=0;i<ed.length;i++){
      entries.push({
        id:ed[i][0],
        authorName:ed[i][1]||'',
        phones:ed[i][2]||'',
        email:ed[i][3]||'',
        book:ed[i][4]||'',
        isbn:ed[i][5]||'',
        address:ed[i][6]||'',
        assignedAgentId:ed[i][7],
        assignedAgentName:am[ed[i][7]]||'Unassigned',
        status:ed[i][8]||'',
        createdAt:ed[i][9]||''
      });
    }
    return{
      success:true,
      entries:entries,
      summary:{
        pipeCount:entries.filter(function(e){return e.status==='Pipe';}).length,
        soldCount:entries.filter(function(e){return e.status==='Sold';}).length,
        pendingCount:entries.filter(function(e){return e.status==='Pending';}).length,
        vmCount:entries.filter(function(e){return e.status==='VM';}).length,
        dncCount:entries.filter(function(e){return e.status==='DNC';}).length,
        exclusiveCount:entries.filter(function(e){return e.status==='Exclusive';}).length,
        deleteCount:entries.filter(function(e){return e.status==='Delete';}).length,
        totalAssigned:entries.length
      }
    };
  } catch(e){return{success:false,entries:[],summary:{pipeCount:0,soldCount:0,totalAssigned:0}};}
}

// ========== STATUS MANAGEMENT ==========
function updateEntryStatus(sessionId, entryId, newStatus) {
  try { 
    var u=getUserFromCache(sessionId); 
    if(!u)return{success:false,message:'Session expired.'}; 
    var validStatuses=['','Pipe','Sold','Pending','DNC','VM','Exclusive','Delete'];
    if(validStatuses.indexOf(newStatus)===-1)return{success:false,message:'Invalid status.'}; 
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries'); 
    var d=s.getDataRange().getValues(); 
    for(var i=1;i<d.length;i++){
      if(d[i][0]==entryId){
        s.getRange(i+1,9).setValue(newStatus);
        logActivity(u.id,u.name,'Status #'+entryId+' to '+(newStatus||'None'));
        addSystemRemark(entryId,u.name,'Status changed to '+(newStatus||'None'));
        return{success:true,message:'Updated!'};
      }
    } 
    return{success:false,message:'Not found.'}; 
  } catch(e){return{success:false,message:'Error'};}
}

// ========== REMARKS ==========
function addRemark(sessionId, entryId, remarkText) {
  try { 
    var u=getUserFromCache(sessionId); 
    if(!u||!remarkText||!remarkText.trim())return{success:false,message:'Invalid.'}; 
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks'); 
    s.appendRow([s.getDataRange().getValues().length,entryId,new Date().toISOString(),u.id,u.name,remarkText.trim()]); 
    logActivity(u.id,u.name,'Added remark to #'+entryId); 
    return{success:true,message:'Remark added!'}; 
  } catch(e){return{success:false,message:'Error'};}
}

function getRemarks(entryId) {
  try { 
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks'); 
    var d=s.getDataRange().getValues();d.shift();
    var r=[];
    for(var i=0;i<d.length;i++){
      if(d[i][1]==entryId)r.push({
        id:d[i][0],entryId:d[i][1],timestamp:d[i][2],
        userId:d[i][3],userName:d[i][4],remark:d[i][5]
      });
    }
    r.sort(function(a,b){return new Date(b.timestamp)-new Date(a.timestamp);});
    return r; 
  } catch(e){return[];}
}

// ========== ACTIVITY ==========
function getEntryActivity(entryId) {
  try { 
    var rs=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks').getDataRange().getValues();rs.shift(); 
    var as=SpreadsheetApp.openById(SHEET_ID).getSheetByName('ActivityLog').getDataRange().getValues();as.shift(); 
    var items=[]; 
    for(var i=0;i<rs.length;i++){
      if(rs[i][1]==entryId)items.push({type:'remark',timestamp:rs[i][2],userName:rs[i][4],content:rs[i][5]});
    } 
    for(var j=0;j<as.length;j++){
      if(as[j][3]&&as[j][3].toString().indexOf('#'+entryId)>-1)items.push({type:'activity',timestamp:as[j][4],userName:as[j][2],content:as[j][3]});
    } 
    items.sort(function(a,b){return new Date(b.timestamp)-new Date(a.timestamp);}); 
    return items.slice(0,50); 
  } catch(e){return[];}
}

function logActivity(uid,un,action) { 
  try { 
    SpreadsheetApp.openById(SHEET_ID).getSheetByName('ActivityLog').appendRow([
      SpreadsheetApp.openById(SHEET_ID).getSheetByName('ActivityLog').getDataRange().getValues().length,
      uid,un,action,new Date().toISOString()
    ]); 
  } catch(e){} 
}

function addSystemRemark(eid,un,msg) { 
  try { 
    SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks').appendRow([
      SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks').getDataRange().getValues().length,
      eid,new Date().toISOString(),0,un,msg
    ]); 
  } catch(e){} 
}

// ========== SIMPLE DISTRIBUTION ==========
function distributeNewEntry(entryId) {
  try {
    var as=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues();as.shift();
    var agents=[]; 
    for(var i=0;i<as.length;i++){
      if(as[i][8]==='Active' && as[i][7]!=='Super Admin') {
        agents.push(as[i]);
      }
    }
    if(agents.length === 0) return;
    var sa=agents[Math.floor(Math.random()*agents.length)];
    var es=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries'); 
    var ed=es.getDataRange().getValues();
    for(var i=1;i<ed.length;i++){
      if(ed[i][0]==entryId){
        es.getRange(i+1,8).setValue(sa[0]);
        es.getRange(i+1,10).setValue(new Date().toISOString());
        break;
      }
    }
    logActivity(0,'System','Auto-assigned #'+entryId+' to '+sa[3]); 
    addSystemRemark(entryId,'System','Lead auto-assigned to '+sa[3]);
  } catch(e){}
}

function getAgentName(id) { 
  try { 
    var d=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues(); 
    for(var i=1;i<d.length;i++){if(d[i][0]==id)return d[i][3];} 
    return 'Unknown'; 
  } catch(e){return'Unknown';} 
}

function getTransferRequestNotifications(sessionId) {
  try {
    var viewer = getUserFromCache(sessionId);
    if (!viewer) return { success: false, count: 0, requests: [], message: 'Session expired.' };
    var viewerRole = (viewer.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (viewerRole !== 'superadmin' && viewerRole !== 'admin') return { success: false, count: 0, requests: [], message: 'Only an Admin or Super Admin can view transfer requests.' };
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TransferRequests');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, count: 0, requests: [] };
    var values = sheet.getDataRange().getValues();
    var headers = values.shift().map(function(header) { return String(header || '').trim().toLowerCase(); });
    var agents = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues();
    var agentNames = {};
    for (var i = 1; i < agents.length; i++) agentNames[String(agents[i][0])] = agents[i][3] || '';
    var nameIndex = -1;
    for (var h = 0; h < headers.length; h++) {
      if (/from.*(agent|user).*name|request(ed)?\s*by.*name/.test(headers[h])) { nameIndex = h; break; }
    }
    var idIndex = -1;
    for (var j = 0; j < headers.length; j++) {
      if (/from.*(id|agent)|requester.*id/.test(headers[j])) { idIndex = j; break; }
    }
    var toAgentNameIndex = headers.indexOf('toagentname');
    if (toAgentNameIndex === -1 && values.length && values[0].length > 14) toAgentNameIndex = 14;
    var statusIndex = -1;
    for (var s = 0; s < headers.length; s++) {
      if (/status|decision|approval|review/.test(headers[s])) { statusIndex = s; break; }
    }
    var requests = [];
    values.forEach(function(row, rowIndex) {
      if (!row.some(function(value) { return value !== '' && value !== null; })) return;
      if (statusIndex > -1 && /^(approve|approved|disapprove|disapproved|rejected)$/i.test(String(row[statusIndex] || '').trim())) return;
      var agentName = nameIndex > -1 ? String(row[nameIndex] || '').trim() : '';
      if (!agentName && idIndex > -1) agentName = agentNames[String(row[idIndex])] || '';
      requests.push({
        rowNumber: rowIndex + 2,
        agentName: agentName || 'Unknown Agent',
        toAgentName: toAgentNameIndex > -1 ? String(row[toAgentNameIndex] || '').trim() : ''
      });
    });
    return { success: true, count: requests.length, requests: requests };
  } catch (e) { return { success: false, count: 0, requests: [], message: 'Unable to load transfer requests.' }; }
}

function reviewTransferRequest(sessionId, rowNumber, decision, note) {
  try {
    var reviewer = getUserFromCache(sessionId);
    if (!reviewer) return { success: false, message: 'Session expired.' };
    var reviewerRole = (reviewer.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (reviewerRole !== 'superadmin' && reviewerRole !== 'admin') return { success: false, message: 'Only an Admin or Super Admin can review transfer requests.' };
    if (decision !== 'Approve' && decision !== 'Disapprove') return { success: false, message: 'Choose Approve or Disapprove.' };
    var transferSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TransferRequests');
    var row = Number(rowNumber);
    if (!transferSheet || !row || row < 2 || row > transferSheet.getLastRow()) return { success: false, message: 'Transfer request not found.' };
    var headers = transferSheet.getRange(1, 1, 1, transferSheet.getLastColumn()).getValues()[0];
    var statusColumn = 0;
    for (var i = 0; i < headers.length; i++) {
      if (/status|decision|approval|review/.test((headers[i] || '').toString().toLowerCase())) { statusColumn = i + 1; break; }
    }
    if (!statusColumn) statusColumn = transferSheet.getLastColumn() + 1;
    if (statusColumn > transferSheet.getLastColumn()) transferSheet.insertColumnAfter(transferSheet.getLastColumn());
    transferSheet.getRange(1, statusColumn).setValue('Review Status');
    transferSheet.getRange(row, statusColumn).setValue(decision);

    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var noteSheet = spreadsheet.getSheetByName('reviewNote');
    if (!noteSheet) {
      var allSheets = spreadsheet.getSheets();
      for (var n = 0; n < allSheets.length; n++) {
        if (allSheets[n].getName().toLowerCase() === 'reviewnote') { noteSheet = allSheets[n]; break; }
      }
    }
    if (!noteSheet) noteSheet = spreadsheet.insertSheet('reviewNote');
    var noteText = (note || '').toString().trim();
    var noteHeaders = noteSheet.getLastRow() > 0 ? noteSheet.getRange(1, 1, 1, noteSheet.getLastColumn()).getValues()[0] : [];
    if (!noteHeaders.length) {
      noteHeaders = ['Reviewed At', 'Transfer Request Row', 'Decision', 'Note', 'Reviewer ID', 'Reviewer Name'];
      noteSheet.getRange(1, 1, 1, noteHeaders.length).setValues([noteHeaders]);
    }
    var noteColumn = -1;
    for (var h = 0; h < noteHeaders.length; h++) {
      if (/note|comment|remark/.test((noteHeaders[h] || '').toString().toLowerCase())) { noteColumn = h; break; }
    }
    if (noteColumn === -1) {
      noteColumn = noteHeaders.length;
      noteSheet.getRange(1, noteColumn + 1).setValue('Note');
      noteHeaders.push('Note');
    }
    var reviewRow = new Array(noteHeaders.length).fill('');
    for (var c = 0; c < noteHeaders.length; c++) {
      var noteHeader = (noteHeaders[c] || '').toString().toLowerCase();
      if (/date|time|reviewed|created/.test(noteHeader)) reviewRow[c] = new Date().toISOString();
      else if (/decision|status|approval/.test(noteHeader)) reviewRow[c] = decision;
      else if (/reviewer.*name|user.*name/.test(noteHeader)) reviewRow[c] = reviewer.name || '';
      else if (/reviewer.*id|user.*id/.test(noteHeader)) reviewRow[c] = reviewer.id || '';
      else if (/request|row|id/.test(noteHeader)) reviewRow[c] = row;
    }
    reviewRow[noteColumn] = noteText;
    noteSheet.getRange(noteSheet.getLastRow() + 1, 1, 1, reviewRow.length).setValues([reviewRow]);
    SpreadsheetApp.flush();
    logActivity(reviewer.id, reviewer.name, decision + ' transfer request on row ' + row);
    return { success: true, message: 'Transfer request reviewed.' };
  } catch (e) { return { success: false, message: 'Unable to review transfer request: ' + e.toString() }; }
}

function getTeamColumn(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var header = (headers[i] || '').toString().trim().toLowerCase().replace(/[\s_/-]+/g, '');
    if (header === 'team' || header === 'position' || header === 'teamposition') return i + 1;
  }
  return 10;
}

function addUser(sessionId, userData) {
  try {
    var creator = getUserFromCache(sessionId);
    if (!creator) return { success: false, message: 'Session expired.' };
    var creatorRole = (creator.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (creatorRole !== 'superadmin' && creatorRole !== 'admin') return { success: false, message: 'Only an Admin or Super Admin can add users.' };
    if (!userData || !userData.username || !userData.password || !userData.name || !userData.email) {
      return { success: false, message: 'Username, password, name, and email are required.' };
    }
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents');
    var data = sheet.getDataRange().getValues();
    var username = userData.username.toString().trim().toLowerCase();
    var email = userData.email.toString().trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][1] || '').toString().trim().toLowerCase() === username) return { success: false, message: 'Username already exists.' };
      if ((data[i][4] || '').toString().trim().toLowerCase() === email) return { success: false, message: 'Email already exists.' };
    }
    var newUserId = data.length;
    var role = creatorRole === 'admin' ? 'Agent' : (userData.role || 'Agent');
    var position = role === 'Super Admin' ? '' : (userData.position || 'BM');
    sheet.appendRow([newUserId, username, userData.password.toString(), userData.name.toString().trim(), email, '', '', role, userData.status || 'Active']);
    sheet.getRange(sheet.getLastRow(), getTeamColumn(sheet)).setValue(position);
    SpreadsheetApp.flush();
    logActivity(creator.id, creator.name, 'Created user #' + newUserId + ': ' + username);
    return { success: true, message: 'User added.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.toString() }; }
}

function getUsers(sessionId) {
  try {
    var viewer = getUserFromCache(sessionId);
    if (!viewer) return { success: false, users: [], message: 'Session expired.' };
    var role = (viewer.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (role !== 'superadmin' && role !== 'admin') return { success: false, users: [], message: 'Only an Admin or Super Admin can view users.' };
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents');
    var data = sheet.getDataRange().getValues();
    var teamColumn = getTeamColumn(sheet);
    var users = [];
    for (var i = 1; i < data.length; i++) {
      users.push({
        id: data[i][0],
        username: data[i][1] || '',
        name: data[i][3] || '',
        email: data[i][4] || '',
        ip: data[i][6] || '',
        role: data[i][7] || '',
        status: data[i][8] || '',
        position: data[i][teamColumn - 1] || ''
      });
    }
    return { success: true, users: users };
  } catch (e) { return { success: false, users: [], message: 'Error: ' + e.toString() }; }
}

function updateUser(sessionId, userId, userData) {
  try {
    var editor = getUserFromCache(sessionId);
    if (!editor) return { success: false, message: 'Session expired.' };
    var editorRole = (editor.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (editorRole !== 'superadmin' && editorRole !== 'admin') return { success: false, message: 'Only an Admin or Super Admin can update users.' };
    if (!userData || !userData.username || !userData.name || !userData.email) return { success: false, message: 'Username, name, and email are required.' };
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents');
    var data = sheet.getDataRange().getValues();
    var teamColumn = getTeamColumn(sheet);
    var username = userData.username.toString().trim().toLowerCase();
    var email = userData.email.toString().trim().toLowerCase();
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] != userId && (data[j][1] || '').toString().trim().toLowerCase() === username) return { success: false, message: 'Username already exists.' };
      if (data[j][0] != userId && (data[j][4] || '').toString().trim().toLowerCase() === email) return { success: false, message: 'Email already exists.' };
    }
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == userId) {
        sheet.getRange(i + 1, 2).setValue(username);
        if (userData.password) sheet.getRange(i + 1, 3).setValue(userData.password.toString());
        sheet.getRange(i + 1, 4).setValue(userData.name.toString().trim());
        sheet.getRange(i + 1, 5).setValue(email);
        sheet.getRange(i + 1, 7).setValue(userData.ip || '');
        var role = editorRole === 'admin' ? 'Agent' : (userData.role || 'Agent');
        sheet.getRange(i + 1, 8).setValue(role);
        sheet.getRange(i + 1, 9).setValue(userData.status || 'Active');
        sheet.getRange(i + 1, teamColumn).setValue(role === 'Super Admin' ? '' : (userData.position || 'BM'));
        logActivity(editor.id, editor.name, 'Updated user #' + userId + ': ' + username);
        return { success: true, message: 'User updated.' };
      }
    }
    return { success: false, message: 'User not found.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.toString() }; }
}

function deleteUser(sessionId, userId) {
  try {
    var editor = getUserFromCache(sessionId);
    if (!editor) return { success: false, message: 'Session expired.' };
    var editorRole = (editor.role || '').toString().toLowerCase().replace(/\s+/g, '');
    if (editorRole !== 'superadmin') return { success: false, message: 'Only a Super Admin can delete users.' };
    if (String(editor.id) === String(userId)) return { success: false, message: 'You cannot delete your own account.' };
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents');
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(userId)) {
        var deletedUsername = data[i][1] || data[i][3] || userId;
        sheet.deleteRow(i + 1);
        logActivity(editor.id, editor.name, 'Deleted user #' + userId + ': ' + deletedUsername);
        return { success: true, message: 'User deleted.' };
      }
    }
    return { success: false, message: 'User not found.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.toString() }; }
}