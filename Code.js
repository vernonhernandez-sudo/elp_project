const SHEET_ID = '1eysLnBQszwqm7g9B23mPlHco1rB2NpCd33dACk6OoCw';

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
        if (row[8] && row[8].toString() !== 'Active') return { success: false, message: 'Account is not active.' };
        var user = { id: row[0], username: row[1], name: row[3], email: row[4] };
        var sessionId = Utilities.getUuid();
        CacheService.getScriptCache().put(sessionId, JSON.stringify(user), 21600);
        return { success: true, sessionId: sessionId, user: { name: user.name, email: user.email } };
      }
    }
    return { success: false, message: 'Invalid username or password.' };
  } catch (error) { return { success: false, message: 'Error: ' + error.toString() }; }
}

function getUserFromCache(sessionId) {
  if (!sessionId) return null;
  try { var data = CacheService.getScriptCache().get(sessionId); return data ? JSON.parse(data) : null; } catch (e) { return null; }
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
    var newEntryId = data.length;
    sheet.appendRow([
      newEntryId, 
      entryData.authorName, 
      entryData.phones || '', 
      entryData.email || '', 
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
    for(var i=1;i<d.length;i++){
      if(d[i][0]==entryId){
        s.getRange(i+1,2).setValue(entryData.authorName);
        s.getRange(i+1,3).setValue(entryData.phones);
        s.getRange(i+1,4).setValue(entryData.email);
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
    var validStatuses=['','Pipe','Sold','DNC','VM']; 
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