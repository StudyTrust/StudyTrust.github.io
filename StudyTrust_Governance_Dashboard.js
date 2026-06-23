'use strict';

// ════════════════════════════════════════════
// CONFIGURATION
// ════════════════════════════════════════════
const CLIENT_ID    = 'c5c93f6e-957d-431f-8e83-ff22e7be10a9';
const REDIRECT_URI = 'https://studytrust.github.io/auth.html';
const SP_HOST      = 'studietrust.sharepoint.com';
const SP_PATH      = '/sites/StudyTrustAuditTracker';
const GRAPH        = 'https://graph.microsoft.com/v1.0';

const LISTS = { pol:'GovPolicies', sop:'GovSOPs', log:'GovChangelog', trustees:'GovTrustees', ack:'GovAcknowledgements' };

const MSAL_CFG = {
  auth:{ clientId:CLIENT_ID, authority:'https://login.microsoftonline.com/studietrust.onmicrosoft.com', redirectUri:REDIRECT_URI },
  cache:{ cacheLocation:'sessionStorage', storeAuthStateInCookie:true }
};
const SCOPES = ['Sites.ReadWrite.All','Sites.Manage.All','User.Read','GroupMember.Read.All'];

// Status transitions
const STATUS_FLOW = {
  Initiated:   ['Draft','Archived'],          // idea registered, not yet drafting
  Draft:       ['Consult','Archived'],
  Consult:     ['PTT Review','Draft','Archived'],
  'PTT Review':['Approved','Consult','Draft'],
  Approved:    ['Published','Draft'],
  New:         ['Published','Draft'],
  Published:   ['Review Due','Archived'],
  'Review Due':['Draft','Archived'],
  Archived:    []
};
const DRAFT_STATUSES     = ['Initiated','Draft','Consult','PTT Review'];
const PUBLISHED_STATUSES = ['Approved','New','Published'];

// ════════════════════════════════════════════
// MSAL + GRAPH API
// ════════════════════════════════════════════════
// GOVERNANCE DOCUMENT UPLOAD
// ════════════════════════════════════════════════

let _govDriveCache = {};

async function getGovDriveId(libraryName) {
  if (_govDriveCache[libraryName]) return _govDriveCache[libraryName];
  const sid = await getSiteId();
  const data = await gFetch(`/sites/${sid}/drives`);
  const norm = s => s.toLowerCase().replace(/[_\s\-]+/g,' ').trim();
  const drive = (data.value||[]).find(d =>
    d.name === libraryName ||
    norm(d.name||'') === norm(libraryName) ||
    norm((d.webUrl||'').split('/').pop()) === norm(libraryName)
  );
  if (!drive) {
    const avail = (data.value||[]).map(d=>d.name).join(', ');
    throw new Error(`Library "${libraryName}" not found. Available: ${avail}`);
  }
  _govDriveCache[libraryName] = drive.id;
  return drive.id;
}

function govDocFolder(docType, docStatus, category) {
  const base = docType === 'pol' ? '01_Policies' : '02_SOPs';
  const cat  = (category||'General').replace(/[^a-zA-Z0-9 ]/g,'').trim().replace(/\s+/g,'_');
  const year = new Date().getFullYear();
  switch(docStatus) {
    case 'Initiated':
    case 'Draft': case 'Consult': case 'PTT Review':
      return `${base}/01_Drafts/${cat}`;
    case 'Approved':
      return `${base}/03_Approved_Pending_Publication`;
    case 'Published': case 'Review Due':
      return `${base}/04_Published/${cat}`;
    case 'Archived':
      return `${base}/05_Archived/${year}`;
    default:
      return `${base}/01_Drafts/${cat}`;
  }
}

async function uploadGovFile(libraryName, folderPath, file) {
  const tok = await getToken();
  const driveId = await getGovDriveId(libraryName);
  const sid = await getSiteId();
  const encodedPath = (folderPath+'/'+file.name).split('/').map(encodeURIComponent).join('/');
  const url = `${GRAPH}/sites/${sid}/drives/${driveId}/root:/${encodedPath}:/content`;
  const r = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: 'Bearer '+tok, 'Content-Type': file.type||'application/octet-stream' },
    body: await file.arrayBuffer()
  });
  if (!r.ok) { const e = await r.json().catch(()=>({})); throw new Error(e?.error?.message||'HTTP '+r.status); }
  return await r.json();
}

async function createGovFolder(libraryName, parentPath, folderName) {
  const tok = await getToken();
  const driveId = await getGovDriveId(libraryName);
  const sid = await getSiteId();
  const encodedParent = parentPath.split('/').map(encodeURIComponent).join('/');
  const url = `${GRAPH}/sites/${sid}/drives/${driveId}/root:/${encodedParent}:/children`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { Authorization: 'Bearer '+tok, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' })
  });
  if (!r.ok) { const e = await r.json().catch(()=>({})); throw new Error(e?.error?.message||'HTTP '+r.status); }
  return await r.json();
}

function triggerGovUpload(fieldId) {
  const inp = document.getElementById('gov-file-input');
  inp.dataset.field = fieldId; inp.click();
}
function triggerGovMultiUpload(fieldId) {
  const inp = document.getElementById('gov-multi-input');
  inp.dataset.field = fieldId; inp.click();
}

async function handleGovFileUpload(input) {
  const file = input.files[0]; if (!file) return;
  const fieldId = input.dataset.field;
  const {id, type} = urlTarget||{};
  const arr = type === 'pol' ? polData : sopData;
  const doc = arr ? arr.find(d => (type==='pol' ? d.polId : d.sopId) === id) : null;
  if (!doc) { toast('Open the document URL editor first.','err'); input.value=''; return; }
  let status = doc.status;
  if (fieldId === 'f-url-pub') status = 'Published';
  if (fieldId === 'f-url-arch') status = 'Archived';
  const folder = govDocFolder(type, status, doc.category);
  toast('Uploading ' + file.name + '…', 'ok');
  try {
    const result = await uploadGovFile('Governance Published', folder, file);
    document.getElementById(fieldId).value = result.webUrl||'';
    toast('Uploaded: ' + file.name, 'ok');
  } catch(e) { toast('Upload failed: ' + e.message, 'err'); }
  input.value = '';
}

async function handleGovMultiUpload(input) {
  const files = Array.from(input.files); if (!files.length) return;
  const fieldId = input.dataset.field;
  const {id, type} = urlTarget||{};
  const arr = type === 'pol' ? polData : sopData;
  const doc = arr ? arr.find(d => (type==='pol' ? d.polId : d.sopId) === id) : null;
  if (!doc) { toast('Open the document URL editor first.','err'); input.value=''; return; }
  let status = doc.status;
  if (fieldId === 'f-url-pub') status = 'Published';
  if (fieldId === 'f-url-arch') status = 'Archived';
  const parentFolder = govDocFolder(type, status, doc.category);
  const datePart = new Date().toISOString().slice(0,10);
  const titlePart = (doc.title||'Documents').replace(/[^a-zA-Z0-9 ]/g,'').trim().split(/\s+/).slice(0,4).join('_');
  const folderName = datePart + '_' + titlePart;
  toast(`Creating folder and uploading ${files.length} file${files.length!==1?'s':''}…`, 'ok');
  try {
    const folderResult = await createGovFolder('Governance Published', parentFolder, folderName);
    const folderPath = parentFolder + '/' + folderResult.name;
    for (const file of files) { await uploadGovFile('Governance Published', folderPath, file); }
    document.getElementById(fieldId).value = folderResult.webUrl||'';
    toast(`Folder created with ${files.length} file${files.length!==1?'s':''}: ${folderResult.name}`, 'ok');
  } catch(e) { toast('Upload failed: ' + e.message, 'err'); }
  input.value = '';
}

function clearPolFilters(){
  ['pol-q','pol-cat','pol-tier','pol-status','pol-risk','pol-approver','pol-custodian','pol-owner','pol-rev-from','pol-rev-to']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  renderPol();
}
function clearSopFilters(){
  ['sop-q','sop-cat','sop-status','sop-custodian','sop-owner','sop-rev-from','sop-rev-to']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  renderSop();
}

// ════════════════════════════════════════════════
// GRAPH API
// ════════════════════════════════════════════════
let msalApp, graphToken, graphSiteId;

async function initMsal(){
  msalApp = new msal.PublicClientApplication(MSAL_CFG);
  await msalApp.initialize();
  let redir = null;
  try{
    redir = await msalApp.handleRedirectPromise();
  }catch(e){
    try{ msalApp.clearCache(); }catch(_){}
    try{ sessionStorage.clear(); }catch(_){}
  }
  if(redir) graphToken = redir.accessToken;
  const accs = msalApp.getAllAccounts();
  if(accs.length){
    try{
      const r = await msalApp.acquireTokenSilent({scopes:SCOPES,account:accs[0]});
      graphToken = r.accessToken;
      document.getElementById('user-name').textContent = accs[0].name || accs[0].username;
      return true;
    }catch(e){graphToken=null;}
  }
  return false;
}

async function doSignIn(){
  const btn=document.getElementById('si-btn'), err=document.getElementById('si-err');
  btn.disabled=true; btn.textContent='Signing in…'; err.classList.remove('show');
  try{
    const result = await msalApp.loginPopup({
      scopes: SCOPES,
      redirectUri: REDIRECT_URI
    });
    graphToken = result.accessToken;
    document.getElementById('user-name').textContent = result.account.name || result.account.username;
    document.getElementById('signin-screen').classList.add('hidden');
    await loadAndRender();
  }catch(e){
    btn.disabled=false;
    btn.innerHTML='<svg width="18" height="18" viewBox="0 0 21 21" fill="none"><rect width="10" height="10" fill="#F25022"/><rect x="11" width="10" height="10" fill="#7FBA00"/><rect y="11" width="10" height="10" fill="#00A4EF"/><rect x="11" y="11" width="10" height="10" fill="#FFB900"/></svg> Sign in with Microsoft';
    let msg = e.message||'Sign in failed.';
    err.textContent=msg; err.classList.add('show');
  }
}

async function getToken(){
  if(graphToken) return graphToken;
  const accs=msalApp.getAllAccounts();
  if(!accs.length) throw new Error('Not signed in.');
  const r=await msalApp.acquireTokenSilent({scopes:SCOPES,account:accs[0]});
  graphToken=r.accessToken; return graphToken;
}

async function gFetch(url,opts={}){
  const tok=await getToken();
  const full=url.startsWith('http')?url:GRAPH+url;
  const r=await fetch(full,{...opts,headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json',Accept:'application/json',...(opts.headers||{})}});
  if(!r.ok&&r.status!==204){const e=await r.json().catch(()=>({}));throw new Error(e?.error?.message||'Graph API error HTTP '+r.status);}
  return r.status===204?null:r.json();
}

async function getSiteId(){
  if(graphSiteId) return graphSiteId;
  const d=await gFetch('/sites/'+SP_HOST+':'+SP_PATH);
  graphSiteId=d.id; return graphSiteId;
}

async function spGet(listName){
  const sid=await getSiteId();
  let url=`/sites/${sid}/lists/${listName}/items?$expand=fields&$top=5000`;
  let items=[];
  while(url){const d=await gFetch(url);items=items.concat(d.value||[]);url=d['@odata.nextLink']||null;}
  return items;
}

async function spPost(listName,fields){
  const sid=await getSiteId();
  return gFetch(`/sites/${sid}/lists/${listName}/items`,{method:'POST',body:JSON.stringify({fields})});
}

async function spPatch(listName,itemId,fields){
  const sid=await getSiteId();
  await gFetch(`/sites/${sid}/lists/${listName}/items/${itemId}`,{method:'PATCH',body:JSON.stringify({fields})});
}

// Columns that may not yet exist in SharePoint — gracefully skip if they cause errors.
// Simon must create these columns to enable full functionality (see Admin Manual Section 2.2).
const SP_OPTIONAL_COLS = {
  sop: [],   // Custodian + Notes added to GovSOPs (2026-06)
  pol: []    // Notes added to GovPolicies (2026-06)
};

async function spPatchResilient(listName, itemId, fields, docType) {
  try {
    await spPatch(listName, itemId, fields);
  } catch(e) {
    // Only retry for column-not-found errors — re-throw everything else
    const msg = (e.message||'').toLowerCase();
    const msg2 = e.message||'';
    const isColErr = /column|field|property|not recognized|does not exist|invalid property/i.test(msg2);
    const isInvalidReq = /invalid request|invalid value|bad request/i.test(msg2);
    if (!isColErr && !isInvalidReq) throw e;
    // If 'invalid request', convert any object-type URL values to plain strings and retry
    if (isInvalidReq && !isColErr) {
      const fixedUrl = Object.fromEntries(Object.entries(fields).map(([k,v])=>[
        k, (v && typeof v==='object' && 'Url' in v) ? (v.Url||null) : v
      ]));
      await spPatch(listName, itemId, fixedUrl);
      return;
    }
    // Strip the optional columns that don't exist yet
    const optional = SP_OPTIONAL_COLS[docType] || [];
    if (!optional.length) throw e;
    const stripped = Object.fromEntries(
      Object.entries(fields).filter(([k]) => !optional.includes(k))
    );
    await spPatch(listName, itemId, stripped);  // retry without optional cols
    // Warn which columns need to be created by Simon
    const missing = optional.filter(c => c in fields && fields[c] !== null && fields[c] !== undefined);
    if (missing.length) {
      toast(
        'Saved \u2713 (most fields) — but the '
        + missing.join(' and ')
        + ' column' + (missing.length > 1 ? 's do' : ' does')
        + ' not exist yet in SharePoint. Ask Simon to add '
        + (missing.length > 1 ? 'them' : 'it')
        + ' to the Gov' + (docType==='sop' ? 'SOPs' : 'Policies')
        + ' list as Single line of text.'
      , 'warning');
    }
  }
}

// ════════════════════════════════════════════
// FIELD MAPPINGS
// ════════════════════════════════════════════
const POL_MAP={title:'Title',polId:'PolicyID',category:'Category',tier:'Tier',version:'PolicyVersion',status:'Status',risk:'Risk',approver:'Approver',reviewDate:'ReviewDate',owner:'Owner',custodian:'Custodian',notes:'Notes',draftUrl:'DraftUrl',publishedUrl:'PublishedUrl',archiveUrl:'ArchiveUrl'};
const SOP_MAP={title:'Title',sopId:'SopID',category:'Category',linkedPolicy:'LinkedPolicy',version:'SopVersion',status:'Status',reviewDate:'ReviewDate',owner:'Owner',custodian:'Custodian',notes:'Notes',draftUrl:'DraftUrl',publishedUrl:'PublishedUrl',archiveUrl:'ArchiveUrl',sobCompliant:'SOBCompliant',sobReviewDate:'SOBReviewDate',sobReviewedBy:'SOBReviewedBy'};
const LOG_MAP={note:'Title',docId:'DocID',docType:'DocType',fromStatus:'FromStatus',toStatus:'ToStatus',changedBy:'ChangedBy',changeDate:'ChangeDate'};

// Hyperlink columns: SharePoint "Hyperlink or Picture" type returns/accepts {Url, Description} objects via Graph.
// Backward compatible: if your SP column is plain text these helpers pass strings through unchanged.
// URL columns are Single line of text in SharePoint — save as plain strings.
// fromSpRead handles both string and {Url,Description} object formats on reading.
const HYPERLINK_FIELDS = new Set(); // empty — URLs saved as plain strings
const DATE_FIELDS = new Set(['reviewDate','invited','confirmed','changeDate']);
function spWriteValue(field, value, item){
  if (HYPERLINK_FIELDS.has(field)) {
    if (!value) return null;
    return { Url: value, Description: (item && item.title) || value };
  }
  if (DATE_FIELDS.has(field)) return value ? value + 'T00:00:00Z' : null;
  return value || null;
}
function spReadValue(field, v){
  if (HYPERLINK_FIELDS.has(field) && v && typeof v === 'object') return v.Url || '';
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) return v.substring(0,10);
  return v == null ? '' : v;
}
// Build a SharePoint fields object for one URL field, respecting hyperlink shape
function spUrlField(field, value, item){
  return spWriteValue(field, value, item);
}

function toSpPol(p){const f={};for(const[ak,sk]of Object.entries(POL_MAP)){f[sk]=spWriteValue(ak,p[ak],p);}return f;}
function toSpSop(s){const f={};for(const[ak,sk]of Object.entries(SOP_MAP)){f[sk]=spWriteValue(ak,s[ak],s);}return f;}

function fromSpPol(sp){const f=sp.fields||{},p={_spId:sp.id};for(const[ak,sk]of Object.entries(POL_MAP)){p[ak]=spReadValue(ak,f[sk]);}return p;}
function fromSpSop(sp){const f=sp.fields||{},s={_spId:sp.id};for(const[ak,sk]of Object.entries(SOP_MAP)){s[ak]=spReadValue(ak,f[sk]);}return s;}
function fromSpLog(sp){const f=sp.fields||{},l={_spId:sp.id};for(const[ak,sk]of Object.entries(LOG_MAP)){let v=f[sk];if(typeof v==='string'&&/^\d{4}-\d{2}-\d{2}T/.test(v))v=v.substring(0,10);l[ak]=v||'';}return l;}

// ════════════════════════════════════════════
// STATE
// ════════════════════════════════════════════
let polData=[], sopData=[], changeLog=[], trusteesData=[];

const TR_MAP={name:'Title',email:'Email',title:'TrusteeTitle',invited:'DateInvited',confirmed:'DateConfirmed',status:'Status',notes:'Notes'};

function toSpTrustee(t){const f={};for(const[ak,sk]of Object.entries(TR_MAP)){const v=t[ak];f[sk]=(ak==='invited'||ak==='confirmed')?(v?v+'T00:00:00Z':null):(v||null);}return f;}
function fromSpTrustee(sp){const f=sp.fields||{},t={_spId:sp.id};for(const[ak,sk]of Object.entries(TR_MAP)){let v=f[sk];if(typeof v==='string'&&/^\d{4}-\d{2}-\d{2}T/.test(v))v=v.substring(0,10);t[ak]=v||'';}return t;}
let currentRole='personnel', userName='User';
let statusTarget=null, urlTarget=null;
let ADMIN_GROUP='Dashboards - Governance-Admins', TRUSTEE_GROUP='Dashboards - Board-of-Trustees', AUDITOR_GROUP='Dashboards - External-Auditors', PTT_GROUP='Dashboards - Policy Task Team', PUBLISHERS_GROUP='Dashboards - Policy-Publishers', LEGAL_REVIEWER_GROUP='Dashboards - Legal-Reviewer';

// Load group config from localStorage
function loadGroupConfig(){
  let cfg=JSON.parse(localStorage.getItem('st_gov_group_cfg')||'{}');
  // Migration: if stored config uses the old StudyTrust- prefix, discard it so
  // the new Dashboards- defaults take effect without manual intervention.
  if(cfg.admin && cfg.admin.toLowerCase().startsWith('studytrust-')){
    localStorage.removeItem('st_gov_group_cfg'); cfg={};
    console.info('Stale group config cleared — using new Dashboards- defaults.');
  }
  ADMIN_GROUP=cfg.admin||'Dashboards - Governance-Admins';
  TRUSTEE_GROUP=cfg.trustee||'Dashboards - Board-of-Trustees';
  AUDITOR_GROUP=cfg.auditor||'Dashboards - External-Auditors';
  PTT_GROUP=cfg.ptt||'Dashboards - Policy Task Team';
  PUBLISHERS_GROUP=cfg.publishers||'Dashboards - Policy-Publishers';
  LEGAL_REVIEWER_GROUP=cfg.legal||'Dashboards - Legal-Reviewer';
  const ag=document.getElementById('cfg-admin-grp'), tg=document.getElementById('cfg-trustee-grp'), aug=document.getElementById('cfg-auditor-grp'), pg=document.getElementById('cfg-ptt-grp');
  if(ag) ag.value=ADMIN_GROUP; if(tg) tg.value=TRUSTEE_GROUP; if(aug) aug.value=AUDITOR_GROUP; if(pg) pg.value=PTT_GROUP;
}
function saveGroupConfig(){
  const ag=document.getElementById('cfg-admin-grp').value.trim();
  const tg=document.getElementById('cfg-trustee-grp').value.trim();
  const aug=document.getElementById('cfg-auditor-grp').value.trim();
  const pg=(document.getElementById('cfg-ptt-grp')||{}).value?document.getElementById('cfg-ptt-grp').value.trim():PTT_GROUP;
  const pubg=(document.getElementById('cfg-publishers-grp')||{}).value?document.getElementById('cfg-publishers-grp').value.trim():PUBLISHERS_GROUP;
  const lg=(document.getElementById('cfg-legal-grp')||{}).value?document.getElementById('cfg-legal-grp').value.trim():LEGAL_REVIEWER_GROUP;
  localStorage.setItem('st_gov_group_cfg',JSON.stringify({admin:ag,trustee:tg,auditor:aug,ptt:pg,publishers:pubg,legal:lg}));
  ADMIN_GROUP=ag; TRUSTEE_GROUP=tg; AUDITOR_GROUP=aug; PTT_GROUP=pg; PUBLISHERS_GROUP=pubg; LEGAL_REVIEWER_GROUP=lg;
  toast('Group configuration saved.','success');
}
function resetGroupConfig(){
  if(!confirm('Reset all group names to the Dashboards- defaults?')) return;
  localStorage.removeItem('st_gov_group_cfg');
  ADMIN_GROUP='Dashboards - Governance-Admins';
  TRUSTEE_GROUP='Dashboards - Board-of-Trustees';
  AUDITOR_GROUP='Dashboards - External-Auditors';
  PTT_GROUP='Dashboards - Policy Task Team';
  PUBLISHERS_GROUP='Dashboards - Policy-Publishers';
  LEGAL_REVIEWER_GROUP='Dashboards - Legal-Reviewer';
  loadGroupConfig();
  toast('Group names reset to defaults. Sign out and back in to re-detect your role.','success');
}

// ════════════════════════════════════════════
// ROLE DETECTION via Graph API
// ════════════════════════════════════════════
// Stores diagnostic info for the help panel
let _diagGroups=[], _diagRole='', _diagError='';

async function detectRole(){
  try{
    const d=await gFetch('/me/memberOf?$select=displayName&$top=200');
    _diagGroups=(d.value||[]).map(g=>g.displayName||'').filter(Boolean);
    if(_diagGroups.some(g=>g.toLowerCase()===ADMIN_GROUP.toLowerCase())){ _diagRole='admin'; return 'admin'; }
    if(_diagGroups.some(g=>g.toLowerCase()===TRUSTEE_GROUP.toLowerCase())){ _diagRole='trustee'; return 'trustee'; }
    if(_diagGroups.some(g=>g.toLowerCase()===PTT_GROUP.toLowerCase())){ _diagRole='ptt'; return 'ptt'; }
    if(_diagGroups.some(g=>g.toLowerCase()===LEGAL_REVIEWER_GROUP.toLowerCase())){ _diagRole='legal'; return 'legal'; }
    if(_diagGroups.some(g=>g.toLowerCase()===AUDITOR_GROUP.toLowerCase())){ _diagRole='auditor'; return 'auditor'; }
    _diagRole='personnel';
  }catch(e){
    _diagError=e.message;
    _diagGroups=[];
    _diagRole='personnel (error)';
    console.warn('Role detection failed:',e);
  }
  return 'personnel';
}

// ════════════════════════════════════════════
// DEFAULT SEED DATA (from v3)
// ════════════════════════════════════════════
const DEFAULT_POL=[
  {polId:'POL-001',title:'Foundational Governance Policy',category:'Governance',tier:'T1',version:'1.2',status:'Published',risk:'High Strategic',approver:'Dashboards - Board-of-Trustees',reviewDate:'2027-05-01',owner:'National Director',custodian:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-002',title:'Code of Conduct & Ethics Policy',category:'Governance',tier:'T1',version:'1.1',status:'Published',risk:'High Strategic',approver:'Dashboards - Board-of-Trustees',reviewDate:'2027-05-01',owner:'National Director',custodian:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-003',title:'Human Resources Policy',category:'HR',tier:'T2',version:'2.0',status:'Published',risk:'High Operational',approver:'National Director',reviewDate:'2026-06-30',owner:'HR Manager',custodian:'HR Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-004',title:'Leave & Attendance Policy',category:'HR',tier:'T3',version:'1.0',status:'Published',risk:'Medium',approver:'National Director',reviewDate:'2026-12-31',owner:'HR Manager',custodian:'HR Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-005',title:'IT Security & Access Control Policy',category:'IT',tier:'T2',version:'1.1',status:'Published',risk:'High Operational',approver:'National Director',reviewDate:'2026-04-30',owner:'IT Lead',custodian:'IT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-006',title:'Data Protection & Privacy (POPIA)',category:'IT',tier:'T1',version:'1.0',status:'Published',risk:'High Strategic',approver:'Dashboards - Board-of-Trustees',reviewDate:'2026-07-01',owner:'Governance Lead',custodian:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-007',title:'Financial Management Policy',category:'Finance',tier:'T1',version:'2.1',status:'Review Due',risk:'High Strategic',approver:'Dashboards - Board-of-Trustees',reviewDate:'2026-03-31',owner:'CFO',custodian:'CFO',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-008',title:'Procurement & Supply Chain Policy',category:'Finance',tier:'T2',version:'1.0',status:'Published',risk:'High Operational',approver:'National Director',reviewDate:'2026-09-01',owner:'Manager: Operations',custodian:'Manager: Operations',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-009',title:'Bursary Management Policy',category:'Bursaries',tier:'T2',version:'1.1',status:'Published',risk:'High Operational',approver:'National Director',reviewDate:'2026-08-01',owner:'Bursary Manager',custodian:'Bursary Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-010',title:'Funder Relations Policy',category:'Funders',tier:'T2',version:'1.0',status:'Published',risk:'Medium',approver:'National Director',reviewDate:'2026-04-15',owner:'Manager: Operations',custodian:'Manager: Operations',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-011',title:'Risk Management Policy',category:'Operations',tier:'T2',version:'1.0',status:'Published',risk:'High Operational',approver:'National Director',reviewDate:'2026-10-15',owner:'Manager: Operations',custodian:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {polId:'POL-012',title:'Business Continuity Policy',category:'Operations',tier:'T2',version:'0.9',status:'Consult',risk:'High Operational',approver:'National Director',reviewDate:'2026-09-01',owner:'CEO',custodian:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''}
];
const DEFAULT_SOP=[
  {sopId:'SOP-001',title:'Policy Development Procedure',category:'Governance',linkedPolicy:'POL-001',status:'Published',version:'1.0',reviewDate:'2027-05-01',owner:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-002',title:'Board Meeting Procedure',category:'Governance',linkedPolicy:'POL-001',status:'Published',version:'1.1',reviewDate:'2027-05-01',owner:'National Director',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-003',title:'Conflict of Interest Declaration',category:'Governance',linkedPolicy:'POL-002',status:'Published',version:'1.0',reviewDate:'2027-05-01',owner:'PTT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-004',title:'Whistleblowing & Protected Disclosure',category:'Governance',linkedPolicy:'POL-002',status:'Published',version:'1.0',reviewDate:'2027-05-01',owner:'National Director',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-005',title:'Recruitment & Onboarding',category:'HR',linkedPolicy:'POL-003',status:'Review Due',version:'2.1',reviewDate:'2025-12-31',owner:'HR Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-006',title:'Leave Application Process',category:'HR',linkedPolicy:'POL-004',status:'Published',version:'1.0',reviewDate:'2026-12-31',owner:'HR Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-007',title:'IT Access Provisioning',category:'IT',linkedPolicy:'POL-005',status:'Published',version:'1.0',reviewDate:'2026-09-15',owner:'IT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-008',title:'Incident Response Procedure',category:'IT',linkedPolicy:'POL-005',status:'Review Due',version:'1.0',reviewDate:'2025-10-01',owner:'IT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-009',title:'Data Backup & Recovery',category:'IT',linkedPolicy:'POL-006',status:'Published',version:'1.0',reviewDate:'2026-07-01',owner:'IT Lead',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-010',title:'Payment Authorisation Process',category:'Finance',linkedPolicy:'POL-007',status:'Published',version:'1.1',reviewDate:'2026-03-31',owner:'CFO',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-011',title:'Budget Preparation Procedure',category:'Finance',linkedPolicy:'POL-007',status:'Published',version:'1.0',reviewDate:'2026-08-31',owner:'CFO',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-012',title:'Bursary Application & Selection',category:'Bursaries',linkedPolicy:'POL-009',status:'Published',version:'1.0',reviewDate:'2026-08-01',owner:'Bursary Manager',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-013',title:'Funder Reporting Process',category:'Funders',linkedPolicy:'POL-010',status:'Draft',version:'0.1',reviewDate:'2026-05-01',owner:'Manager: Operations',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-014',title:'Risk Assessment Procedure',category:'Operations',linkedPolicy:'POL-011',status:'Published',version:'1.0',reviewDate:'2026-10-15',owner:'Manager: Operations',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-015',title:'Business Continuity Procedure',category:'Operations',linkedPolicy:'POL-012',status:'Published',version:'1.0',reviewDate:'2026-09-01',owner:'CEO',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-016',title:'Emergency Decisions Procedure',category:'Governance',linkedPolicy:'POL-001',status:'PTT Review',version:'0.2',reviewDate:'2028-05-01',owner:'CEO',draftUrl:'',publishedUrl:'',archiveUrl:''},
  {sopId:'SOP-017',title:'Third-Party Due Diligence',category:'Governance',linkedPolicy:'POL-001',status:'Published',version:'1.0',reviewDate:'2028-05-01',owner:'Manager: Operations',draftUrl:'',publishedUrl:'',archiveUrl:''}
];

// ════════════════════════════════════════════
// LOAD DATA
// ════════════════════════════════════════════
async function loadData(){
  const [pRows,sRows,lRows,tRows]=await Promise.all([spGet(LISTS.pol),spGet(LISTS.sop),spGet(LISTS.log),spGet(LISTS.trustees).catch(()=>[])]);
  polData=pRows.map(fromSpPol);
  sopData=sRows.map(fromSpSop);
  changeLog=lRows.map(fromSpLog).sort((a,b)=>(b.changeDate||'').localeCompare(a.changeDate||''));
  trusteesData=tRows.map(fromSpTrustee);
  const sb=document.getElementById('seed-banner');
  if(polData.length===0&&sopData.length===0) sb.style.display='flex'; else sb.style.display='none';
}

// ════════════════════════════════════════════
// SEED DATA
// ════════════════════════════════════════════
async function seedData(){
  const total=DEFAULT_POL.length+DEFAULT_SOP.length+6;
  setLoad('Seeding '+total+' items into SharePoint…','Posting to SharePoint lists one by one…');
  const errors=[], fieldErrors=[];
  let seeded=0;

  function clean(fields){
    const out={};
    for(const[k,v]of Object.entries(fields)){
      if(v!==null&&v!==undefined&&v!=='') out[k]=v;
    }
    return out;
  }

  // Try posting a full payload. If it fails, fall back field by field,
  // posting only the fields SharePoint accepts, then report which failed.
  async function safePost(listName, fullFields, itemLabel){
    const cf=clean(fullFields);
    try{
      await spPost(listName,cf);
      seeded++;
      return true;
    }catch(e){
      // Full post failed — try fields one by one to identify the bad column
      // Post with just Title first to create the item
      const base={Title:cf.Title||itemLabel};
      let itemId=null;
      try{
        const r=await spPost(listName,base);
        itemId=r?.id;
        seeded++;
      }catch(e2){
        errors.push({item:itemLabel,list:listName,msg:'Even Title-only post failed: '+e2.message});
        return false;
      }
      // Now PATCH additional fields one at a time
      const sid=await getSiteId();
      const badFields=[];
      for(const[k,v]of Object.entries(cf)){
        if(k==='Title') continue;
        try{
          await gFetch(`/sites/${sid}/lists/${listName}/items/${itemId}`,{
            method:'PATCH',body:JSON.stringify({fields:{[k]:v}})
          });
        }catch(fe){
          badFields.push(k);
        }
      }
      if(badFields.length){
        fieldErrors.push({item:itemLabel,list:listName,fields:badFields});
      }
      return true;
    }
  }

  for(const p of DEFAULT_POL){
    await safePost(LISTS.pol, toSpPol({...p,title:p.title||p.polId}), p.polId||p.title);
  }
  for(const s of DEFAULT_SOP){
    await safePost(LISTS.sop, toSpSop({...s,title:s.title||s.sopId}), s.sopId||s.title);
  }

  const defaultTrustees=[
    {name:'Andre Bartlett',title:'Chairperson',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
    {name:'Sethlogane Manchidi',title:'Trustee',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
    {name:'Naledi Maricowitz',title:'Trustee',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
    {name:'Itumeleng Matlou',title:'Trustee',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
    {name:'Isaah Mhlanga',title:'Trustee',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
    {name:'Doc Sethole',title:'Trustee',email:'',invited:'',confirmed:'',status:'Pending Invitation',notes:''},
  ];
  for(const t of defaultTrustees){
    await safePost(LISTS.trustees, clean(toSpTrustee(t)), t.name);
  }

  await loadData();
  await loadAckData();
  initTemplateUrl(); renderAll(); hideLoad();

  // Remove old error panel
  const oldP=document.getElementById('seed-error-panel');
  if(oldP) oldP.remove();

  if(errors.length===0&&fieldErrors.length===0){
    toast('Seeded '+seeded+' items successfully.','success');
    return;
  }

  // Build error report
  let report='';
  if(errors.length) report+=errors.map(e=>`• ${e.list} — ${e.item}: ${e.msg}`).join('\n');
  if(fieldErrors.length){
    report+='\n\nColumns that do not exist in SharePoint (create these then seed again):\n';
    const allBad=[...new Set(fieldErrors.flatMap(fe=>fe.fields))];
    report+=allBad.map(f=>`  ✗ ${f}`).join('\n');
  }

  let panel=document.getElementById('seed-error-panel');
  if(!panel){
    panel=document.createElement('div');
    panel.id='seed-error-panel';
    panel.style.cssText='background:#fde8e8;border:1px solid #ef9a9a;border-radius:8px;padding:14px 16px;margin:0 20px 14px;font-size:12px;color:#7a2020';
    document.getElementById('content').insertBefore(panel,document.getElementById('content').firstChild);
  }
  panel.innerHTML=`<strong>Seed report — ${seeded} items created, ${fieldErrors.length} had missing columns:</strong>
    <pre style="margin-top:8px;white-space:pre-wrap;font-size:11px">${report}</pre>
    <div style="margin-top:8px;font-size:11px">Items were created in SharePoint with available fields. Create the missing columns listed above, then use Seed again — duplicates will not be created for existing items.</div>`;
  toast(seeded+' items seeded. Check the error panel for missing columns.','warning');
}

// ════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════
const td=()=>new Date().toISOString().slice(0,10);
const esc=t=>String(t||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const isOverdue=d=>d&&new Date(d)<new Date();
const isPublished=s=>['Approved','New','Published'].includes(s);
const isDraft=s=>['Initiated','Draft','Consult','PTT Review'].includes(s);

function sbadge(s){const m={Initiated:'bpr',Draft:'ba',Consult:'ba','PTT Review':'bt',Approved:'bt',New:'bb',Published:'bg','Review Due':'ba',Archived:'bgy'};return`<span class="bd ${m[s]||'bgy'}">${esc(s)}</span>`;}
function tbadge(t){return`<span class="bd bb">${esc(t)}</span>`;}
function rbadge(r){const c=r==='High Strategic'||r==='High Operational'?'br':r==='Medium'?'ba':'bg';return`<span class="bd ${c}">${esc(r)}</span>`;}
function datecell(d){if(!d)return'—';return isOverdue(d)?`<span class="overdue">${esc(d)} ⚠</span>`:esc(d);}

function docLinks(doc,type){
  const isA=currentRole==='admin', isP=currentRole==='ptt';
  const isPub=isPublished(doc.status), isDr=isDraft(doc.status), isAr=doc.status==='Archived';
  let h='<div class="doc-links">';
  if(isDr){
    if((isA||isP)&&doc.draftUrl) h+=`<a class="doc-btn edit" href="${esc(doc.draftUrl)}?web=1" target="_blank">✏ Edit Draft</a>`;
    else if(isA||isP) h+='<span style="font-size:11px;color:var(--tx3)">No draft URL</span>';
  }
  if(isPub){
    if(doc.publishedUrl) h+=`<a class="doc-btn view" href="${esc(doc.publishedUrl)}" target="_blank">📄 View</a>`;
    else h+='<span style="font-size:11px;color:var(--tx3)">No URL yet</span>';
    if(isA&&doc.publishedUrl) h+=`<a class="doc-btn edit" href="${esc(doc.publishedUrl)}?web=1" target="_blank">✏ Edit</a>`;
  }
  if(isAr&&isA){
    if(doc.archiveUrl) h+=`<a class="doc-btn arch" href="${esc(doc.archiveUrl)}" target="_blank">🗄 Archive</a>`;
    else h+='<span style="font-size:11px;color:var(--tx3)">No archive URL</span>';
  }
  const id=doc.polId||doc.sopId||doc._spId;
  // Admin: full URL editor. PTT: only draft URL editor, only on documents currently in a draft phase.
  if(isA) {
    h+=`<button class="doc-btn" style="background:var(--st-p);color:var(--st);border-color:var(--st-l)" onclick="openUrlModal('${esc(id)}','${type}')">⚙ URLs</button>`;
  } else if(isP && isDr) {
    h+=`<button class="doc-btn" style="background:var(--st-p);color:var(--st);border-color:var(--st-l)" onclick="openUrlModal('${esc(id)}','${type}')">✏ Draft URL</button>`;
  }
  h+='</div>'; return h;
}

// ════════════════════════════════════════════
// RENDER FUNCTIONS
// ════════════════════════════════════════════
function renderPol(){
  const q=(document.getElementById('pol-q')||{}).value||'';
  const cf=(document.getElementById('pol-cat')||{}).value||'';
  const tf=(document.getElementById('pol-tier')||{}).value||'';
  const sf=(document.getElementById('pol-status')||{}).value||'';
  const rf=(document.getElementById('pol-risk')||{}).value||'';
  const af=(document.getElementById('pol-approver')||{}).value||'';
  const cuf=(document.getElementById('pol-custodian')||{}).value||'';
  const owf=(document.getElementById('pol-owner')||{}).value||'';
  const prf=(document.getElementById('pol-rev-from')||{}).value||'';
  const prt=(document.getElementById('pol-rev-to')||{}).value||'';
  const ql=q.toLowerCase();

  // Admin and PTT see all statuses; trustees, auditors and personnel see only published
  const seesAll = currentRole==='admin' || currentRole==='ptt' || currentRole==='legal';
  const visibleData = seesAll ? polData : polData.filter(p=>isPublished(p.status));

  const tbody=document.getElementById('pol-body'); if(!tbody) return;
  let html='',shown=0;
  visibleData.forEach(p=>{
    const txt=[p.polId,p.title,p.category,p.tier,p.version,p.status,p.risk,p.approver,p.reviewDate,p.owner,p.custodian,p.notes].join(' ').toLowerCase();
    if(ql&&txt.indexOf(ql)===-1) return;
    if(cf&&p.category!==cf) return;
    if(tf&&p.tier!==tf) return;
    if(sf&&p.status!==sf) return;
    if(rf&&p.risk!==rf) return;
    if(af&&p.approver!==af) return;
    if(cuf&&p.custodian!==cuf) return;
    if(owf&&p.owner!==owf) return;
    if(prf&&p.reviewDate&&p.reviewDate<prf) return;
    if(prt&&p.reviewDate&&p.reviewDate>prt) return;
    shown++;
    const canChange = (currentRole==='admin') || (currentRole==='ptt' && allowedTransitions(p.status,'ptt').length>0) || (currentRole==='legal' && allowedTransitions(p.status,'legal').length>0);
    const act = canChange
      ? `<td onclick="event.stopPropagation()"><button class="btn btn-sec btn-sm" onclick="openStatusModal('${esc(p.polId)}','pol','${esc(p.status)}')">Change Status</button></td>`
      : (seesAll ? '<td></td>' : '');
    html+=`<tr class="reg-row" onclick="openEditModal('${esc(p.polId)}','pol')" style="cursor:pointer" title="Click to edit this policy"><td>${esc(p.polId)}</td><td>${esc(p.title)}</td><td>${esc(p.category)}</td><td>${tbadge(p.tier)}</td><td>${esc(p.version)}</td><td>${sbadge(p.status)}</td><td>${rbadge(p.risk)}</td><td style="font-size:11px">${esc(p.approver)}</td><td>${datecell(p.reviewDate)}</td><td style="font-size:11px">${esc(p.custodian||'')}</td><td style="font-size:11px">${esc(p.owner)}</td><td>${docLinks(p,'pol')}</td>${act}</tr>`;
  });
  if(!shown){
    if(polData.length===0){
      html=`<tr><td colspan="13" style="text-align:center;padding:24px">
        <div style="color:#856404;font-weight:600;margin-bottom:6px">No policies loaded from SharePoint.</div>
        <div style="color:#5a6172;font-size:11px">The GovPolicies list may be empty or the dashboard could not read it.<br>
        Check the ℹ Access Diagnostics panel → Test GovPolicies to verify the list connection.</div>
      </td></tr>`;
    } else {
      html=`<tr><td colspan="13" style="text-align:center;color:var(--tx3);padding:20px">No policies match the current filters. Clear all filters to see all ${polData.length} policies.</td></tr>`;
    }
  }
  const polTbl=document.querySelector('#sec-pol .tw');
  const polCg=document.getElementById('pol-card-grid');
  if(polView==='card'){
    if(polTbl) polTbl.style.display='none';
    if(polCg) polCg.style.display='grid';
    renderPolCards(visibleData);
  } else {
    if(polCg) polCg.style.display='none';
    if(polTbl) polTbl.style.display='';
    tbody.innerHTML=html;
  }

  const hdr=document.getElementById('pol-act-hdr');
  if(hdr) hdr.style.display=(currentRole==='admin'||currentRole==='ptt')?'':'none';
  const lbl=document.getElementById('pol-prog-lbl');
  const pub=polData.filter(p=>isPublished(p.status)).length;
  if(lbl) lbl.textContent=`${pub}/${polData.length} published`;

  // Populate category filter
  const cats=[...new Set(polData.map(p=>p.category).filter(Boolean))].sort();
  const cf2=document.getElementById('pol-cat');
  // Populate Approver, Custodian, Owner dropdowns
  ['pol-approver','pol-custodian','pol-owner'].forEach((id,i)=>{
    const key=['approver','custodian','owner'][i];
    const el=document.getElementById(id); if(!el||el.options.length>1) return;
    [...new Set(polData.map(p=>p[key]).filter(Boolean))].sort()
      .forEach(v=>{ const o=new Option(v,v); el.appendChild(o); });
  });
  if(cf2&&cf2.options.length<2) cats.forEach(c=>{const o=new Option(c,c);cf2.appendChild(o);});

  // Status filter
  const sf2=document.getElementById('pol-status');
  if(sf2&&sf2.options.length<2){
    const stats=[...new Set(polData.map(p=>p.status))];
    stats.forEach(s=>{const o=new Option(s,s);sf2.appendChild(o);});
  }
}

function renderSop(){
  const q=(document.getElementById('sop-q')||{}).value||'';
  const cf=(document.getElementById('sop-cat')||{}).value||'';
  const sf=(document.getElementById('sop-status')||{}).value||'';
  const scuf=(document.getElementById('sop-custodian')||{}).value||'';
  const sowf=(document.getElementById('sop-owner')||{}).value||'';
  const srf=(document.getElementById('sop-rev-from')||{}).value||'';
  const srt=(document.getElementById('sop-rev-to')||{}).value||'';
  const ql=q.toLowerCase();

  const seesAll = currentRole==='admin' || currentRole==='ptt' || currentRole==='legal';
  const visibleData = seesAll ? sopData : sopData.filter(s=>isPublished(s.status));

  const tbody=document.getElementById('sop-body'); if(!tbody) return;
  let html='',shown=0;
  visibleData.forEach(s=>{
    const txt=[s.sopId,s.title,s.category,s.linkedPolicy,s.status,s.version,s.reviewDate,s.owner,s.custodian,s.notes].join(' ').toLowerCase();
    if(ql&&txt.indexOf(ql)===-1) return;
    if(cf&&s.category!==cf) return;
    if(sf&&s.status!==sf) return;
    if(scuf&&s.custodian!==scuf) return;
    if(sowf&&s.owner!==sowf) return;
    if(srf&&s.reviewDate&&s.reviewDate<srf) return;
    if(srt&&s.reviewDate&&s.reviewDate>srt) return;
    shown++;
    const canChange = (currentRole==='admin') || (currentRole==='ptt' && allowedTransitions(s.status,'ptt').length>0) || (currentRole==='legal' && allowedTransitions(s.status,'legal').length>0);
    const act = canChange
      ? `<td onclick="event.stopPropagation()"><button class="btn btn-sec btn-sm" onclick="openStatusModal('${esc(s.sopId)}','sop','${esc(s.status)}')">Change Status</button></td>`
      : (seesAll ? '<td></td>' : '');
    html+=`<tr class="reg-row" onclick="openEditModal('${esc(s.sopId)}','sop')" style="cursor:pointer" title="Click to edit this SOP"><td>${esc(s.sopId)}</td><td>${esc(s.title)}</td><td>${esc(s.category)}</td><td>${esc(s.linkedPolicy)}</td><td>${sbadge(s.status)} ${sobBadge(s)}</td><td>${esc(s.version)}</td><td>${datecell(s.reviewDate)}</td><td style="font-size:11px">${esc(s.custodian||'')}</td><td style="font-size:11px">${esc(s.owner)}</td><td>${docLinks(s,'sop')}</td>${act}</tr>`;
  });
  if(!shown) html=`<tr><td colspan="11" style="text-align:center;color:var(--tx3);padding:20px">No SOPs match the current filters.</td></tr>`;
    const sopTbl=document.querySelector('#sec-sop .tw');
  const sopCg=document.getElementById('sop-card-grid');
  if(sopView==='card'){
    if(sopTbl) sopTbl.style.display='none';
    if(sopCg) sopCg.style.display='grid';
    renderSopCards(visibleData);
  } else {
    if(sopCg) sopCg.style.display='none';
    if(sopTbl) sopTbl.style.display='';
    tbody.innerHTML=html;
  }

  const hdr=document.getElementById('sop-act-hdr');
  if(hdr) hdr.style.display=(currentRole==='admin'||currentRole==='ptt')?'':'none';
  const lbl=document.getElementById('sop-prog-lbl');
  const pub=sopData.filter(s=>isPublished(s.status)).length;
  if(lbl) lbl.textContent=`${pub}/${sopData.length} published`;

  const cats=[...new Set(sopData.map(s=>s.category).filter(Boolean))].sort();
  const cf2=document.getElementById('sop-cat');
  ['sop-custodian','sop-owner'].forEach((id,i)=>{
    const key=['custodian','owner'][i];
    const el=document.getElementById(id); if(!el||el.options.length>1) return;
    [...new Set(sopData.map(s=>s[key]).filter(Boolean))].sort()
      .forEach(v=>{ const o=new Option(v,v); el.appendChild(o); });
  });
  if(cf2&&cf2.options.length<2) cats.forEach(c=>{const o=new Option(c,c);cf2.appendChild(o);});
  const sf2=document.getElementById('sop-status');
  if(sf2&&sf2.options.length<2){const stats=[...new Set(sopData.map(s=>s.status))];stats.forEach(s=>{const o=new Option(s,s);sf2.appendChild(o);});}
}

function renderExec(){
  const totalP=polData.length, pubP=polData.filter(p=>isPublished(p.status)).length;
  const totalS=sopData.length, pubS=sopData.filter(s=>isPublished(s.status)).length;
  const overP=polData.filter(p=>isOverdue(p.reviewDate)&&p.status!=='Archived').length;
  const overS=sopData.filter(s=>isOverdue(s.reviewDate)&&s.status!=='Archived').length;
  const hiR=polData.filter(p=>(p.risk==='High Strategic'||p.risk==='High Operational')&&p.status!=='Archived').length;
  const misM=polData.filter(p=>!p.owner||!p.reviewDate||!p.custodian).length;
  const comp=totalP?((totalP-misM)/totalP):1;
  const revC=totalP?Math.max(0,(totalP-overP)/totalP):1;
  const withSop=polData.filter(p=>sopData.some(s=>s.linkedPolicy===p.polId)).length;
  const sopL=totalP?(withSop/totalP):1;
  const gmi=Math.round(((comp*0.4)+(revC*0.4)+(sopL*0.2))*5*10)/10;
  const gmiLbls=['Poor','Developing','Functional','Strong','Mature'];
  const gmiLbl=gmiLbls[Math.min(4,Math.floor(gmi))];

  document.getElementById('kpi-grid').innerHTML=`
    <div class="kpi hl"><div class="kpi-l">Governance Maturity Index</div><div class="kpi-v" style="color:var(--st)">${gmi.toFixed(1)}/5</div><div class="kpi-s">${gmiLbl}</div></div>
    <div class="kpi"><div class="kpi-l">Total Policies</div><div class="kpi-v">${totalP}</div><div class="kpi-s">${pubP} published</div></div>
    <div class="kpi"><div class="kpi-l">Total SOPs</div><div class="kpi-v">${totalS}</div><div class="kpi-s">${pubS} published</div></div>
    <div class="kpi"><div class="kpi-l">Overdue Reviews</div><div class="kpi-v" style="color:${overP+overS>0?'var(--rd-tx)':'var(--gn-tx)'}">${overP+overS}</div><div class="kpi-s">${overP} policies · ${overS} SOPs</div></div>
    <div class="kpi"><div class="kpi-l">High-Risk Policies</div><div class="kpi-v" style="color:var(--am-tx)">${hiR}</div><div class="kpi-s">Annual review required</div></div>
    <div class="kpi"><div class="kpi-l">Policies with SOPs</div><div class="kpi-v" style="color:var(--gn-tx)">${withSop}</div><div class="kpi-s">of ${totalP} total</div></div>`;

  // Trustee summary
  const ts=document.getElementById('trustee-summary');
  if(ts){
    ts.style.display=currentRole==='trustee'?'block':'none';
    const txt=document.getElementById('trustee-summary-text');
    if(txt) txt.textContent=`There are currently ${totalP} registered policies, of which ${pubP} are published and active. ${overP+overS>0?`${overP+overS} document(s) are overdue for review and require attention.`:'All documents are within their review cycles.'} The Governance Maturity Index is ${gmi.toFixed(1)}/5 (${gmiLbl}). ${hiR} high-risk policies are in effect, all requiring annual review by the Board or National Director.`;
  }

  // Heatmap
  const cats=['Governance','HR','IT','Finance','Operations','Bursaries','Funders'];
  const all=[...polData,...sopData];
  document.getElementById('heatmap-body').innerHTML=cats.map(c=>{
    const cDocs=all.filter(d=>d.category===c&&d.status!=='Archived');
    const lo=cDocs.filter(d=>d.risk==='Low'||(d.tier==='T3'||d.tier==='T4')).length;
    const hi=cDocs.filter(d=>d.risk==='High Strategic'||d.risk==='High Operational').length;
    const me=Math.max(0,cDocs.length-lo-hi);
    return`<tr><td class="hm-l">${c}</td><td class="hl">${lo||0}</td><td class="hm">${me||0}</td><td class="hh">${hi||0}</td></tr>`;
  }).join('');

  // Cross filter
  document.getElementById('cross-filter-body').innerHTML=polData.filter(p=>p.status!=='Archived').slice(0,8).map(p=>{
    const sops=sopData.filter(s=>s.linkedPolicy===p.polId).length;
    return`<tr><td style="font-size:12px">${esc(p.polId)}</td><td>${esc(p.version)}</td><td>${sops}</td><td>${datecell(p.reviewDate)}</td><td>${rbadge(p.risk)}</td></tr>`;
  }).join('');

  // Pipeline
  const stages=[
    {label:'DRAFT',statuses:['Draft'],color:'#90a4ae',bg:'#eceff1'},
    {label:'CONSULT',statuses:['Consult'],color:'#1976d2',bg:'#e3f2fd'},
    {label:'PTT REVIEW',statuses:['PTT Review'],color:var2('--st-a')},
    {label:'APPROVED',statuses:['Approved','New'],color:'#388e3c',bg:'#e8f5e9'},
    {label:'PUBLISHED',statuses:['Published'],color:'#2e7d32',bg:'#c8e6c9'},
    {label:'REVIEW DUE',statuses:['Review Due'],color:'#f57f17',bg:var2('--am-bg')},
  ];
  function var2(n){return getComputedStyle(document.documentElement).getPropertyValue(n).trim();}
  const pAll=[...polData,...sopData];
  document.getElementById('pipeline-bar').innerHTML=stages.map(st=>{
    const n=pAll.filter(d=>st.statuses.includes(d.status)).length;
    return`<div class="ps" style="background:${st.bg||'#f5f5f5'};border:1px solid ${st.color}20"><div class="ps-l" style="color:${st.color}">${st.label}</div><div class="ps-v" style="color:${st.color}">${n}</div><div class="ps-sub" style="color:${st.color}">${n===1?'document':'documents'}</div></div>`;
  }).join('<div style="display:flex;align-items:center;color:#bbb;font-size:14px;padding:0 2px">›</div>');
  document.getElementById('pipeline-total').textContent=`${pAll.filter(d=>d.status!=='Archived').length} active`;
}

function renderMaturity(){
  const totalP=polData.length||1;
  const pubP=polData.filter(p=>isPublished(p.status)).length;
  const overP=polData.filter(p=>isOverdue(p.reviewDate)&&p.status!=='Archived').length;
  const withSop=polData.filter(p=>sopData.some(s=>s.linkedPolicy===p.polId)).length;
  const fullMeta=polData.filter(p=>p.owner&&p.reviewDate&&p.custodian&&p.approver).length;

  const comp=Math.round(pubP/totalP*100);
  const revC=Math.round(Math.max(0,(totalP-overP)/totalP)*100);
  const sopL=Math.round(withSop/totalP*100);
  const meta=Math.round(fullMeta/totalP*100);
  const avg=Math.round((comp+revC+sopL+meta)/4);

  document.getElementById('mat-cards').innerHTML=`
    <div class="kpi"><div class="kpi-l">Overall Maturity</div><div class="kpi-v" style="color:var(--st)">${avg}%</div><div class="kpi-s">Composite score</div></div>
    <div class="kpi"><div class="kpi-l">Publication Rate</div><div class="kpi-v" style="color:var(--gn-tx)">${comp}%</div><div class="kpi-s">${pubP}/${totalP} published</div></div>
    <div class="kpi"><div class="kpi-l">Review Compliance</div><div class="kpi-v" style="color:${revC<80?'var(--rd-tx)':'var(--gn-tx)'}">${revC}%</div><div class="kpi-s">${overP} overdue</div></div>
    <div class="kpi"><div class="kpi-l">SOP Linkage</div><div class="kpi-v" style="color:${sopL<70?'var(--am-tx)':'var(--gn-tx)'}">${sopL}%</div><div class="kpi-s">${withSop}/${totalP} linked</div></div>`;

  const cats=['Governance','HR','IT','Finance','Operations','Bursaries','Funders'];
  document.getElementById('mat-cats').innerHTML=cats.map(c=>{
    const cP=polData.filter(p=>p.category===c);
    const pct=cP.length?Math.round(cP.filter(p=>isPublished(p.status)).length/cP.length*100):0;
    const col=pct>=80?'#2e7d32':pct>=50?'#f57f17':'#c62828';
    return`<div class="prog-row"><div class="prog-meta"><span>${c}</span><span style="font-weight:600;color:${col}">${pct}%</span></div><div class="prog-bar"><div class="prog-fill" style="width:${pct}%;background:${col}"></div></div></div>`;
  }).join('');

  const frameworks=[
    ['NPO Act Compliance','Published NPO-related policies',90],
    ['POPIA Alignment','Data protection policy & procedures',75],
    ['King V Alignment','Governance structure & oversight',80],
    ['Financial Controls','Financial management & procurement',85],
    ['HR Framework','HR policies and SOPs',70],
  ];
  document.getElementById('mat-comp').innerHTML=frameworks.map(([n,d,v])=>{
    const col=v>=80?'#2e7d32':v>=60?'#f57f17':'#c62828';
    return`<div class="prog-row"><div class="prog-meta"><span style="font-size:11px">${n}</span><span style="font-size:11px;font-weight:600;color:${col}">${v}%</span></div><div class="prog-bar"><div class="prog-fill" style="width:${v}%;background:${col}"></div></div><div style="font-size:10px;color:var(--tx3);margin-top:2px">${d}</div></div>`;
  }).join('');
}

function renderChangelog(){
  const tbody=document.getElementById('changelog-body'); if(!tbody) return;
  if(!changeLog.length){tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:var(--tx3);padding:20px">No changes recorded.</td></tr>';return;}
  tbody.innerHTML=changeLog.slice(0,50).map(e=>`<tr><td style="font-size:11px">${esc(e.changeDate||'')}</td><td>${esc(e.docId)}</td><td>${sbadge(e.fromStatus)}</td><td>${sbadge(e.toStatus)}</td><td>${esc(e.changedBy)}</td><td style="font-size:11px;color:var(--tx2)">${esc(e.note||'')}</td></tr>`).join('');
}

function renderAll(){
  const gsw=document.getElementById('gs-wrap'); if(gsw) gsw.style.display=(currentRole==='admin'||currentRole==='ptt'||currentRole==='legal')?'block':'none';
  renderExec();
  renderPol();
  renderSop();
  renderMaturity();
  renderChangelog();
  renderTrustees();
  renderAck();
  renderPendingAck();
  checkReviewReminders();
}

// ════════════════════════════════════════════
// STATUS CHANGE
// ════════════════════════════════════════════
// ════════════════════════════════════════════
// EDIT ROW MODAL
// ════════════════════════════════════════════
function openEditModal(docId, type){
  const arr = type==='pol' ? polData : sopData;
  const doc = arr.find(d => (type==='pol' ? d.polId : d.sopId) === docId);
  if(!doc){ toast('Document not found.','error'); return; }
  document.getElementById('edit-modal-title').textContent = (type==='pol' ? 'Edit Policy — ' : 'Edit SOP — ') + docId;
  document.getElementById('edit-id').value = docId;
  document.getElementById('edit-type').value = type;
  document.getElementById('edit-docid').value = docId;
  document.getElementById('edit-title').value = doc.title||'';
  document.getElementById('edit-version').value = doc.version||'';
  setSelectVal('edit-cat', doc.category||'Governance');
  if(type==='pol'){
    setSelectVal('edit-tier', doc.tier||'T2');
    setSelectVal('edit-risk', doc.risk||'Medium');
    document.getElementById('edit-approver').value = doc.approver||'';
    document.getElementById('edit-linked-pol-row').style.display='none';
  } else {
    document.getElementById('edit-linked-pol-row').style.display='';
    document.getElementById('edit-linked-pol').value = doc.linkedPolicy||'';
    document.getElementById('edit-approver').value='';
  }
  setSelectVal('edit-status', doc.status||'Draft');
  document.getElementById('edit-owner').value = doc.owner||'';
  document.getElementById('edit-custodian').value = doc.custodian||'';
  document.getElementById('edit-rev-date').value = doc.reviewDate||'';
  document.getElementById('edit-notes').value = doc.notes||'';
  document.getElementById('edit-draft-url').value = doc.draftUrl||'';
  document.getElementById('edit-pub-url').value = doc.publishedUrl||'';
  document.getElementById('edit-arch-url').value = doc.archiveUrl||'';
  openM('m-edit');
  loadVersionHistory(docId, type);
}

function setSelectVal(id, val){
  const el=document.getElementById(id); if(!el) return;
  for(let i=0;i<el.options.length;i++){ if(el.options[i].value===val||el.options[i].text===val){ el.selectedIndex=i; return; } }
}

async function saveEdit(){
  const type = document.getElementById('edit-type').value;
  const id   = document.getElementById('edit-id').value;
  const arr  = type==='pol' ? polData : sopData;
  const doc  = arr.find(d => (type==='pol' ? d.polId : d.sopId) === id);
  if(!doc){ toast('Document not found.','error'); return; }

  doc.title      = document.getElementById('edit-title').value.trim();
  doc.version    = document.getElementById('edit-version').value.trim();
  doc.category   = document.getElementById('edit-cat').value;
  doc.status     = document.getElementById('edit-status').value;
  doc.owner      = document.getElementById('edit-owner').value.trim();
  doc.custodian  = document.getElementById('edit-custodian').value.trim();
  doc.reviewDate = document.getElementById('edit-rev-date').value;
  doc.notes      = document.getElementById('edit-notes').value.trim();
  doc.draftUrl   = document.getElementById('edit-draft-url').value.trim();
  doc.publishedUrl = document.getElementById('edit-pub-url').value.trim();
  doc.archiveUrl = document.getElementById('edit-arch-url').value.trim();
  if(type==='pol'){
    doc.tier     = document.getElementById('edit-tier').value;
    doc.risk     = document.getElementById('edit-risk').value;
    doc.approver = document.getElementById('edit-approver').value.trim();
  } else {
    doc.linkedPolicy = document.getElementById('edit-linked-pol').value.trim();
  }

  try {
    const listKey = type==='pol' ? LISTS.pol : LISTS.sop;
    const fields  = type==='pol' ? toSpPol(doc) : toSpSop(doc);
    await spPatchResilient(listKey, doc._spId, fields, type);
    renderAll();
    closeM('m-edit');
    // Only show success toast if no warning was already shown by spPatchResilient
    const warnShown = (document.querySelector('.toast-warn')||null);
    if(!warnShown) toast((type==='pol'?'Policy':'SOP')+' updated: '+id, 'success');
  } catch(e){ toast('Save failed: '+e.message, 'error'); }
}

function triggerEditUpload(field){
  const inp=document.getElementById('edit-file-input');
  inp.dataset.field=field; inp.click();
}

async function handleEditUpload(input){
  const file=input.files[0]; if(!file) return;
  const field=input.dataset.field;
  const type=document.getElementById('edit-type').value;
  const id=document.getElementById('edit-id').value;
  const arr=type==='pol'?polData:sopData;
  const doc=arr.find(d=>(type==='pol'?d.polId:d.sopId)===id);
  if(!doc){toast('Open the editor first.','error');input.value='';return;}
  let status=doc.status;
  if(field==='pub') status='Published';
  if(field==='arch') status='Archived';
  const folder=govDocFolder(type,status,doc.category);
  toast('Uploading '+file.name+'…','ok');
  try{
    const result=await uploadGovFile('Governance Published',folder,file);
    const urlField={draft:'edit-draft-url',pub:'edit-pub-url',arch:'edit-arch-url'}[field];
    if(urlField) document.getElementById(urlField).value=result.webUrl||'';
    toast('Uploaded: '+file.name,'ok');
  }catch(e){toast('Upload failed: '+e.message,'error');}
  input.value='';
}

// ════════════════════════════════════════════
function openStatusModal(docId,type,currentStatus){
  statusTarget={id:docId,type};
  document.getElementById('status-modal-title').textContent='Change Status — '+docId;
  const cur=document.getElementById('status-current');
  cur.textContent=currentStatus; cur.className='bd '+(({Initiated:'bpr',Draft:'ba',Consult:'ba','PTT Review':'bt',Approved:'bt',New:'bb',Published:'bg','Review Due':'ba',Archived:'bgy'})[currentStatus]||'bgy');
  document.getElementById('f-status-note').value='';
  document.getElementById('status-warn').style.display='none';
  // Reset step panels
  const pubStep=document.getElementById('publish-step');
  const archStep=document.getElementById('archive-step');
  if(pubStep)  pubStep.style.display='none';
  if(archStep) archStep.style.display='none';
  ['f-pub-url','f-arch-url','f-arch-url-arch'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  const next=allowedTransitions(currentStatus,currentRole);
  const badges={Draft:'ba',Consult:'ba','PTT Review':'bt',Approved:'bt',New:'bb',Published:'bg','Review Due':'ba',Archived:'bgy'};
  document.getElementById('status-opts').innerHTML=next.length?next.map(s=>`<button class="bd ${badges[s]||'bgy'}" style="cursor:pointer;border:2px solid transparent" onclick="selectStatus(this,'${s}')" data-s="${s}">${s}</button>`).join(''):'<span style="font-size:12px;color:var(--tx3)">No transitions available for your role.</span>';
  openM('m-status');
}

let selStatus=null;
function selectStatus(btn,s){
  document.querySelectorAll('#status-opts button').forEach(b=>b.style.border='2px solid transparent');
  btn.style.border='2px solid #333'; selStatus=s;

  // Show/hide contextual step panels
  const pubStep  = document.getElementById('publish-step');
  const archStep = document.getElementById('archive-step');
  if(pubStep)  pubStep.style.display  = s==='Published' ? 'block' : 'none';
  if(archStep) archStep.style.display = s==='Archived'  ? 'block' : 'none';

  // Clear URL fields when switching
  ['f-pub-url','f-arch-url','f-arch-url-arch'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });

  const warns={Archived:'This document will be removed from all user views. Ensure the file has been moved to Governance Archive in SharePoint first.'};
  const w=document.getElementById('status-warn');
  if(warns[s]){w.textContent=warns[s];w.style.display='block';}else w.style.display='none';
}

async function applyStatus(){
  if(!selStatus){toast('Please select a status.','warning');return;}
  const note=document.getElementById('f-status-note').value.trim();
  if(!note){toast('Change notes are mandatory for audit trail compliance.','warning');return;}

  // Publisher gate — only admins and members of Dashboards - Policy-Publishers may publish
  if(selStatus==='Published' && currentRole !== 'admin'){
    if(!_diagGroups.some(g=>g.toLowerCase()===PUBLISHERS_GROUP.toLowerCase())){
      toast('Only authorised publishers may set a document to Published status.','error');
      return;
    }
  }

  // Validate URL fields for publishing/archiving transitions
  if(selStatus==='Published'){
    const pubUrl=(document.getElementById('f-pub-url')||{}).value||'';
    if(!pubUrl){toast('Please paste the Published document URL before applying.','warning');return;}
  }
  if(selStatus==='Archived'){
    const archUrl=(document.getElementById('f-arch-url-arch')||{}).value||'';
    if(!archUrl){toast('Please paste the Archive document URL before applying.','warning');return;}
  }

  const {id,type}=statusTarget;
  const arr=type==='pol'?polData:sopData;
  const doc=arr.find(d=>(type==='pol'?d.polId:d.sopId)===id);
  if(!doc){toast('Document not found.','error');return;}
  const oldStatus=doc.status;
  doc.status=selStatus;

  // Build the SharePoint field update — include URLs if provided
  const spFields={Status:selStatus};
  if(selStatus==='Published'){
    const pubUrl=(document.getElementById('f-pub-url')||{}).value.trim()||'';
    const archUrl=(document.getElementById('f-arch-url')||{}).value.trim()||'';
    if(pubUrl){ doc.publishedUrl=pubUrl; spFields.PublishedUrl=spWriteValue('publishedUrl',pubUrl,doc); }
    if(archUrl){ doc.archiveUrl=archUrl; spFields.ArchiveUrl=spWriteValue('archiveUrl',archUrl,doc); doc.draftUrl=''; spFields.DraftUrl=null; }
  }
  if(selStatus==='Archived'){
    const archUrl=(document.getElementById('f-arch-url-arch')||{}).value.trim()||'';
    if(archUrl){ doc.archiveUrl=archUrl; spFields.ArchiveUrl=spWriteValue('archiveUrl',archUrl,doc); }
    doc.publishedUrl=''; spFields.PublishedUrl=null;
  }

  if(type==='sop'&&!validateSOBBeforeTransition(doc,selStatus))return;
  try{
    await spPatch(type==='pol'?LISTS.pol:LISTS.sop,doc._spId,spFields);
    const logItem={note,docId:id,docType:type,fromStatus:oldStatus,toStatus:selStatus,changedBy:userName,changeDate:td()};
    const lr=await spPost(LISTS.log,{Title:note,DocID:id,DocType:type,FromStatus:oldStatus,ToStatus:selStatus,ChangedBy:userName,ChangeDate:td()+'T00:00:00Z'});
    logItem._spId=lr?.id||'';
    changeLog.unshift(logItem);
    renderAll(); closeM('m-status');
    const urlMsg = selStatus==='Published'?' · Published URL saved.' : selStatus==='Archived'?' · Archive URL saved.' : '';
    toast(`${id} → Status changed to "${selStatus}". Change logged.${urlMsg}`,'success');
    // Fire Teams notifications asynchronously
    fireStatusChangeNotif(id, doc.title||id, oldStatus, selStatus, userName);
  }catch(e){toast('Failed to save: '+e.message,'error');}
  selStatus=null;
}

// ════════════════════════════════════════════
// URL MANAGEMENT
// ════════════════════════════════════════════
function openUrlModal(docId,type){
  urlTarget={id:docId,type};
  document.getElementById('url-modal-title').textContent='Edit Document URLs — '+docId;
  const arr=type==='pol'?polData:sopData;
  const doc=arr.find(d=>(type==='pol'?d.polId:d.sopId)===docId);
  document.getElementById('f-url-draft').value=doc?.draftUrl||'';
  document.getElementById('f-url-pub').value=doc?.publishedUrl||'';
  document.getElementById('f-url-arch').value=doc?.archiveUrl||'';
  // PTT scope: lock Published and Archive URL fields (read-only)
  const isP = currentRole === 'ptt' || currentRole === 'legal';
  const pub = document.getElementById('f-url-pub');
  const arc = document.getElementById('f-url-arch');
  [pub,arc].forEach(el=>{
    if(!el) return;
    el.readOnly = isP;
    el.style.background = isP ? '#f1f3f5' : '';
    el.title = isP ? 'Only an Admin can edit Published and Archive URLs' : '';
  });
  openM('m-urls');
}

async function saveUrls(){
  const {id,type}=urlTarget;
  const arr=type==='pol'?polData:sopData;
  const doc=arr.find(d=>(type==='pol'?d.polId:d.sopId)===id);
  if(!doc){closeM('m-urls');return;}
  // PTT may only update the draft URL
  if(currentRole === 'ptt'){
    doc.draftUrl=document.getElementById('f-url-draft').value.trim();
    try{
      await spPatch(type==='pol'?LISTS.pol:LISTS.sop,doc._spId,
        {DraftUrl: spWriteValue('draftUrl', doc.draftUrl, doc)});
      renderAll(); closeM('m-urls');
      toast('Draft URL saved for '+id,'success');
    }catch(e){toast('Failed to save URL: '+e.message,'error');}
    return;
  }
  doc.draftUrl=document.getElementById('f-url-draft').value.trim();
  doc.publishedUrl=document.getElementById('f-url-pub').value.trim();
  doc.archiveUrl=document.getElementById('f-url-arch').value.trim();
  try{
    await spPatch(type==='pol'?LISTS.pol:LISTS.sop,doc._spId,{
      DraftUrl: spWriteValue('draftUrl', doc.draftUrl, doc),
      PublishedUrl: spWriteValue('publishedUrl', doc.publishedUrl, doc),
      ArchiveUrl: spWriteValue('archiveUrl', doc.archiveUrl, doc)
    });
    renderAll(); closeM('m-urls');
    toast('Document URLs saved for '+id,'success');
  }catch(e){toast('Failed to save URLs: '+e.message,'error');}
}

// ════════════════════════════════════════════
// ADD POLICY / SOP
// ════════════════════════════════════════════
function autoFillApprover(){
  const tier=document.getElementById('f-pol-tier').value;
  const risk=document.getElementById('f-pol-risk').value;
  const el=document.getElementById('f-pol-app'); if(!el) return;
  if(tier==='T1'||risk==='High Strategic') el.value='Dashboards - Board-of-Trustees (quorum ≥50%+1, legal consultation mandatory)';
  else if(risk==='High Operational') el.value='National Director (or Deputy ND if formally delegated in writing)';
  else if(tier==='T2'||tier==='T3') el.value='National Director + Leadership Team';
  else el.value='Leadership Team (majority quorum)';
}

async function submitPol(){
  const pid=document.getElementById('f-pol-id').value.trim();
  const title=document.getElementById('f-pol-title').value.trim();
  if(!pid||!title){toast('Policy ID and Title are required.','warning');return;}
  if(polData.some(p=>p.polId===pid)){toast('Policy ID '+pid+' already exists.','error');return;}
  const p={polId:pid,title,category:document.getElementById('f-pol-cat').value,tier:document.getElementById('f-pol-tier').value,version:document.getElementById('f-pol-ver').value||'0.1',status:document.getElementById('f-pol-status').value||'Draft',risk:document.getElementById('f-pol-risk').value||'Medium',approver:document.getElementById('f-pol-app').value,owner:document.getElementById('f-pol-owner').value,custodian:document.getElementById('f-pol-cust').value,notes:document.getElementById('f-pol-notes').value,reviewDate:document.getElementById('f-pol-rev').value,draftUrl:document.getElementById('f-pol-draft-url').value,publishedUrl:'',archiveUrl:''};
  try{
    const r=await spPost(LISTS.pol,toSpPol(p));
    p._spId=r.id; polData.unshift(p);
    const lr=await spPost(LISTS.log,{Title:'New policy registered',DocID:pid,DocType:'pol',FromStatus:'—',ToStatus:p.status,ChangedBy:userName,ChangeDate:td()+'T00:00:00Z'});
    changeLog.unshift({note:'New policy registered',docId:pid,docType:'pol',fromStatus:'—',toStatus:p.status,changedBy:userName,changeDate:td(),_spId:lr?.id});
    renderAll(); closeM('m-pol');
    toast('Policy '+pid+' registered as '+p.status+'.','success');
    fireNewDocNotif(pid, p.title, 'Policy', p.status, userName);
  }catch(e){toast('Failed to add policy: '+e.message,'error');}
}

async function submitSop(){
  const sid=document.getElementById('f-sop-id').value.trim();
  const title=document.getElementById('f-sop-title').value.trim();
  if(!sid||!title){toast('SOP ID and Title are required.','warning');return;}
  if(sopData.some(s=>s.sopId===sid)){toast('SOP ID '+sid+' already exists.','error');return;}
  const s={sopId:sid,title,category:document.getElementById('f-sop-cat').value,linkedPolicy:document.getElementById('f-sop-pol').value,version:document.getElementById('f-sop-ver').value||'0.1',status:document.getElementById('f-sop-status').value||'Draft',reviewDate:'',owner:document.getElementById('f-sop-owner').value,custodian:(document.getElementById('f-sop-cust')||{}).value||'',notes:(document.getElementById('f-sop-notes')||{}).value||'',draftUrl:document.getElementById('f-sop-draft-url').value,publishedUrl:'',archiveUrl:''};
  try{
    const r=await spPost(LISTS.sop,toSpSop(s));
    s._spId=r.id; sopData.unshift(s);
    const lr=await spPost(LISTS.log,{Title:'New SOP registered',DocID:sid,DocType:'sop',FromStatus:'—',ToStatus:s.status,ChangedBy:userName,ChangeDate:td()+'T00:00:00Z'});
    changeLog.unshift({note:'New SOP registered',docId:sid,docType:'sop',fromStatus:'—',toStatus:s.status,changedBy:userName,changeDate:td(),_spId:lr?.id});
    renderAll(); closeM('m-sop');
    toast('SOP '+sid+' registered.','success');
    fireNewDocNotif(sid, s.title, 'SOP', s.status, userName);
  }catch(e){toast('Failed to add SOP: '+e.message,'error');}
}

function polDropdown(){
  const sel=document.getElementById('f-sop-pol'); if(!sel) return;
  sel.innerHTML=polData.map(p=>`<option value="${esc(p.polId)}">${esc(p.polId)} — ${esc(p.title)}</option>`).join('');
}

// ════════════════════════════════════════════
// BACKUP / RESTORE
// ════════════════════════════════════════════
// ════════════════════════════════════════════
// DEDUPLICATE
// ════════════════════════════════════════════
async function deduplicateAll(){
  const polDups = findDups(polData, 'polId');
  const sopDups = findDups(sopData, 'sopId');
  const total = polDups.length + sopDups.length;
  if(total === 0){ toast('No duplicates found.', 'success'); return; }
  if(!confirm(`Found ${polDups.length} duplicate policy item(s) and ${sopDups.length} duplicate SOP item(s).\n\nThe duplicate with the least data will be deleted, keeping the most complete record.\n\nProceed?`)) return;
  setLoad('Removing ' + total + ' duplicate item(s)…');
  try{
    const sid = await getSiteId();
    for(const dup of polDups){
      await gFetch(`/sites/${sid}/lists/${LISTS.pol}/items/${dup._spId}`, {method:'DELETE'});
    }
    for(const dup of sopDups){
      await gFetch(`/sites/${sid}/lists/${LISTS.sop}/items/${dup._spId}`, {method:'DELETE'});
    }
    await loadData(); renderAll();
    toast(`Removed ${total} duplicate(s). Registers are now clean.`, 'success');
  }catch(e){
    toast('Deduplicate error: ' + e.message, 'error');
  }finally{ hideLoad(); }
}

function findDups(arr, idField){
  // Group by ID field. For each group with more than one item,
  // keep the one with the most non-empty fields and mark the rest for deletion.
  const groups = {};
  arr.forEach(item => {
    const key = item[idField] || item.title || item._spId;
    if(!groups[key]) groups[key] = [];
    groups[key].push(item);
  });
  const toDelete = [];
  Object.values(groups).forEach(group => {
    if(group.length <= 1) return;
    // Score each by number of non-empty fields — keep the highest score
    const scored = group.map(item => ({
      item,
      score: Object.values(item).filter(v => v && v !== '' && v !== item._spId).length
    })).sort((a,b) => b.score - a.score);
    // Keep first (highest score), delete the rest
    scored.slice(1).forEach(s => toDelete.push(s.item));
  });
  return toDelete;
}

function downloadBackup(){
  const pkg={version:4,exportDate:new Date().toISOString(),org:'StudyTrust — 000-601 NPO · IT3895/11(T)',polData,sopData,changeLog};
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([JSON.stringify(pkg,null,2)],{type:'application/json'}));
  a.download='StudyTrust_GovHub_Backup_'+td()+'.json'; a.click();
  toast('Backup downloaded.','success');
}

async function doRestore(){
  let pkg;
  try{pkg=JSON.parse(document.getElementById('restore-data').value.trim());}
  catch(e){toast('Invalid JSON: '+e.message,'error');return;}
  if(!confirm('Post all backup items to SharePoint? Existing items are not deleted.')) return;

  setLoad('Restoring to SharePoint…');
  const fieldErrors=[], restoreErrors=[];
  let posted=0;

  function clean(fields){
    const out={};
    for(const[k,v]of Object.entries(fields)){
      if(v!==null&&v!==undefined&&v!=='') out[k]=v;
    }
    return out;
  }

  async function safePost(listName, fullFields, label){
    const cf=clean(fullFields);
    try{
      await spPost(listName,cf);
      posted++;
      return;
    }catch(e){/* fall through to field-by-field */}
    // Post with Title only, then PATCH each field individually
    const base={Title:cf.Title||label};
    let itemId=null;
    try{
      const r=await spPost(listName,base);
      itemId=r?.id; posted++;
    }catch(e2){
      restoreErrors.push(`${listName} — ${label}: ${e2.message}`);
      return;
    }
    const sid=await getSiteId();
    const bad=[];
    for(const[k,v]of Object.entries(cf)){
      if(k==='Title') continue;
      try{
        await gFetch(`/sites/${sid}/lists/${listName}/items/${itemId}`,{
          method:'PATCH',body:JSON.stringify({fields:{[k]:v}})
        });
      }catch(fe){ bad.push(k); }
    }
    if(bad.length) fieldErrors.push(...bad);
  }

  for(const p of (pkg.polData||[])) await safePost(LISTS.pol, toSpPol(p), p.polId||p.title||'Policy');
  for(const s of (pkg.sopData||[])) await safePost(LISTS.sop, toSpSop(s), s.sopId||s.title||'SOP');

  closeM('m-restore');
  await loadData(); renderAll(); hideLoad();

  const uniqueBad=[...new Set(fieldErrors)];
  if(restoreErrors.length===0 && uniqueBad.length===0){
    toast('Restore complete — '+posted+' items posted to SharePoint.','success');
  } else {
    let msg='Restore: '+posted+' items created.';
    if(uniqueBad.length) msg+=' Missing columns: '+uniqueBad.join(', ')+'. Create these in SharePoint then restore again.';
    toast(msg,'warning');
    if(uniqueBad.length){
      alert('The following SharePoint columns do not exist in GovPolicies:\n\n'+uniqueBad.map(f=>'  • '+f).join('\n')+'\n\nCreate these columns (Single line of text) in SharePoint, then run Restore again.');
    }
  }
}

function exportPolCsv(){
  const rows=[['PolicyID','Title','Category','Tier','Version','Status','Risk','Approver','Review Date','Owner','Published URL'],...polData.map(p=>[p.polId,p.title,p.category,p.tier,p.version,p.status,p.risk,p.approver,p.reviewDate,p.owner,p.publishedUrl].map(v=>'"'+(v||'').replace(/"/g,'""')+'"'))];
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([rows.map(r=>r.join(',')).join('\n')],{type:'text/csv'}));
  a.download='StudyTrust_Policies_'+td()+'.csv'; a.click();
}
function exportSopCsv(){
  const rows=[['SopID','Title','Category','Linked Policy','Version','Status','Review Date','Owner','Published URL'],...sopData.map(s=>[s.sopId,s.title,s.category,s.linkedPolicy,s.version,s.status,s.reviewDate,s.owner,s.publishedUrl].map(v=>'"'+(v||'').replace(/"/g,'""')+'"'))];
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([rows.map(r=>r.join(',')).join('\n')],{type:'text/csv'}));
  a.download='StudyTrust_SOPs_'+td()+'.csv'; a.click();
}

// ════════════════════════════════════════════
// TABS & ROLE
// ════════════════════════════════════════════
const TABS_ADMIN=[{id:'sec-exec',l:'📊 Executive Overview'},{id:'sec-workflow',l:'🔄 Approval Workflow'},{id:'sec-pol',l:'📋 Policy Register'},{id:'sec-sop',l:'📑 SOP Register'},{id:'sec-mat',l:'📈 Maturity Tracker'},{id:'sec-docs',l:'📁 Document Libraries'},{id:'sec-trustees',l:'🏛️ Board of Trustees'},{id:'sec-ack',l:'✅ Acknowledgements'},{id:'sec-training',l:'🎓 Training'},{id:'sec-hub',l:'🏠 Governance Hub'}];
const TABS_TRUSTEE=[{id:'sec-exec',l:'Governance Overview'},{id:'sec-pol',l:'Policy Register'},{id:'sec-sop',l:'SOP Register'},{id:'sec-workflow',l:'Approval Framework'},{id:'sec-mat',l:'Maturity Overview'},{id:'sec-trustees',l:'Board of Trustees'},{id:'sec-hub',l:'Governance Hub'}];
const TABS_AUDITOR=[{id:'sec-aud-docs',l:'Policy & SOP Documents'},{id:'sec-aud-overview',l:'Governance Overview'},{id:'sec-aud-request',l:'Document Requests'}];
const TABS_PTT=[{id:'sec-exec',l:'Governance Overview'},{id:'sec-workflow',l:'Approval Workflow'},{id:'sec-pol',l:'Policy Register'},{id:'sec-sop',l:'SOP Register'},{id:'sec-mat',l:'Maturity Tracker'},{id:'sec-docs',l:'Document Libraries'},{id:'sec-hub',l:'Quick Access'}];
const TABS_PERSONNEL=[{id:'sec-pol',l:'Policies'},{id:'sec-sop',l:'SOPs'},{id:'sec-hub',l:'Quick Access'}];

// Status transitions allowed by role.
// PTT may move documents within the drafting pipeline but not to Published or Archived.

// ════════════════════════════════════════════════════════════════
// v4.5 NEW FEATURES
// ════════════════════════════════════════════════════════════════

// ── GLOBAL SEARCH ────────────────────────────────────────────────
function globalSearch(q){
  const box = document.getElementById('gs-results');
  if(!box) return;
  if(!q||q.trim().length<2){ box.style.display='none'; box.innerHTML=''; return; }
  box.style.display='block';
  const t = q.toLowerCase();
  const polHits = polData.filter(p=>(p.polId+p.title+p.category+p.owner+p.custodian+p.notes).toLowerCase().includes(t)).slice(0,5);
  const sopHits = sopData.filter(s=>(s.sopId+s.title+s.category+s.owner+s.custodian).toLowerCase().includes(t)).slice(0,4);
  if(!polHits.length && !sopHits.length){
    box.innerHTML='<div class="gs-empty">No results for "'+esc(q)+'"</div>'; return;
  }
  box.innerHTML = [
    ...polHits.map(p=>`<div class="gsr" onclick="gotoDoc('${esc(p.polId)}','pol')">
      <span class="gsr-badge gsr-pol">${esc(p.polId)}</span>
      <div><div class="gsr-title">${esc(p.title)}</div>
      <div class="gsr-meta">${esc(p.category)} · ${esc(p.status)} · Owner: ${esc(p.owner||'—')}</div></div>
    </div>`),
    ...sopHits.map(s=>`<div class="gsr" onclick="gotoDoc('${esc(s.sopId)}','sop')">
      <span class="gsr-badge gsr-sop">${esc(s.sopId)}</span>
      <div><div class="gsr-title">${esc(s.title)}</div>
      <div class="gsr-meta">${esc(s.category)} · ${esc(s.status)} · Owner: ${esc(s.owner||'—')}</div></div>
    </div>`)
  ].join('');
}

function gotoDoc(docId, type){
  document.getElementById('gs-input').value='';
  const box=document.getElementById('gs-results'); if(box) box.style.display='none';
  showTab(type==='pol'?'sec-pol':'sec-sop');
  setTimeout(()=>openEditModal(docId,type),150);
}

// Close search on outside click
document.addEventListener('click', e=>{
  const wrap=document.getElementById('gs-wrap');
  if(wrap && !wrap.contains(e.target)){
    const box=document.getElementById('gs-results');
    if(box) box.style.display='none';
  }
});

// ── REVIEW REMINDER BANNER ────────────────────────────────────────
function checkReviewReminders(){
  const banner=document.getElementById('rev-reminder-banner'); if(!banner) return;
  if(currentRole!=='admin'&&currentRole!=='ptt'){banner.style.display='none';return;}
  const today=new Date(); today.setHours(0,0,0,0);
  const in30=new Date(today); in30.setDate(in30.getDate()+30);
  const in60=new Date(today); in60.setDate(in60.getDate()+60);
  const parseDate=s=>{ if(!s) return null; const d=new Date(s); return isNaN(d)?null:d; };
  const soon30=[...polData,...sopData].filter(d=>{
    const rd=parseDate(d.reviewDate); if(!rd||d.status==='Archived') return false;
    return rd>today && rd<=in30;
  });
  const soon60=[...polData,...sopData].filter(d=>{
    const rd=parseDate(d.reviewDate); if(!rd||d.status==='Archived') return false;
    return rd>today && rd<=in60 && rd>in30;
  });
  if(!soon30.length && !soon60.length){ banner.style.display='none'; return; }
  banner.style.display='block';
  let msg='';
  if(soon30.length) msg+=`<strong>⚠ Due within 30 days (${soon30.length}):</strong> `+soon30.map(d=>`${d.polId||d.sopId} (${d.reviewDate})`).join(', ')+'. ';
  if(soon60.length) msg+=`<strong>📅 Due within 60 days (${soon60.length}):</strong> `+soon60.map(d=>`${d.polId||d.sopId} (${d.reviewDate})`).join(', ')+'.';
  banner.innerHTML=msg;
  // Fire webhook once per week max
  const wkey='st_gov_rev_notif_'+new Date().toISOString().slice(0,10);
  const fired=localStorage.getItem(wkey);
  if(!fired&&soon30.length){
    const wh=loadWebhooks();
    if(wh.ptt){
      sendAdaptiveCard(wh.ptt,'📅 Policy Reviews Due Soon',`${soon30.length} policy/SOP review(s) are due within 30 days.`,
        soon30.map(d=>`• ${d.polId||d.sopId}: ${d.title} — due ${d.reviewDate}`).join('\n'),'ptt');
      localStorage.setItem(wkey,'1');
    }
  }
}

// ── BOARD REPORT ──────────────────────────────────────────────────
function generateBoardReport(){
  const today=new Date().toLocaleDateString('en-ZA',{year:'numeric',month:'long',day:'numeric'});
  const statuses=['Initiated','Draft','Consult','PTT Review','Approved','Published','Review Due','Archived'];
  const counts={};
  statuses.forEach(s=>counts[s]=0);
  [...polData,...sopData].forEach(d=>{counts[d.status]=(counts[d.status]||0)+1;});
  const published=polData.filter(p=>p.status==='Published'||p.status==='Review Due').length;
  const overdue=[...polData,...sopData].filter(d=>isOverdue(d.reviewDate)&&d.status!=='Archived');
  const hiRisk=polData.filter(p=>p.risk==='High Strategic'&&(p.status==='Published'||p.status==='Review Due'||p.status==='Approved'));
  const initiated=[...polData,...sopData].filter(d=>d.status==='Initiated');
  const matPct=Math.round((published/Math.max(polData.length+sopData.length,1))*100);

  const catCounts={};
  polData.filter(p=>p.status==='Published'||p.status==='Review Due').forEach(p=>{
    catCounts[p.category]=(catCounts[p.category]||0)+1;
  });

  const rpt=window.open('','_blank');
  rpt.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>StudyTrust Governance Report — ${today}</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Aptos,Calibri,sans-serif;font-size:13px;color:#222;padding:32px;max-width:900px;margin:0 auto}
h1{color:#226397;font-size:22px;margin-bottom:4px}
h2{color:#226397;font-size:15px;margin:24px 0 10px;border-bottom:2px solid #e3f0fa;padding-bottom:5px}
h3{font-size:13px;color:#555;margin-bottom:8px}
.subtitle{color:#888;font-size:12px;margin-bottom:24px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px}
.kpi{background:#f8fafc;border:1px solid #dde;border-radius:8px;padding:14px;text-align:center}
.kpi-val{font-size:28px;font-weight:700;color:#226397}
.kpi-lbl{font-size:11px;color:#888;margin-top:4px}
.kpi.warn .kpi-val{color:#e65100}
table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px}
th{background:#226397;color:#fff;padding:8px 10px;text-align:left;font-weight:400}
td{padding:7px 10px;border-bottom:1px solid #eee}
tr:nth-child(even) td{background:#f8fafc}
.bd{display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700}
.bg{background:#e8f5e9;color:#1b5e20}.ba{background:#fff3e0;color:#e65100}
.bt{background:#e3f0fa;color:#1565c0}.br{background:#fce4ec;color:#c62828}
.bgy{background:#eee;color:#555}.bpr{background:#ede7f6;color:#4527a0}
.pipeline{display:flex;border-radius:6px;overflow:hidden;height:24px;margin-bottom:8px}
.ps{display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:600;color:#fff}
.footer{margin-top:32px;padding-top:16px;border-top:1px solid #dde;font-size:11px;color:#aaa;text-align:center}
@media print{body{padding:20px}.no-print{display:none}}
</style>
</head><body>
<div class="no-print" style="margin-bottom:16px">
  <button onclick="window.print()" style="background:#226397;color:#fff;border:none;border-radius:6px;padding:8px 18px;font-size:13px;cursor:pointer">🖨 Print / Save as PDF</button>
</div>
<h1>StudyTrust Governance Report</h1>
<div class="subtitle">000-601 NPO · IT3895/11(T) · Prepared ${today}</div>

<div class="kpi-grid">
  <div class="kpi"><div class="kpi-val">${polData.length+sopData.length}</div><div class="kpi-lbl">Total Registered</div></div>
  <div class="kpi"><div class="kpi-val">${published}</div><div class="kpi-lbl">Published Policies</div></div>
  <div class="kpi ${overdue.length?'warn':''}"><div class="kpi-val">${overdue.length}</div><div class="kpi-lbl">Overdue Reviews</div></div>
  <div class="kpi"><div class="kpi-val">${matPct}%</div><div class="kpi-lbl">Publication Rate</div></div>
</div>

<h2>Lifecycle Pipeline</h2>
<div class="pipeline">
  ${statuses.map((s,i)=>{
    const n=counts[s]||0; if(!n) return '';
    const total=polData.length+sopData.length||1;
    const pct=Math.max(Math.round(n/total*100),4);
    const cols=['#4527a0','#1565c0','#0277bd','#01579b','#2e7d32','#1b5e20','#e65100','#757575'];
    return `<div class="ps" style="width:${pct}%;background:${cols[i]}" title="${s}: ${n}">${n}</div>`;
  }).join('')}
</div>
<div style="display:flex;flex-wrap:wrap;gap:10px;font-size:11px;color:#666;margin-bottom:16px">
  ${statuses.map((s,i)=>counts[s]?`<span>${s}: ${counts[s]}</span>`:'').filter(Boolean).join(' · ')}
</div>

<h2>Published Policies by Category</h2>
<table><tr><th>Category</th><th>Published</th><th>SOPs</th></tr>
${['Governance','Personnel','Finance','IT','Bursaries','Funders','Operations'].map(cat=>{
  const pols=polData.filter(p=>(p.status==='Published'||p.status==='Review Due')&&p.category===cat).length;
  const sops=sopData.filter(s=>(s.status==='Published'||s.status==='Review Due')&&s.category===cat).length;
  return pols||sops?`<tr><td>${cat}</td><td>${pols}</td><td>${sops}</td></tr>`:'';
}).join('')}
</table>

${hiRisk.length?`<h2>High Strategic Policies (Board Approval Required)</h2>
<table><tr><th>ID</th><th>Title</th><th>Status</th><th>Review Date</th><th>Approver</th></tr>
${hiRisk.map(p=>`<tr><td>${esc(p.polId)}</td><td>${esc(p.title)}</td><td>${esc(p.status)}</td><td>${esc(p.reviewDate||'—')}</td><td>${esc(p.approver||'—')}</td></tr>`).join('')}
</table>`:''}

${overdue.length?`<h2>Overdue Reviews (${overdue.length} items)</h2>
<table><tr><th>ID</th><th>Title</th><th>Type</th><th>Due Date</th><th>Owner</th></tr>
${overdue.map(d=>`<tr><td>${esc(d.polId||d.sopId)}</td><td>${esc(d.title)}</td><td>${d.polId?'Policy':'SOP'}</td><td style="color:#e65100;font-weight:600">${esc(d.reviewDate)}</td><td>${esc(d.owner||'—')}</td></tr>`).join('')}
</table>`:''}

${initiated.length?`<h2>Initiated — Planned Policies Not Yet Drafted (${initiated.length})</h2>
<table><tr><th>ID</th><th>Title</th><th>Category</th><th>Owner</th><th>Notes</th></tr>
${initiated.map(d=>`<tr><td>${esc(d.polId||d.sopId)}</td><td>${esc(d.title)}</td><td>${esc(d.category||'—')}</td><td>${esc(d.owner||'—')}</td><td>${esc(d.notes||'—')}</td></tr>`).join('')}
</table>`:''}

<div class="footer">Generated by StudyTrust Governance Hub v4.5 · ${today} · Confidential</div>
</body></html>`);
  rpt.document.close();
}

// ── VERSION HISTORY IN EDIT MODAL ────────────────────────────────
function loadVersionHistory(docId, type){
  const vh=document.getElementById('edit-ver-hist'); if(!vh) return;
  const entries=(logData||[]).filter(l=>(l.docId||l.polId||l.sopId)===docId||l.title===docId);
  if(!entries.length){vh.innerHTML='<div style="font-size:12px;color:var(--tx3);padding:8px 0">No changelog entries found for this document.</div>';return;}
  vh.innerHTML=entries.slice().reverse().slice(0,20).map(l=>`
    <div class="vh-row">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <span><strong>${esc(l.fromStatus||'—')}</strong> → <strong>${esc(l.toStatus||'—')}</strong></span>
        <span class="vh-date">${esc(l.date||l.changed||'')}</span>
      </div>
      ${l.note||l.notes?'<div style="color:var(--tx3);font-size:11px;margin-top:3px">'+esc(l.note||l.notes)+'</div>':''}
      ${l.changedBy||l.user?'<div style="color:var(--tx3);font-size:11px">By: '+esc(l.changedBy||l.user)+'</div>':''}
    </div>`).join('');
}

// Patch openEditModal to load version history
// openEditModal extension handled via patch below

// ── TEMPLATES DOWNLOAD ────────────────────────────────────────────
function openTemplateFolder(){
  const url=localStorage.getItem('st_tpl_folder_url');
  if(url){ window.open(url,'_blank'); return; }
  closeM('m-templates');
  toast('Please configure the Templates folder URL in the Templates modal first.','warning');
  openM('m-templates');
}
function saveTemplateUrl(){
  const url=(document.getElementById('tpl-folder-url')||{}).value||'';
  localStorage.setItem('st_tpl_folder_url',url);
  toast('Template folder URL saved.','success');
}
function initTemplateUrl(){
  const url=localStorage.getItem('st_tpl_folder_url')||'';
  const el=document.getElementById('tpl-folder-url');
  if(el) el.value=url;
  if(url){
    ['tpl-pol-link','tpl-sop-link','tpl-ann-link'].forEach(id=>{
      const a=document.getElementById(id); if(a) a.href=url;
    });
  }
}

// ── POLICY ACKNOWLEDGEMENTS ───────────────────────────────────────
let ackData = [];
const ACK_MAP = { title:'Title', version:'PolicyVersion', userEmail:'UserEmail',
  userName:'UserDisplayName', date:'AcknowledgedDate', polTitle:'PolicyTitle' };

async function loadAckData(){
  try{
    const sid=await getSiteId();
    const data=await gFetch(`/sites/${sid}/lists/${LISTS.ack}/items?$expand=fields&$top=999&$select=id,fields`);
    ackData=(data.value||[]).map(item=>{
      const f=item.fields||{}; const r={_spId:item.id};
      for(const[ak,sk] of Object.entries(ACK_MAP)) r[ak]=f[sk]||'';
      return r;
    });
  } catch(e){ ackData=[]; }
}

async function acknowledgePolicy(polId, version, polTitle){
  if(!polId) return;
  const user=document.getElementById('user-name')?.textContent||'Unknown';
  const email=typeof msalApp!=='undefined'?
    (msalApp.getAllAccounts()[0]?.username||'unknown') : 'unknown';
  try{
    const sid=await getSiteId();
    const fields={
      Title: polId, PolicyVersion: version||'', PolicyTitle: polTitle||'',
      UserEmail: email, UserDisplayName: user,
      AcknowledgedDate: new Date().toISOString().slice(0,10)
    };
    await gFetch(`/sites/${sid}/lists/${LISTS.ack}/items`,{method:'POST',
      headers:{'Content-Type':'application/json'},body:JSON.stringify({fields})});
    toast('Acknowledged: '+polId+' v'+version,'success');
    await loadAckData();
    renderAck();
    renderPendingAck();
  } catch(e){ toast('Acknowledgement failed: '+e.message,'error'); }
}

function renderAck(){
  const body=document.getElementById('body-ack'); if(!body) return;
  const polF=(document.getElementById('ack-pol-filter')||{}).value||'';
  const rows=ackData.filter(a=>!polF||a.title===polF);

  if(!ackData.length){
    body.innerHTML='<tr><td colspan="5" style="text-align:center;padding:24px;color:var(--tx3)">No acknowledgements recorded yet. Requires GovAcknowledgements SharePoint list.</td></tr>';
    document.getElementById('ack-summary').textContent='';
  } else {
    const uniq=[...new Set(ackData.map(a=>a.title))].length;
    const users=[...new Set(ackData.map(a=>a.userEmail))].length;
    document.getElementById('ack-summary').textContent=ackData.length+' acknowledgements · '+uniq+' policies · '+users+' employees';
    body.innerHTML=rows.map((a,i)=>`<tr>
      <td>${esc(a.title)}</td><td>${esc(a.version)}</td>
      <td>${esc(a.userName)}</td><td>${esc(a.userEmail)}</td>
      <td>${esc(a.date)}</td></tr>`).join('');
  }

  // Populate filter dropdown
  const pf=document.getElementById('ack-pol-filter');
  if(pf&&pf.options.length===1){
    [...new Set(ackData.map(a=>a.title))].sort().forEach(id=>{
      const o=new Option(id,id); pf.appendChild(o);
    });
  }
}

function renderPendingAck(){
  const sec=document.getElementById('ack-pending-section');
  const list=document.getElementById('ack-pending-list');
  if(!sec||!list) return;
  if(currentRole!=='personnel'&&currentRole!=='trustee'){ sec.style.display='none'; return; }
  const email=typeof msalApp!=='undefined'?
    (msalApp.getAllAccounts()[0]?.username||'') : '';
  const myAcked=new Set(ackData.filter(a=>a.userEmail===email).map(a=>a.title+'::'+a.version));
  const toAck=polData.filter(p=>(p.status==='Published'||p.status==='Review Due')&&
    !myAcked.has(p.polId+'::'+p.version));
  if(!toAck.length){
    sec.style.display='block';
    list.innerHTML='<div style="padding:16px;text-align:center;color:#388e3c;font-weight:600">✅ You are up to date — all published policies acknowledged.</div>';
    return;
  }
  sec.style.display='block';
  list.innerHTML=toAck.map(p=>`<div class="ack-row">
    <div>
      <div class="ack-pol">${esc(p.polId)} — ${esc(p.title)}</div>
      <div class="ack-meta">${esc(p.category)} · v${esc(p.version)} · Owner: ${esc(p.owner||'—')}</div>
    </div>
    <div style="display:flex;gap:8px;align-items:center;flex-shrink:0">
      ${p.publishedUrl?`<a href="${esc(p.publishedUrl)}" target="_blank" class="btn btn-sec btn-sm">📄 Read</a>`:''}
      <button class="btn btn-pri btn-sm" onclick="acknowledgePolicy('${esc(p.polId)}','${esc(p.version)}','${esc(p.title)}')">✓ Acknowledge</button>
    </div></div>`).join('');
}

// ── CARD VIEW ────────────────────────────────────────────────────
let polView='table', sopView='table';

function setPolView(mode){
  polView=mode;
  document.getElementById('pol-vt-tbl').className='vt-btn'+(mode==='table'?' active':'');
  document.getElementById('pol-vt-card').className='vt-btn'+(mode==='card'?' active':'');
  renderPol();
}
function setSopView(mode){
  sopView=mode;
  document.getElementById('sop-vt-tbl').className='vt-btn'+(mode==='table'?' active':'');
  document.getElementById('sop-vt-card').className='vt-btn'+(mode==='card'?' active':'');
  renderSop();
}

function renderPolCards(data){
  const body=document.getElementById('body-pol');
  const tbl=document.querySelector('#sec-pol .tw');
  if(tbl) tbl.style.display='none';
  let grid=document.getElementById('pol-card-grid');
  if(!grid){
    grid=document.createElement('div'); grid.id='pol-card-grid';
    tbl?.parentNode.appendChild(grid);
  }
  grid.style.display='grid';
  grid.className='pol-cards';
  grid.innerHTML=data.map(p=>`<div class="pc" onclick="openEditModal('${esc(p.polId)}','pol')">
    <div class="pc-id">${esc(p.polId)}</div>
    <div class="pc-title">${esc(p.title)}</div>
    <div class="pc-badges">${sbadge(p.status)}${rbadge(p.risk)}${tbadge(p.tier)}</div>
    <div class="pc-meta">
      <span>👤 ${esc(p.owner||'—')}</span>
      <span>📋 ${esc(p.category||'—')}</span>
      ${p.reviewDate?'<span>📅 Review: '+esc(p.reviewDate)+'</span>':''}
    </div>
  </div>`).join('');
}

function renderSopCards(data){
  const tbl=document.querySelector('#sec-sop .tw');
  if(tbl) tbl.style.display='none';
  let grid=document.getElementById('sop-card-grid');
  if(!grid){
    grid=document.createElement('div'); grid.id='sop-card-grid';
    tbl?.parentNode.appendChild(grid);
  }
  grid.style.display='grid';
  grid.className='sop-cards';
  grid.innerHTML=data.map(s=>`<div class="pc" onclick="openEditModal('${esc(s.sopId)}','sop')">
    <div class="pc-id">${esc(s.sopId)}</div>
    <div class="pc-title">${esc(s.title)}</div>
    <div class="pc-badges">${sbadge(s.status)}</div>
    <div class="pc-meta">
      <span>👤 ${esc(s.owner||'—')}</span>
      <span>📋 ${esc(s.category||'—')}</span>
      ${s.reviewDate?'<span>📅 Review: '+esc(s.reviewDate)+'</span>':''}
    </div>
  </div>`).join('');
}

// ── XLSX EXPORT ───────────────────────────────────────────────────
function exportXlsx(tab){
  if(typeof XLSX==='undefined'){
    toast('XLSX library not loaded. Check network connection.','error'); return;
  }
  const wb=XLSX.utils.book_new();
  if(tab==='pol'||tab==='all'){
    const polRows=[
      ['Policy ID','Title','Category','Tier','Version','Status','Risk','Approver','Review Date','Custodian','Owner','Notes'],
      ...polData.map(p=>[p.polId,p.title,p.category,p.tier,p.version,p.status,p.risk,p.approver,p.reviewDate,p.custodian,p.owner,p.notes])
    ];
    const ws=XLSX.utils.aoa_to_sheet(polRows);
    ws['!cols']=[8,30,14,6,8,14,18,20,12,20,20,30].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,'Policies');
  }
  if(tab==='sop'||tab==='all'){
    const sopRows=[
      ['SOP ID','Title','Category','Linked Policy','Version','Status','Review Date','Custodian','Owner'],
      ...sopData.map(s=>[s.sopId,s.title,s.category,s.linkedPolicy,s.version,s.status,s.reviewDate,s.custodian,s.owner])
    ];
    const ws=XLSX.utils.aoa_to_sheet(sopRows);
    ws['!cols']=[8,30,14,14,8,14,12,20,20].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,'SOPs');
  }
  if(tab==='ack'){
    const ackRows=[
      ['Policy ID','Policy Title','Version','Acknowledged By','Email','Date'],
      ...ackData.map(a=>[a.title,a.polTitle,a.version,a.userName,a.userEmail,a.date])
    ];
    const ws=XLSX.utils.aoa_to_sheet(ackRows);
    ws['!cols']=[10,30,8,25,30,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,'Acknowledgements');
  }
  if(!wb.SheetNames.length){ toast('No data to export.','warning'); return; }
  const date=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`StudyTrust_Governance_${tab.toUpperCase()}_${date}.xlsx`);
  toast('XLSX exported.','success');
}

// ── INIT (called after loadData) ──────────────────────────────────
// loadData extension handled inline in loadAndRender patch below

// ════════════════════════════════════════════════════════════════
// END v4.5
// ════════════════════════════════════════════════════════════════

function allowedTransitions(currentStatus, role){
  const all = STATUS_FLOW[currentStatus] || [];
  if(role === 'admin') return all;
  if(role === 'ptt')   return all.filter(s => s !== 'Published' && s !== 'Archived');
  if(role === 'legal') return all.filter(s => s !== 'Published' && s !== 'Archived');
  return [];
}

function buildTabs(){
  const tabs=currentRole==='admin'?TABS_ADMIN:currentRole==='trustee'?TABS_TRUSTEE:currentRole==='auditor'?TABS_AUDITOR:currentRole==='ptt'?TABS_PTT:currentRole==='legal'?TABS_PTT:TABS_PERSONNEL;
  document.getElementById('tab-bar').innerHTML=tabs.map((t,i)=>`<button class="tab${i===0?' on':''}" onclick="showTab('${t.id}',this)">${t.l}</button>`).join('');
}

function showTab(id,btn){
  document.querySelectorAll('.sec').forEach(s=>s.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('on'));
  const sec=document.getElementById(id); if(sec) sec.classList.add('on');
  if(btn) btn.classList.add('on');
  if(id==='sec-training') loadTrainingTab();
}

function setRole(role){
  currentRole=role;
  document.getElementById('rb-a').classList.toggle('on',role==='admin');
  document.getElementById('rb-t').classList.toggle('on',role==='trustee');
  document.getElementById('rb-au').classList.toggle('on',role==='auditor');
  document.getElementById('rb-p').classList.toggle('on',role==='personnel');
  const rbpt=document.getElementById('rb-pt'); if(rbpt) rbpt.classList.toggle('on',role==='ptt');

  const badge=document.getElementById('role-badge');
  const labels={admin:'Admin',trustee:'Trustee',auditor:'External Auditor',personnel:'All Personnel',ptt:'PTT (Policy Task Team)',legal:'Legal Reviewer'};
  const badgeCls={admin:'rb-admin',trustee:'rb-trustee',auditor:'rb-auditor',personnel:'rb-personnel',ptt:'rb-ptt',legal:'rb-ptt'};
  badge.textContent=labels[role]||'All Personnel';
  badge.className='role-badge '+(badgeCls[role]||'rb-personnel');

  const banner=document.getElementById('role-banner');
  const bannerCls={admin:'ban-admin',trustee:'ban-trustee',auditor:'ban-auditor',personnel:'ban-personnel',ptt:'ban-ptt',legal:'ban-legal'};
  const bannerText={
    admin:'Admin Mode — Full access. All changes save live to SharePoint.',
    trustee:'Trustee View — Dashboards - Board-of-Trustees access. Published policies and governance information.',
    auditor:'External Auditor View — Read-only access to published documents. Use Document Requests to request additional information.',
    personnel:'All Personnel View — Published policies and SOPs only.',
    ptt:'PTT View — Review and edit drafts (Step 4 of the governance process). You cannot publish, archive, or change hub structure.',
    legal:'Legal Reviewer View — Review and annotate drafts. You cannot publish, archive, or change hub structure.'
  };
  banner.className=bannerCls[role]||'ban-personnel';
  banner.textContent=bannerText[role]||bannerText.personnel;

  document.querySelectorAll('.adm').forEach(el=>el.classList.toggle('show',role==='admin'));
  document.querySelectorAll('.adm-b').forEach(el=>el.classList.toggle('show',role==='admin'));
  document.querySelectorAll('.adm-r').forEach(el=>el.classList.toggle('show',role==='admin'));

  buildTabs();
  document.querySelectorAll('.sec').forEach(s=>s.classList.remove('on'));
  const firstTab={admin:'sec-exec',trustee:'sec-exec',auditor:'sec-aud-docs',personnel:'sec-pol',ptt:'sec-exec'}[role]||'sec-pol';
  const sec=document.getElementById(firstTab); if(sec) sec.classList.add('on');

  renderAll();
  if(role==='admin') polDropdown();
  if(role==='auditor'){ renderAudPol(); renderAudSop(); renderAudOverview(); renderMyReqs(); }
}

// ════════════════════════════════════════════
// MODALS
// ════════════════════════════════════════════
function openM(id){document.getElementById(id).classList.add('show');}
function closeM(id){document.getElementById(id).classList.remove('show');}
document.querySelectorAll('.modal-ov').forEach(el=>el.addEventListener('click',e=>{if(e.target===el)el.classList.remove('show');}));

// ════════════════════════════════════════════
// LOAD / TOAST
// ════════════════════════════════════════════
let _lt;
function setLoad(msg,sub){
  document.getElementById('load-msg').textContent=msg||'Loading…';
  const s=document.getElementById('load-sub'); if(s) s.textContent=sub||'';
  document.getElementById('load').classList.remove('hidden');
  const d=document.getElementById('load-dismiss'); if(d) d.style.display='none';
  clearTimeout(_lt);
  _lt=setTimeout(()=>{const d=document.getElementById('load-dismiss');if(d)d.style.display='inline-block';},8000);
}
function hideLoad(){clearTimeout(_lt);document.getElementById('load').classList.add('hidden');const d=document.getElementById('load-dismiss');if(d)d.style.display='none';}

let _tt;
function toast(msg,type=''){
  const t=document.getElementById('toast'); t.textContent=msg; t.className='toast '+(type?type:'');
  t.classList.add('show'); clearTimeout(_tt); _tt=setTimeout(()=>t.classList.remove('show'),4500);
}

// ════════════════════════════════════════════
// AUDITOR RENDER FUNCTIONS
// ════════════════════════════════════════════
let audReqs=[]; // stored in localStorage only — no SP list needed

function renderAudPol(){
  const q=(document.getElementById('aud-pol-q')||{}).value||'';
  const cf=(document.getElementById('aud-pol-cat')||{}).value||'';
  const ql=q.toLowerCase();
  const data=polData.filter(p=>isPublished(p.status));
  const tbody=document.getElementById('aud-pol-body'); if(!tbody) return;
  let html='',shown=0;
  data.forEach(p=>{
    const txt=[p.polId,p.title,p.category,p.tier,p.risk,p.version,p.reviewDate,p.approver].join(' ').toLowerCase();
    if(ql&&txt.indexOf(ql)===-1) return;
    if(cf&&p.category!==cf) return;
    shown++;
    const viewBtn=p.publishedUrl?`<a class="doc-btn view" href="${esc(p.publishedUrl)}" target="_blank">📄 View</a>`:'<span style="font-size:11px;color:var(--tx3)">Not yet available</span>';
    html+=`<tr><td>${esc(p.polId)}</td><td>${esc(p.title)}</td><td>${esc(p.category)}</td><td>${tbadge(p.tier)}</td><td>${rbadge(p.risk)}</td><td>${esc(p.version)}</td><td>${datecell(p.reviewDate)}</td><td style="font-size:11px">${esc(p.approver)}</td><td>${viewBtn}</td></tr>`;
  });
  if(!shown) html=`<tr><td colspan="9" style="text-align:center;color:var(--tx3);padding:20px">No published policies match the current filters.</td></tr>`;
  tbody.innerHTML=html;
  // Populate category filter
  const cats=[...new Set(data.map(p=>p.category).filter(Boolean))].sort();
  const cf2=document.getElementById('aud-pol-cat');
  if(cf2&&cf2.options.length<=1) cats.forEach(c=>{const o=new Option(c,c);cf2.appendChild(o);});
}

function renderAudSop(){
  const q=(document.getElementById('aud-sop-q')||{}).value||'';
  const cf=(document.getElementById('aud-sop-cat')||{}).value||'';
  const ql=q.toLowerCase();
  const data=sopData.filter(s=>isPublished(s.status));
  const tbody=document.getElementById('aud-sop-body'); if(!tbody) return;
  let html='',shown=0;
  data.forEach(s=>{
    const txt=[s.sopId,s.title,s.category,s.linkedPolicy,s.version,s.reviewDate,s.owner].join(' ').toLowerCase();
    if(ql&&txt.indexOf(ql)===-1) return;
    if(cf&&s.category!==cf) return;
    shown++;
    const viewBtn=s.publishedUrl?`<a class="doc-btn view" href="${esc(s.publishedUrl)}" target="_blank">📄 View</a>`:'<span style="font-size:11px;color:var(--tx3)">Not yet available</span>';
    html+=`<tr><td>${esc(s.sopId)}</td><td>${esc(s.title)}</td><td>${esc(s.category)}</td><td>${esc(s.linkedPolicy)}</td><td>${esc(s.version)}</td><td>${datecell(s.reviewDate)}</td><td style="font-size:11px">${esc(s.owner)}</td><td>${viewBtn}</td></tr>`;
  });
  if(!shown) html=`<tr><td colspan="8" style="text-align:center;color:var(--tx3);padding:20px">No published SOPs match the current filters.</td></tr>`;
  tbody.innerHTML=html;
  const cats=[...new Set(data.map(s=>s.category).filter(Boolean))].sort();
  const cf2=document.getElementById('aud-sop-cat');
  if(cf2&&cf2.options.length<=1) cats.forEach(c=>{const o=new Option(c,c);cf2.appendChild(o);});
}

function renderAudOverview(){
  const totalP=polData.length, pubP=polData.filter(p=>isPublished(p.status)).length;
  const totalS=sopData.length, pubS=sopData.filter(s=>isPublished(s.status)).length;
  const overP=polData.filter(p=>isOverdue(p.reviewDate)&&p.status!=='Archived').length;
  const overS=sopData.filter(s=>isOverdue(s.reviewDate)&&s.status!=='Archived').length;

  const kpig=document.getElementById('aud-kpi-grid');
  if(kpig) kpig.innerHTML=`
    <div class="kpi"><div class="kpi-l">Published Policies</div><div class="kpi-v" style="color:var(--st)">${pubP}</div><div class="kpi-s">of ${totalP} total</div></div>
    <div class="kpi"><div class="kpi-l">Published SOPs</div><div class="kpi-v" style="color:var(--st)">${pubS}</div><div class="kpi-s">of ${totalS} total</div></div>
    <div class="kpi"><div class="kpi-l">Overdue Reviews</div><div class="kpi-v" style="color:${overP+overS>0?'var(--rd-tx)':'var(--gn-tx)'}">${overP+overS}</div><div class="kpi-s">${overP} policies · ${overS} SOPs</div></div>
    <div class="kpi"><div class="kpi-l">My Requests</div><div class="kpi-v" style="color:#065f46">${audReqs.length}</div><div class="kpi-s">${audReqs.filter(r=>r.status==='Pending').length} pending</div></div>`;

  // Status summary table
  const statuses=['Published','Approved','New','Review Due','Draft','PTT Review','Consult','Archived'];
  const allDocs=[...polData,...sopData];
  const sb=document.getElementById('aud-status-body');
  if(sb) sb.innerHTML=statuses.map(s=>{
    const n=allDocs.filter(d=>d.status===s).length;
    return n?`<tr><td>${sbadge(s)}</td><td style="font-weight:600">${n}</td><td style="color:var(--tx2)">${Math.round(n/allDocs.length*100)}%</td></tr>`:'';
  }).join('');

  // Overdue table
  const ob=document.getElementById('aud-overdue-body');
  const overdueItems=[...polData.filter(p=>isOverdue(p.reviewDate)&&p.status!=='Archived').map(p=>({id:p.polId,title:p.title,type:'Policy',date:p.reviewDate,owner:p.owner})),...sopData.filter(s=>isOverdue(s.reviewDate)&&s.status!=='Archived').map(s=>({id:s.sopId,title:s.title,type:'SOP',date:s.reviewDate,owner:s.owner}))];
  if(ob) ob.innerHTML=overdueItems.length?overdueItems.map(d=>`<tr><td style="font-size:12px">${esc(d.id)} — ${esc(d.title)}</td><td><span class="bd ${d.type==='Policy'?'bb':'bgy'}">${d.type}</span></td><td class="overdue">${esc(d.date)} ⚠</td><td style="font-size:11px">${esc(d.owner)}</td></tr>`).join(''):'<tr><td colspan="4" style="text-align:center;color:var(--gn-tx);padding:16px">No overdue reviews.</td></tr>';

  // Pipeline (auditor sees published only)
  const pp=document.getElementById('aud-pipeline');
  if(pp){
    const stages=[{label:'DRAFT',statuses:['Draft','Consult','PTT Review'],color:'#90a4ae',bg:'#eceff1'},{label:'APPROVED',statuses:['Approved','New'],color:'#388e3c',bg:'#e8f5e9'},{label:'PUBLISHED',statuses:['Published'],color:'#2e7d32',bg:'#c8e6c9'},{label:'REVIEW DUE',statuses:['Review Due'],color:'#f57f17',bg:'#fff8e1'},{label:'ARCHIVED',statuses:['Archived'],color:'#757575',bg:'#f5f5f5'}];
    pp.innerHTML=stages.map(st=>{const n=allDocs.filter(d=>st.statuses.includes(d.status)).length;return`<div class="ps" style="background:${st.bg};border:1px solid ${st.color}20"><div class="ps-l" style="color:${st.color}">${st.label}</div><div class="ps-v" style="color:${st.color}">${n}</div></div>`;}).join('<div style="display:flex;align-items:center;color:#bbb;font-size:14px;padding:0 2px">›</div>');
  }
}

function submitAudRequest(){
  const name=document.getElementById('aud-req-name').value.trim();
  const org=document.getElementById('aud-req-org').value.trim();
  const doc=document.getElementById('aud-req-doc').value.trim();
  const urgency=document.getElementById('aud-req-urgency').value;
  const note=document.getElementById('aud-req-note').value.trim();
  if(!doc){toast('Please specify the document or information you need.','warning');return;}
  const req={id:Date.now(),date:td(),name:name||'Auditor',org:org||'External',doc,urgency,note,status:'Pending'};
  audReqs.unshift(req);
  localStorage.setItem('st_aud_reqs',JSON.stringify(audReqs));
  document.getElementById('aud-req-name').value='';
  document.getElementById('aud-req-doc').value='';
  document.getElementById('aud-req-note').value='';
  renderMyReqs();
  renderAudOverview();
  toast('Request submitted. The governance team will respond via email or SharePoint.','success');
  // Notify admin team via webhook
  const wh = loadWebhooks();
  if(wh.admin){
    sendAdaptiveCard(wh.admin,
      '📋 New Auditor Document Request',
      `From: ${name||'Auditor'} (${org||'External'})`,
      `**Document requested:** ${doc}\n**Urgency:** ${urgency}${note ? '\n**Context:** '+note : ''}`,
      'admin'
    );
  }
  // Push to in-app bell
  pushNotif('request', `Document request from ${name||'Auditor'}`, `${doc} · Urgency: ${urgency}`);
}

function renderMyReqs(){
  audReqs=JSON.parse(localStorage.getItem('st_aud_reqs')||'[]');
  const el=document.getElementById('aud-req-list'); if(!el) return;
  if(!audReqs.length){el.innerHTML='<div style="color:var(--tx3);font-size:12px;padding:12px">No requests submitted yet.</div>';return;}
  el.innerHTML='<div class="tbl-wrap"><table class="dt" style="width:100%"><thead><tr>'
    +['Date','Document Requested','Urgency','Status','Your Notes'].map(h=>`<th style="background:#065f46">${h}</th>`).join('')
    +'</tr></thead><tbody>'
    +audReqs.map(r=>`<tr>
      <td style="white-space:nowrap;color:var(--tx2)">${esc(r.date)}</td>
      <td>${esc(r.doc)}</td>
      <td><span class="bd ${r.urgency==='Urgent'?'br':r.urgency==='High'?'ba':'bgy'}">${esc(r.urgency)}</span></td>
      <td><span class="bd ${r.status==='Resolved'?'bg':'ba'}">${esc(r.status)}</span></td>
      <td style="font-size:11px;color:var(--tx2)">${esc(r.note||'—')}</td>
    </tr>`).join('')
    +'</tbody></table></div>';
}

// ════════════════════════════════════════════
// NOTIFICATION SYSTEM
// ════════════════════════════════════════════
const NOTIF_KEY   = 'st_gov_notifs';
const SEEN_KEY    = 'st_gov_last_seen';
const WEBHOOK_KEY = 'st_gov_webhooks';

// Each notification: { id, type, title, meta, ts, read }
// Types: 'new-doc', 'status-change', 'overdue', 'request'

function loadNotifs(){
  try{ return JSON.parse(localStorage.getItem(NOTIF_KEY)||'[]'); }
  catch(e){ return []; }
}
function saveNotifs(arr){
  // Keep at most 100 notifications
  localStorage.setItem(NOTIF_KEY, JSON.stringify(arr.slice(0,100)));
}

function pushNotif(type, title, meta){
  const arr = loadNotifs();
  arr.unshift({ id: Date.now()+Math.random(), type, title, meta, ts: new Date().toISOString(), read: false });
  saveNotifs(arr);
  renderNotifBell();
}

function renderNotifBell(){
  const arr   = loadNotifs();
  const unread = arr.filter(n => !n.read);
  const nc    = document.getElementById('notif-count');
  const btn   = document.getElementById('notif-btn');
  if(!nc) return;
  if(unread.length > 0){
    nc.style.display = 'flex';
    nc.textContent   = unread.length > 99 ? '99+' : unread.length;
  } else {
    nc.style.display = 'none';
  }
  if(btn) btn.title = unread.length ? unread.length+' unread notification'+(unread.length>1?'s':'') : 'No new notifications';
}

function renderNotifPanel(){
  const arr  = loadNotifs();
  const list = document.getElementById('notif-list');
  const ttl  = document.getElementById('notif-title');
  if(!list) return;
  const unread = arr.filter(n => !n.read).length;
  if(ttl) ttl.textContent = 'Notifications' + (unread ? ' ('+unread+' new)' : '');

  const typeStyle = {
    'new-doc'      : { cls:'ni-new',    icon:'📄', label:'NEW' },
    'status-change': { cls:'ni-change',  icon:'🔄', label:'STATUS' },
    'overdue'      : { cls:'ni-late',    icon:'⚠',  label:'OVERDUE' },
    'request'      : { cls:'ni-req',     icon:'📋', label:'REQUEST' }
  };

  if(!arr.length){
    list.innerHTML = '<div class="notif-empty">No notifications yet.</div>';
    return;
  }

  list.innerHTML = arr.slice(0, 30).map(n => {
    const s = typeStyle[n.type] || { cls:'ni-change', icon:'•', label:'INFO' };
    const d = new Date(n.ts);
    const when = d.toLocaleDateString('en-ZA') + ' ' + d.toLocaleTimeString('en-ZA', {hour:'2-digit',minute:'2-digit'});
    return `<div class="ni" style="${n.read?'opacity:.6':''}">
      <div class="ni-title"><span class="ni-badge ${s.cls}">${s.icon} ${s.label}</span>${esc(n.title)}</div>
      <div class="ni-meta">${esc(n.meta)} &nbsp;·&nbsp; ${when}</div>
    </div>`;
  }).join('');
}

function toggleNotif(){
  const panel = document.getElementById('notif-panel');
  if(!panel) return;
  const isOpen = panel.classList.contains('open');
  panel.classList.toggle('open');
  if(!isOpen) renderNotifPanel();
}

function markAllRead(){
  const arr = loadNotifs().map(n => ({...n, read:true}));
  saveNotifs(arr);
  renderNotifBell();
  renderNotifPanel();
}

// Scan for overdue items and push notifications if they are new since last check
function checkOverdueNotifs(){
  const lastKey = 'st_gov_overdue_checked';
  const lastCheck = localStorage.getItem(lastKey) || '';
  const today = td();
  if(lastCheck === today) return; // Already checked today
  localStorage.setItem(lastKey, today);

  const all = [...polData, ...sopData];
  const overItems = all.filter(d => isOverdue(d.reviewDate) && d.status !== 'Archived');
  overItems.forEach(d => {
    const id = d.polId || d.sopId;
    pushNotif('overdue', `${id} — ${d.title||d.polId||d.sopId}`, `Review was due ${d.reviewDate} · Owner: ${d.owner||'—'}`);
  });
  if(overItems.length){
    // Optionally send to admin webhook
    const wh = loadWebhooks();
    if(wh.admin && loadWebhookPrefs().notifyOverdue){
      sendAdaptiveCard(wh.admin, '⚠ Overdue Governance Reviews', `${overItems.length} document(s) are overdue for review.`,
        overItems.slice(0,5).map(d=>`• ${d.polId||d.sopId}: due ${d.reviewDate}`).join('\n'),
        'all');
    }
  }
}

// ════════════════════════════════════════════
// WEBHOOK MANAGEMENT
// ════════════════════════════════════════════
function loadWebhooks(){
  try{ return JSON.parse(localStorage.getItem(WEBHOOK_KEY)||'{}'); }
  catch(e){ return {}; }
}
function loadWebhookPrefs(){
  try{ return JSON.parse(localStorage.getItem(WEBHOOK_KEY+'_prefs')||'{}'); }
  catch(e){ return {}; }
}

function saveWebhooks(){
  const wh = {
    admin:   (document.getElementById('wh-admin')  ||{}).value||'',
    trustee: (document.getElementById('wh-trustee')||{}).value||'',
    ptt:     (document.getElementById('wh-ptt')    ||{}).value||'',
    auditor: (document.getElementById('wh-auditor')||{}).value||''
  };
  const prefs = {
    notifyOverdue: (document.getElementById('wh-notify-overdue')||{}).checked||false
  };
  localStorage.setItem(WEBHOOK_KEY, JSON.stringify(wh));
  localStorage.setItem(WEBHOOK_KEY+'_prefs', JSON.stringify(prefs));
  updateWebhookDots(wh);
  toast('Webhook URLs saved.', 'success');
}

function loadWebhookUI(){
  const wh = loadWebhooks();
  const prefs = loadWebhookPrefs();
  const a = document.getElementById('wh-admin');
  const t = document.getElementById('wh-trustee');
  const p = document.getElementById('wh-ptt');
  const au = document.getElementById('wh-auditor');
  const ov = document.getElementById('wh-notify-overdue');
  if(a) a.value = wh.admin||'';
  if(t) t.value = wh.trustee||'';
  if(p) p.value = wh.ptt||'';
  if(au) au.value = wh.auditor||'';
  if(ov) ov.checked = prefs.notifyOverdue||false;
  updateWebhookDots(wh);
}

function updateWebhookDots(wh){
  const dots = { admin:'wh-admin-dot', trustee:'wh-trustee-dot', ptt:'wh-ptt-dot', auditor:'wh-auditor-dot' };
  Object.entries(dots).forEach(([key, id]) => {
    const el = document.getElementById(id);
    if(!el) return;
    el.className = 'wh-dot ' + (wh[key] ? 'wh-ok' : 'wh-empty');
  });
}

async function testWebhook(audience){
  const wh = loadWebhooks();
  const url = wh[audience];
  if(!url){ toast('No webhook URL configured for ' + audience + '.', 'warning'); return; }
  const ok = await sendAdaptiveCard(url,
    '🔔 StudyTrust Governance Hub — Test Notification',
    'This is a test notification from the StudyTrust Governance Hub.',
    `Audience: ${audience} · Sent by: ${userName} · ${new Date().toLocaleString('en-ZA')}`,
    audience
  );
  toast(ok ? 'Test notification sent to ' + audience + ' channel ✓' : 'Webhook test failed — check the URL.', ok ? 'success' : 'error');
}

// ── Send an Adaptive Card to a Teams Incoming Webhook ──
async function sendAdaptiveCard(webhookUrl, title, subtitle, body, audience){
  if(!webhookUrl) return false;
  // Teams Incoming Webhook uses the MessageCard format (Adaptive Cards require approval in some tenants)
  const card = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "226397",
    "summary": title,
    "sections": [{
      "activityTitle": `**${title}**`,
      "activitySubtitle": subtitle,
      "activityText": body,
      "facts": [
        { "name": "System:", "value": "StudyTrust Governance Hub" },
        { "name": "Audience:", "value": audience === 'all' ? 'Internal Team' : audience.charAt(0).toUpperCase() + audience.slice(1) },
        { "name": "Timestamp:", "value": new Date().toLocaleString('en-ZA') }
      ],
      "markdown": true
    }],
    "potentialAction": [{
      "@type": "OpenUri",
      "name": "Open Governance Hub",
      "targets": [{ "os": "default", "uri": "https://studytrust.github.io/StudyTrust_Governance_Dashboard_v4.html" }]
    }]
  };
  try{
    const r = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(card)
    });
    return r.ok;
  }catch(e){
    console.warn('Webhook error:', e);
    return false;
  }
}

// ── Fire notifications for a status change event ──
// audience rules:
//   admin   → all status changes
//   trustee → only Published, Review Due, Archived
//   auditor → only Published
async function fireStatusChangeNotif(docId, docTitle, fromStatus, toStatus, changedBy){
  const wh = loadWebhooks();
  const title  = `${docId} — Status changed to ${toStatus}`;
  const sub    = `${docTitle}`;
  const body   = `**From:** ${fromStatus}  →  **To:** ${toStatus}\n**Changed by:** ${changedBy}`;

  // Push to in-app bell for current user
  pushNotif('status-change', title, `${fromStatus} → ${toStatus} · by ${changedBy}`);

  const trusteeStatuses = ['Published','Review Due','Archived'];
  const auditorStatuses = ['Published'];
  // PTT is interested in drafts moving through review-relevant states
  const pttStatuses     = ['Consult','PTT Review'];

  const sends = [];
  if(wh.admin)   sends.push(sendAdaptiveCard(wh.admin,   '🔄 ' + title, sub, body, 'admin'));
  if(wh.trustee && trusteeStatuses.includes(toStatus))
                  sends.push(sendAdaptiveCard(wh.trustee, '🔄 ' + title, sub, body, 'trustee'));
  if(wh.ptt && pttStatuses.includes(toStatus)){
    const pttBody = toStatus==='PTT Review'
      ? `${docTitle} is ready for PTT review. Please open the Governance Hub and review the draft.`
      : `${docTitle} has entered consultation. Please review the draft and provide input.`;
    sends.push(sendAdaptiveCard(wh.ptt, '📝 ' + title, sub, pttBody, 'ptt'));
  }
  if(wh.auditor && auditorStatuses.includes(toStatus))
                  sends.push(sendAdaptiveCard(wh.auditor, '📄 New Published Document: ' + docId, sub, `${docTitle} has been published and is now available for review.`, 'auditor'));

  await Promise.allSettled(sends);
}

// ── Fire notifications for a new document registered ──
async function fireNewDocNotif(docId, docTitle, type, status, registeredBy){
  const wh = loadWebhooks();
  const title = `New ${type} registered: ${docId}`;
  const body  = `**${docTitle}**\nInitial status: ${status} · Registered by: ${registeredBy}`;

  pushNotif('new-doc', title, `Status: ${status} · by ${registeredBy}`);

  const sends = [];
  if(wh.admin) sends.push(sendAdaptiveCard(wh.admin, '📋 ' + title, `A new ${type} has been registered in the Governance Hub.`, body, 'admin'));
  // Notify PTT when a new draft is registered (so they can pick it up for review when it advances)
  if(wh.ptt && ['Draft','Consult','PTT Review'].includes(status)){
    sends.push(sendAdaptiveCard(wh.ptt, '📝 ' + title, `A new ${type} draft has been registered. It will reach PTT review when its drafting and consultation phases complete.`, body, 'ptt'));
  }
  // Auditors and trustees only notified when Published — not for drafts
  await Promise.allSettled(sends);
}

// ════════════════════════════════════════════
// TRUSTEES
// ════════════════════════════════════════════
function renderTrustees(){
  const tbody=document.getElementById('trustees-body'); if(!tbody) return;
  const hdr=document.getElementById('tr-act-hdr');
  if(hdr) hdr.style.display=currentRole==='admin'?'':'none';
  const count=document.getElementById('tr-count');
  const active=trusteesData.filter(t=>t.status==='Active').length;
  if(count) count.textContent=`${trusteesData.length} registered · ${active} active`;

  const statusBadge=(s)=>{
    const m={'Active':'bg','Inactive':'bgy','Pending Invitation':'ba','Invited — Awaiting Acceptance':'bt'};
    return`<span class="bd ${m[s]||'bgy'}">${esc(s)}</span>`;
  };

  if(!trusteesData.length){
    tbody.innerHTML=`<tr><td colspan="8" style="text-align:center;color:var(--tx3);padding:24px">No trustees registered yet. Use the form above to register Board members.</td></tr>`;
    return;
  }

  tbody.innerHTML=trusteesData.map(t=>{
    const act=currentRole==='admin'
      ?`<td><button class="btn btn-sec btn-sm" onclick="deleteTrustee('${esc(t._spId)}','${esc(t.name)}')">Remove</button></td>`
      :'';
    return`<tr>
      <td style="font-weight:600">${esc(t.name)}</td>
      <td style="font-size:11px">${esc(t.title)}</td>
      <td style="font-size:11px">${t.email?`<a href="mailto:${esc(t.email)}" style="color:var(--st)">${esc(t.email)}</a>`:'—'}</td>
      <td style="font-size:11px">${esc(t.invited)||'—'}</td>
      <td style="font-size:11px">${esc(t.confirmed)||'—'}</td>
      <td>${statusBadge(t.status)}</td>
      <td style="font-size:11px;color:var(--tx2)">${esc(t.notes)||'—'}</td>
      ${act}
    </tr>`;
  }).join('');
}

async function addTrustee(){
  const name=document.getElementById('tr-name').value.trim();
  if(!name){toast('Trustee name is required.','warning');return;}
  const t={
    name, email:document.getElementById('tr-email').value.trim(),
    title:document.getElementById('tr-title').value.trim(),
    invited:document.getElementById('tr-invited').value,
    confirmed:document.getElementById('tr-confirmed').value,
    status:document.getElementById('tr-status').value,
    notes:document.getElementById('tr-notes').value.trim()
  };
  try{
    const r=await spPost(LISTS.trustees,toSpTrustee(t));
    t._spId=r.id; trusteesData.push(t);
    // Clear form
    ['tr-name','tr-email','tr-title','tr-invited','tr-confirmed','tr-notes'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
    document.getElementById('tr-status').value='Pending Invitation';
    renderTrustees();
    toast(`${name} registered as a trustee.`,'success');
  }catch(e){toast('Failed to add trustee: '+e.message,'error');}
}

async function deleteTrustee(spId, name){
  if(!confirm(`Remove ${name} from the trustee register? This does not remove their Azure AD guest account.`)) return;
  try{
    const sid=await getSiteId();
    await gFetch(`/sites/${sid}/lists/${LISTS.trustees}/items/${spId}`,{method:'DELETE'});
    trusteesData=trusteesData.filter(t=>t._spId!==spId);
    renderTrustees();
    toast(`${name} removed from the register.`,'success');
  }catch(e){toast('Failed to remove trustee: '+e.message,'error');}
}

function exportTrusteesCsv(){
  const rows=[['Name','Designation','Email','Date Invited','Date Confirmed','Status','Notes'],
    ...trusteesData.map(t=>[t.name,t.title,t.email,t.invited,t.confirmed,t.status,t.notes].map(v=>'"'+(v||'').replace(/"/g,'""')+'"'))];
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([rows.map(r=>r.join(',')).join('\n')],{type:'text/csv'}));
  a.download='StudyTrust_Board_of_Trustees_'+td()+'.csv'; a.click();
}

// ════════════════════════════════════════════
// DIAGNOSTICS
// ════════════════════════════════════════════
function openDiag(){
  document.getElementById('diag-user').textContent = userName || '—';

  const roleEl = document.getElementById('diag-role');
  const roleLabels = {admin:'Admin ✅',trustee:'Trustee ✅',auditor:'Auditor ✅',ptt:'PTT ✅',personnel:'All Personnel (no matching group found)'};
  roleEl.textContent = roleLabels[currentRole] || currentRole;
  roleEl.style.color = currentRole==='personnel' ? '#856404' : '#1b5e20';

  document.getElementById('diag-grp-admin').textContent   = ADMIN_GROUP;
  document.getElementById('diag-grp-trustee').textContent = TRUSTEE_GROUP;
  document.getElementById('diag-grp-auditor').textContent = AUDITOR_GROUP;
  const pttEl = document.getElementById('diag-grp-ptt'); if(pttEl) pttEl.textContent = PTT_GROUP;

  const grpEl = document.getElementById('diag-groups');
  if(_diagGroups.length){
    grpEl.innerHTML = _diagGroups.map(g=>{
      const isMatch = [ADMIN_GROUP,TRUSTEE_GROUP,AUDITOR_GROUP,PTT_GROUP].some(cfg=>cfg.toLowerCase()===g.toLowerCase());
      return `<div style="color:${isMatch?'#1b5e20':'var(--tx)'}${isMatch?';font-weight:600':''}">${isMatch?'✅ ':''} ${g}</div>`;
    }).join('');
  } else {
    grpEl.textContent = _diagError ? 'Error — could not retrieve groups (see error below)' : 'No groups returned by Microsoft. Your account may not be in any Azure AD groups.';
  }

  const errBox = document.getElementById('diag-error-box');
  const errMsg = document.getElementById('diag-error-msg');
  if(_diagError){ errBox.style.display='block'; errMsg.textContent=_diagError; }
  else errBox.style.display='none';

  // Pre-fill editable fields
  document.getElementById('diag-admin-grp').value   = ADMIN_GROUP;
  document.getElementById('diag-trustee-grp').value = TRUSTEE_GROUP;
  document.getElementById('diag-auditor-grp').value = AUDITOR_GROUP;
  const pttIn = document.getElementById('diag-ptt-grp'); if(pttIn) pttIn.value = PTT_GROUP;

  openM('m-diag');
}

async function testList(listName, fields){
  const el=document.getElementById('diag-'+listName);
  if(el) el.innerHTML='<span style="color:#856404">Testing…</span>';
  try{
    const r=await spPost(listName,fields);
    if(r&&r.id){
      const sid=await getSiteId();
      await gFetch(`/sites/${sid}/lists/${listName}/items/${r.id}`,{method:'DELETE'}).catch(()=>{});
    }
    if(el) el.innerHTML='<span style="color:#1b5e20;font-weight:600">✅ Connected — all columns accepted</span>';
  }catch(e){
    if(el) el.innerHTML=`<span style="color:#7a2020;font-weight:600">✗ Error: ${e.message}</span> <button class="btn btn-sec btn-sm" style="margin-left:8px" onclick="isolateField('${listName}')">Find exact field →</button>`;
  }
}

// Tests each field of the failing list individually to identify the bad column
async function isolateField(listName){
  const el=document.getElementById('diag-'+listName);
  if(el) el.innerHTML='<span style="color:#856404">Isolating failing field — testing one by one…</span>';

  const allFields={
    GovPolicies:{
      Title:'__TEST__', PolicyID:'TEST-000', Category:'Test',
      Tier:'T1', PolicyVersion:'0.0', Status:'Draft',
      Risk:'Low', Approver:'Test', ReviewDate:null,
      Owner:'Test', Custodian:'Test',
      DraftUrl:'', PublishedUrl:'', ArchiveUrl:''
    }
  };

  const testFields=allFields[listName];
  if(!testFields){
    if(el) el.innerHTML='<span style="color:#7a2020">No field map for '+listName+'</span>';
    return;
  }

  const failing=[], passing=[];
  const sid=await getSiteId();

  for(const [fieldName, fieldValue] of Object.entries(testFields)){
    if(fieldValue===null||fieldValue==='') continue; // skip empty — not the cause
    try{
      const payload={Title:'__TEST__'};
      if(fieldName!=='Title') payload[fieldName]=fieldValue;
      const r=await gFetch(`/sites/${sid}/lists/${listName}/items`,{
        method:'POST',body:JSON.stringify({fields:payload})
      });
      passing.push(fieldName);
      // Delete test item
      if(r&&r.id) await gFetch(`/sites/${sid}/lists/${listName}/items/${r.id}`,{method:'DELETE'}).catch(()=>{});
    }catch(e){
      failing.push(fieldName);
    }
  }

  if(failing.length===0){
    if(el) el.innerHTML='<span style="color:#1b5e20;font-weight:600">✅ All fields passed individually — issue may be a combination. Try Seed again.</span>';
  } else {
    if(el) el.innerHTML=`<span style="color:#7a2020;font-weight:600">✗ Failing column(s): <strong>${failing.join(', ')}</strong></span><br><span style="font-size:11px;color:#5a1010">These column names do not exist in SharePoint or are named differently. Create them with exactly these names and try seeding again.</span>`;
  }
}

function saveDiagGroups(){
  const ag  = document.getElementById('diag-admin-grp').value.trim();
  const tg  = document.getElementById('diag-trustee-grp').value.trim();
  const aug = document.getElementById('diag-auditor-grp').value.trim();
  const pg  = (document.getElementById('diag-ptt-grp')||{}).value ? document.getElementById('diag-ptt-grp').value.trim() : PTT_GROUP;
  if(!ag||!tg||!aug||!pg){ toast('All four group names are required.','warning'); return; }
  // Always overwrite — this also clears any stale old-prefix config
  localStorage.setItem('st_gov_group_cfg', JSON.stringify({admin:ag,trustee:tg,auditor:aug,ptt:pg}));
  ADMIN_GROUP=ag; TRUSTEE_GROUP=tg; AUDITOR_GROUP=aug; PTT_GROUP=pg;
  // Also update the hub admin panel fields if visible
  const a=document.getElementById('cfg-admin-grp'), t=document.getElementById('cfg-trustee-grp'), au=document.getElementById('cfg-auditor-grp'), p=document.getElementById('cfg-ptt-grp');
  if(a) a.value=ag; if(t) t.value=tg; if(au) au.value=aug; if(p) p.value=pg;
  toast('Group names saved. Sign out and back in to re-detect your role.','success');
  closeM('m-diag');
}

// ════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════
async function loadAndRender(){
  setLoad('Detecting your access level…');
  try{
    loadGroupConfig();
    const role=await detectRole();
    const accs=msalApp.getAllAccounts();
    userName=accs[0]?.name||accs[0]?.username||'User';
    document.getElementById('user-name').textContent=userName;
    setRole(role);
    // Show toggle for admins so they can preview other roles
    if(role==='admin') document.getElementById('role-toggle').style.display='flex';
    setLoad('Loading governance register from SharePoint…');
    await loadData();
    renderAll();
    if(role==='admin') polDropdown();
    toast('Connected to SharePoint — '+userName,'success');
    // Initialise notification bell and webhook UI
    renderNotifBell();
    loadWebhookUI();
    // Check for overdue items (once per day)
    checkOverdueNotifs();
  }catch(e){
    hideLoad();
    toast('Load error: '+e.message,'error');
    // Fall back to personnel view with empty data
    setRole('personnel');
  }finally{hideLoad();}
}

async function init(){
  // If this page loaded inside a MSAL popup (Microsoft redirected here after login),
  // just process the auth response and close — never render the app in a popup window.
  if(window.opener && window.opener !== window){
    try{
      const _t = new msal.PublicClientApplication(MSAL_CFG);
      await _t.initialize();
      await _t.handleRedirectPromise();
    }catch(e){}
    try{ window.close(); }catch(_){}
    return;
  }
  try{
    const ok=await initMsal();
    if(ok){
      document.getElementById('signin-screen').classList.add('hidden');
      await loadAndRender();
    }else hideLoad();
  }catch(e){
    hideLoad();
    const err=document.getElementById('si-err');
    err.textContent=e.message; err.classList.add('show');
  }
}

init();


const SOB_GATE_STATUSES = ['Consult','PTT Review','Approved','Published'];
let _training = [], _policies = [], _sops = [];
let _tFilter = { docType:'All', status:'All', sobMissing:false, q:'' };
let _trainingLoaded = false;

function validateSOBBeforeTransition(sopItem, toStatus) {
  if (!SOB_GATE_STATUSES.includes(toStatus)) return true;
  const compliant = String(sopItem.sobCompliant || '').trim();
  if (compliant === 'Yes') return true;
  _showSOBBlockModal(sopItem, toStatus, compliant);
  return false;
}

function _showSOBBlockModal(sopItem, toStatus, currentValue) {
  const ref   = esc(sopItem.sopId || sopItem.title || 'this SOP');
  const label = currentValue === '' ? 'not yet assessed'
    : `set to <strong>${esc(currentValue)}</strong>`;
  _injectModal('sobBlockModal', `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:18px;">
      <span style="font-size:30px;">⚠️</span>
      <div>
        <h3 style="margin:0;color:#C0392B;font-size:16px;">SOB Validation — Status Transition Blocked</h3>
        <p style="margin:4px 0 0;font-size:13px;color:#555;">
          <strong>${ref}</strong> cannot advance to <strong>${esc(toStatus)}</strong>.
        </p>
      </div>
    </div>
    <p style="font-size:14px;">The <strong>Behavioural Standards</strong> section of this SOP has not been validated.
    The SOB Compliant field is ${label}.</p>
    <div style="background:#FEF9E7;border-left:4px solid #F39C12;padding:12px 16px;margin:16px 0;border-radius:4px;font-size:13px;">
      <strong>To unblock this SOP:</strong>
      <ol style="margin:8px 0 0;padding-left:18px;line-height:1.8;">
        <li>Open the SOP document and complete <strong>Section 6 — Behavioural Standards</strong>.</li>
        <li>Have the Manager: Operations sign the Behavioural Review row in <strong>Section 8</strong>.</li>
        <li>Edit this SOP record in the Hub and set <strong>SOB Compliant → Yes</strong>, plus review date and reviewer.</li>
      </ol>
    </div>
    <div style="display:flex;justify-content:flex-end;margin-top:20px;">
      <button style="padding:7px 18px;border:1px solid #ccc;border-radius:4px;cursor:pointer;"
        onclick="_closeModal('sobBlockModal')">Close</button>
    </div>
  `, '#C0392B');
}

function sobBadge(sopItem) {
  const v = String(sopItem.sobCompliant || '').trim();
  if (v === 'Yes')     return `<span class="bd bg" title="Behavioural Standards validated" style="cursor:default;">SOB ✓</span>`;
  if (v === 'No')      return `<span class="bd br" title="SOB NOT validated — blocks advancement" style="cursor:default;">SOB ✗</span>`;
  if (v === 'Pending') return `<span class="bd ba" title="SOB review pending" style="cursor:default;">SOB ⏳</span>`;
  return `<span class="bd bgy" title="SOB compliance not set" style="cursor:default;">SOB —</span>`;
}

async function loadTrainingTab() {
  const el = document.getElementById('trainingContent');
  if (!el) return;
  el.innerHTML = '<p style="padding:24px;color:#888;">Loading training materials…</p>';
  try {
    const [rawTrn, rawPol, rawSop] = await Promise.all([
      spGet('GovTraining'),
      spGet('GovPolicies'),
      spGet('GovSOPs'),
    ]);
    _training = rawTrn.map(i => ({...(i.fields||{}), _spId: i.id}));
    _policies = rawPol.map(i => ({...(i.fields||{}), _spId: i.id}));
    _sops     = rawSop.map(i => ({...(i.fields||{}), _spId: i.id}));
    _training = _training.map(_checkStaleness);
    _trainingLoaded = true;
    _renderTrainingTab();
  } catch (err) {
    _training = []; _trainingLoaded = false;
    _renderTrainingTab();
    const plog = document.getElementById('provision-log');
    if(plog) plog.innerHTML = `<span style="color:#C0392B;">⚠ ${esc(err.message)}</span>`;
  }
}

function _checkStaleness(item) {
  const linked = item.LinkedDocType === 'Policy'
    ? _policies.find(p => p.PolicyID === item.LinkedDocID)
    : _sops.find(s => s.SopID === item.LinkedDocID);
  if (!linked) return item;
  const currentVer = item.LinkedDocType === 'Policy'
    ? (linked.Version || '') : (linked.SopVersion || '');
  if (currentVer && item.LinkedDocVersion && currentVer !== item.LinkedDocVersion
      && item.TrainingStatus === 'Current') {
    item._autoStale = true;
    item._liveDocVersion = currentVer;
  }
  return item;
}

function _renderTrainingTab() {
  const el = document.getElementById('trainingContent');
  if (!el) return;
  const staleCount = _training.filter(t => t.TrainingStatus === 'Stale' || t._autoStale).length;
  const sobMissing = _training.filter(t => t.LinkedDocType === 'SOP' && t.SOBContentIncluded === 'No').length;
  const filtered = _training.filter(t => {
    if (_tFilter.docType !== 'All' && t.LinkedDocType !== _tFilter.docType) return false;
    if (_tFilter.status  !== 'All' && t.TrainingStatus !== _tFilter.status)  return false;
    if (_tFilter.sobMissing && !(t.LinkedDocType === 'SOP' && t.SOBContentIncluded === 'No')) return false;
    if (_tFilter.q) {
      const q = _tFilter.q.toLowerCase();
      if (!(t.Title||'').toLowerCase().includes(q) && !(t.LinkedDocID||'').toLowerCase().includes(q)) return false;
    }
    return true;
  });
  const provCard = (typeof currentRole !== 'undefined' && currentRole === 'admin') ? `
    <div style="background:#FFF8E1;border:1px solid #FFB300;border-left:4px solid #E67E22;border-radius:6px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
      <div style="font-size:12px;color:#7B5E00;"><strong>First-time setup:</strong> Provision the GovTraining SharePoint list and add SOB columns to GovSOPs. Run once only.</div>
      <button onclick="provisionGovTraining()" id="btn-provision-training"
        style="padding:6px 14px;background:#E67E22;color:#fff;border:none;border-radius:4px;font-size:12px;cursor:pointer;white-space:nowrap;">⚙ Provision GovTraining</button>
    </div>
    <div id="provision-log" style="font-size:11px;margin-bottom:10px;"></div>` : '';
  el.innerHTML = provCard + `
    <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:20px;">
      ${_kpiCard('Total Materials', _training.length, '#226397')}
      ${_kpiCard('Stale / Version Mismatch', staleCount, staleCount > 0 ? '#E67E22' : '#27AE60')}
      ${_kpiCard('SOP Training Missing SOB', sobMissing, sobMissing > 0 ? '#C0392B' : '#27AE60')}
    </div>
    ${staleCount > 0 ? `<div style="background:#FFF3CD;border:1px solid #FBBF24;border-left:4px solid #E67E22;padding:10px 16px;border-radius:4px;margin-bottom:12px;font-size:13px;"><strong>⚠ ${staleCount} training material${staleCount>1?'s are':' is'} stale or version-mismatched.</strong> Review and update, then mark as Current.</div>` : ''}
    ${sobMissing > 0 ? `<div style="background:#FDE8E8;border:1px solid #F87171;border-left:4px solid #C0392B;padding:10px 16px;border-radius:4px;margin-bottom:12px;font-size:13px;"><strong>⛔ ${sobMissing} SOP training material${sobMissing>1?'s do':' does'} not include SOB content.</strong> Update before distributing to personnel.</div>` : ''}
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;align-items:center;">
      <input type="text" placeholder="Search title or document ID…"
        style="flex:1;min-width:200px;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;"
        value="${esc(_tFilter.q)}" oninput="_tFilter.q=this.value;_renderTrainingTab()">
      <select style="padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;" onchange="_tFilter.docType=this.value;_renderTrainingTab()">
        <option value="All" ${_tFilter.docType==='All'?'selected':''}>All types</option>
        <option value="Policy" ${_tFilter.docType==='Policy'?'selected':''}>Policies</option>
        <option value="SOP" ${_tFilter.docType==='SOP'?'selected':''}>SOPs</option>
      </select>
      <select style="padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;" onchange="_tFilter.status=this.value;_renderTrainingTab()">
        <option value="All" ${_tFilter.status==='All'?'selected':''}>All statuses</option>
        <option value="Current" ${_tFilter.status==='Current'?'selected':''}>Current</option>
        <option value="Stale" ${_tFilter.status==='Stale'?'selected':''}>Stale</option>
        <option value="Under Review" ${_tFilter.status==='Under Review'?'selected':''}>Under Review</option>
        <option value="Archived" ${_tFilter.status==='Archived'?'selected':''}>Archived</option>
      </select>
      <label style="display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;white-space:nowrap;">
        <input type="checkbox" ${_tFilter.sobMissing?'checked':''} onchange="_tFilter.sobMissing=this.checked;_renderTrainingTab()"> Missing SOB only
      </label>
      <button onclick="openAddTrainingModal()"
        style="margin-left:auto;padding:7px 16px;background:#226397;color:#fff;border:none;border-radius:4px;font-size:13px;cursor:pointer;">+ Add Material</button>
    </div>
    <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead><tr style="background:#226397;color:#fff;text-align:left;">
          <th style="padding:9px 10px;">Training Title</th>
          <th style="padding:9px 10px;">Linked Document</th>
          <th style="padding:9px 10px;">Format</th>
          <th style="padding:9px 10px;text-align:center;">Trained Ver.</th>
          <th style="padding:9px 10px;text-align:center;">Current Ver.</th>
          <th style="padding:9px 10px;text-align:center;">Status</th>
          <th style="padding:9px 10px;text-align:center;">SOB in Training</th>
          <th style="padding:9px 10px;text-align:center;">Actions</th>
        </tr></thead>
        <tbody>${filtered.length === 0
          ? '<tr><td colspan="8" style="text-align:center;color:#888;padding:28px;">No training materials match the current filters.</td></tr>'
          : filtered.map(_trainingRow).join('')
        }</tbody>
      </table>
    </div>`;
}

function _trainingRow(t) {
  const isStale = t.TrainingStatus === 'Stale' || t._autoStale;
  const liveVer = t._liveDocVersion || '—';
  const statusBadge = ({
    'Current':      '<span style="background:#D1FAE5;color:#065F46;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">Current</span>',
    'Stale':        '<span style="background:#FEE2E2;color:#991B1B;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">Stale</span>',
    'Under Review': '<span style="background:#FEF3C7;color:#92400E;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">Under Review</span>',
    'Archived':     '<span style="background:#F3F4F6;color:#6B7280;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">Archived</span>',
  }[t.TrainingStatus]) || `<span style="background:#F3F4F6;color:#6B7280;padding:2px 8px;border-radius:10px;font-size:11px;">${esc(t.TrainingStatus||'—')}</span>`;
  const autoStaleChip = t._autoStale
    ? `<br><span style="background:#FED7AA;color:#92400E;padding:1px 6px;border-radius:8px;font-size:10px;">⚠ ver. mismatch — doc now ${esc(t._liveDocVersion)}</span>` : '';
  const sobChip = t.LinkedDocType === 'SOP'
    ? ({'Yes':'<span style="background:#D1FAE5;color:#065F46;padding:2px 8px;border-radius:10px;font-size:11px;">SOB ✓</span>',
        'No': '<span style="background:#FEE2E2;color:#991B1B;padding:2px 8px;border-radius:10px;font-size:11px;">SOB ✗</span>'}[t.SOBContentIncluded]
      || '<span style="background:#F3F4F6;color:#6B7280;padding:2px 8px;border-radius:10px;font-size:11px;">N/A</span>')
    : '<span style="background:#F3F4F6;color:#9CA3AF;padding:2px 8px;border-radius:10px;font-size:11px;">N/A</span>';
  const titleCell = t.MaterialURL
    ? `<a href="${esc(t.MaterialURL)}" target="_blank" rel="noopener" style="color:#226397;text-decoration:none;font-weight:500;">${esc(t.Title||'')}</a>${autoStaleChip}`
    : `${esc(t.Title||'')}${autoStaleChip}`;
  const docChip = t.LinkedDocType === 'Policy'
    ? '<span style="background:#DBEAFE;color:#1E40AF;padding:1px 6px;border-radius:8px;font-size:11px;margin-right:4px;">POL</span>'
    : '<span style="background:#EDE9FE;color:#5B21B6;padding:1px 6px;border-radius:8px;font-size:11px;margin-right:4px;">SOP</span>';
  const markStaleBtn = t._autoStale
    ? `<button onclick="markTrainingStale('${esc(t._spId)}')" style="background:#FEF3C7;border:1px solid #F59E0B;color:#92400E;padding:3px 8px;border-radius:4px;font-size:11px;cursor:pointer;margin-left:4px;">Mark Stale</button>` : '';
  return `<tr style="background:${isStale?'#FFF8F0':'#fff'};border-bottom:1px solid #F0F0F0;">
    <td style="padding:9px 10px;">${titleCell}</td>
    <td style="padding:9px 10px;">${docChip}${esc(t.LinkedDocID||'')}</td>
    <td style="padding:9px 10px;color:#555;">${esc(t.MaterialType||'—')}</td>
    <td style="padding:9px 10px;text-align:center;">${esc(t.LinkedDocVersion||'—')}</td>
    <td style="padding:9px 10px;text-align:center;${t._autoStale?'color:#E67E22;font-weight:700;':''}">${esc(liveVer)}</td>
    <td style="padding:9px 10px;text-align:center;">${statusBadge}</td>
    <td style="padding:9px 10px;text-align:center;">${sobChip}</td>
    <td style="padding:9px 10px;text-align:center;">
      <button onclick="openEditTrainingModal('${esc(t._spId)}')" style="background:none;border:none;cursor:pointer;font-size:15px;" title="Edit">✏️</button>
      ${markStaleBtn}
    </td>
  </tr>`;
}

function _kpiCard(label, value, colour) {
  return `<div style="background:#fff;border:1px solid #E5E7EB;border-top:3px solid ${colour};border-radius:6px;padding:14px 20px;min-width:160px;flex:1;">
    <div style="font-size:26px;font-weight:700;color:${colour};">${value}</div>
    <div style="font-size:12px;color:#6B7280;margin-top:2px;">${label}</div>
  </div>`;
}

function openAddTrainingModal() {
  _injectModal('addTrainingModal', `
    <h3 style="margin:0 0 20px;color:#226397;">Add Training Material</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Training Title *</label>
        <input id="trnTitle" type="text" placeholder="e.g. Leave Management SOP — Induction Presentation"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Linked Document ID *</label>
        <input id="trnLinkedID" type="text" placeholder="e.g. SOP-003"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Document Type *</label>
        <select id="trnDocType" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          <option value="SOP">SOP</option><option value="Policy">Policy</option></select></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Document Version *</label>
        <input id="trnVersion" type="text" placeholder="e.g. v1.0"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Material Format *</label>
        <select id="trnType" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          <option>Presentation</option><option>Video</option><option>Quick Reference</option>
          <option>Assessment</option><option>Induction</option></select></div>
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Material URL</label>
        <input id="trnURL" type="url" placeholder="https://…"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">SOB Content Included?</label>
        <select id="trnSOB" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          <option value="Yes">Yes — behavioural expectations embedded</option>
          <option value="No">No — needs updating</option>
          <option value="N/A">N/A (Policy training)</option></select></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Acknowledgement Required?</label>
        <select id="trnAck" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          <option value="Yes">Yes</option><option value="No">No</option></select></div>
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Notes</label>
        <textarea id="trnNotes" rows="2"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;resize:vertical;"></textarea></div>
    </div>
    <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:20px;">
      <button onclick="_closeModal('addTrainingModal')"
        style="padding:7px 18px;border:1px solid #ccc;background:#fff;border-radius:4px;cursor:pointer;font-size:13px;">Cancel</button>
      <button onclick="_saveNewTraining()"
        style="padding:7px 18px;background:#226397;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;">Save</button>
    </div>`);
}

async function _saveNewTraining() {
  const title    = document.getElementById('trnTitle').value.trim();
  const linkedID = document.getElementById('trnLinkedID').value.trim();
  const version  = document.getElementById('trnVersion').value.trim();
  if (!title || !linkedID || !version) {
    alert('Title, Linked Document ID, and Document Version are required.');
    return;
  }
  const nextID = 'TRN-' + String(_training.length + 1).padStart(3, '0');
  try {
    await spPost('GovTraining', {
      Title: title, TrainingID: nextID, LinkedDocID: linkedID,
      LinkedDocType: document.getElementById('trnDocType').value,
      LinkedDocVersion: version,
      MaterialType: document.getElementById('trnType').value,
      MaterialURL:  document.getElementById('trnURL').value.trim(),
      TrainingStatus: 'Current',
      SOBContentIncluded: document.getElementById('trnSOB').value,
      AcknowledgementRequired: document.getElementById('trnAck').value,
      Notes: document.getElementById('trnNotes').value.trim(),
    });
    _closeModal('addTrainingModal');
    await loadTrainingTab();
  } catch (err) { alert('Save failed: ' + err.message); }
}

function openEditTrainingModal(id) {
  const item = _training.find(t => t._spId === id);
  if (!item) return;
  _injectModal('editTrainingModal', `
    <h3 style="margin:0 0 6px;color:#226397;">Edit Training Material</h3>
    <p style="margin:0 0 18px;font-size:12px;color:#888;">${esc(item.TrainingID||'')} — ${esc(item.Title||'')}</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Title *</label>
        <input id="eTrnTitle" type="text" value="${esc(item.Title||'')}"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Document Version (trained on)</label>
        <input id="eTrnVersion" type="text" value="${esc(item.LinkedDocVersion||'')}"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Training Status</label>
        <select id="eTrnStatus" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          ${['Current','Stale','Under Review','Archived'].map(s =>
            `<option ${item.TrainingStatus===s?'selected':''}>${s}</option>`).join('')}
        </select></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">SOB Content Included?</label>
        <select id="eTrnSOB" style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
          ${['Yes','No','N/A'].map(s =>
            `<option ${item.SOBContentIncluded===s?'selected':''}>${s}</option>`).join('')}
        </select></div>
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Material URL</label>
        <input id="eTrnURL" type="url" value="${esc(item.MaterialURL||'')}"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Reviewed By</label>
        <input id="eTrnReviewer" type="text" value="${esc(item.ReviewedBy||'')}"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Last Reviewed Date</label>
        <input id="eTrnReviewDate" type="date"
          value="${item.LastReviewedDate ? item.LastReviewedDate.substring(0,10) : ''}"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;"></div>
      <div style="grid-column:1/-1;"><label style="font-size:12px;font-weight:600;display:block;margin-bottom:4px;">Notes</label>
        <textarea id="eTrnNotes" rows="2"
          style="width:100%;padding:7px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box;resize:vertical;">${esc(item.Notes||'')}</textarea></div>
    </div>
    <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:20px;">
      <button onclick="_closeModal('editTrainingModal')"
        style="padding:7px 18px;border:1px solid #ccc;background:#fff;border-radius:4px;cursor:pointer;font-size:13px;">Cancel</button>
      <button onclick="_updateTraining('${id}')"
        style="padding:7px 18px;background:#226397;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;">Save Changes</button>
    </div>`);
}

async function _updateTraining(id) {
  try {
    await spPatch('GovTraining', id, {
      Title:            document.getElementById('eTrnTitle').value.trim(),
      LinkedDocVersion: document.getElementById('eTrnVersion').value.trim(),
      TrainingStatus:   document.getElementById('eTrnStatus').value,
      SOBContentIncluded: document.getElementById('eTrnSOB').value,
      MaterialURL:      document.getElementById('eTrnURL').value.trim(),
      ReviewedBy:       document.getElementById('eTrnReviewer').value.trim(),
      LastReviewedDate: document.getElementById('eTrnReviewDate').value || null,
      Notes:            document.getElementById('eTrnNotes').value.trim(),
    });
    _closeModal('editTrainingModal');
    await loadTrainingTab();
  } catch (err) { alert('Update failed: ' + err.message); }
}

async function markTrainingStale(id) {
  if (!confirm('Mark this training material as Stale in SharePoint?')) return;
  try { await spPatch('GovTraining', id, {TrainingStatus:'Stale'}); await loadTrainingTab(); }
  catch (err) { alert('Update failed: ' + err.message); }
}

function getTrainingIndicator(docID, docType) {
  if (!_trainingLoaded) return '';
  const related = _training.filter(t => t.LinkedDocID === docID && t.LinkedDocType === docType);
  if (related.length === 0)
    return `<span style="background:#F3F4F6;color:#9CA3AF;padding:1px 7px;border-radius:8px;font-size:11px;"
      title="No training materials linked">🎓 None</span>`;
  const hasStale   = related.some(t => t.TrainingStatus === 'Stale' || t._autoStale);
  const hasMissSOB = related.some(t => t.LinkedDocType === 'SOP' && t.SOBContentIncluded === 'No');
  if (hasStale || hasMissSOB) {
    const tip = hasStale ? 'One or more training materials are stale' : 'SOB content missing from training';
    return `<span style="background:#FEF3C7;color:#92400E;padding:1px 7px;border-radius:8px;font-size:11px;cursor:pointer;"
      title="${tip}" onclick="showTab('sec-training');_tFilter.q='${docID}';_renderTrainingTab();">
      🎓 ⚠ ${related.length}</span>`;
  }
  return `<span style="background:#D1FAE5;color:#065F46;padding:1px 7px;border-radius:8px;font-size:11px;cursor:pointer;"
    title="${related.length} training material(s) — all current"
    onclick="showTab('sec-training');_tFilter.q='${docID}';_renderTrainingTab();">
    🎓 ${related.length}</span>`;
}

async function provisionGovTraining() {
  const logEl = document.getElementById('provision-log');
  function addLog(msg, ok) {
    if (!logEl) return;
    const line = document.createElement('div');
    line.style.marginBottom = '3px';
    line.textContent = (ok === undefined ? '⏳ ' : ok ? '✅ ' : '❌ ') + msg;
    logEl.appendChild(line);
  }
  const btn = document.getElementById('btn-provision-training');
  if (btn) { btn.disabled = true; btn.textContent = 'Provisioning…'; }
  try {
    addLog('Getting SharePoint site ID…');
    const siteId = await getSiteId();
    addLog('Site ID obtained.', true);
    addLog('Creating GovTraining list…');
    const listResp = await gFetch(`/sites/${siteId}/lists`, {method:'POST',body:JSON.stringify({displayName:'GovTraining',columns:[],list:{template:'genericList'}})});
    const listId = listResp.id;
    addLog('GovTraining list created.', true);
    await new Promise(r => setTimeout(r, 2000));
    const cols = [
      {name:'TrainingID',              type:'text'},
      {name:'LinkedDocID',             type:'text'},
      {name:'LinkedDocType',           type:'choice', choices:['Policy','SOP']},
      {name:'LinkedDocVersion',        type:'text'},
      {name:'MaterialType',            type:'choice', choices:['Presentation','Video','Quick Reference','Assessment','Induction']},
      {name:'MaterialURL',             type:'text'},
      {name:'TrainingStatus',          type:'choice', choices:['Current','Stale','Under Review','Archived']},
      {name:'LastReviewedDate',        type:'dateTime'},
      {name:'ReviewedBy',              type:'text'},
      {name:'AcknowledgementRequired', type:'choice', choices:['Yes','No']},
      {name:'SOBContentIncluded',      type:'choice', choices:['Yes','No','N/A']},
      {name:'Notes',                   type:'text'},
    ];
    for (const col of cols) {
      const body = col.type === 'choice'
        ? {name: col.name, choice: {displayAs:'dropDownMenu', choices: col.choices}}
        : col.type === 'dateTime'
        ? {name: col.name, dateTime: {displayAs:'dateOnly'}}
        : {name: col.name, text: {}};
      await gFetch(`/sites/${siteId}/lists/${listId}/columns`, {method:'POST',body:JSON.stringify(body)});
      addLog('Column ' + col.name + ' added.', true);
    }
    addLog('Adding SOB columns to GovSOPs…');
    const sopResp = await gFetch(`/sites/${siteId}/lists?$filter=displayName eq 'GovSOPs'`);
    const sopListId = sopResp.value[0].id;
    await gFetch(`/sites/${siteId}/lists/${sopListId}/columns`, {method:'POST',body:JSON.stringify({name:'SOBCompliant',choice:{displayAs:'dropDownMenu',choices:['Yes','No','Pending']}})});
    addLog('GovTraining provisioning complete.', true);
    if (btn) { btn.disabled = false; btn.textContent = 'Provision Training List'; }
  } catch(e) {
    addLog('Provisioning failed: ' + (e.message || e), false);
    if (btn) { btn.disabled = false; btn.textContent = 'Provision Training List'; }
  }
}
