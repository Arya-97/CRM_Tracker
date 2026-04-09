// Theme
(function(){const t=localStorage.getItem('crm_theme')||'dark';if(t==='light')document.body.classList.add('light');})();
function toggleTheme(){const l=document.body.classList.toggle('light');localStorage.setItem('crm_theme',l?'light':'dark');document.getElementById('theme-btn').textContent=l?'☀️ Light':'🌙 Dark';}
document.addEventListener('DOMContentLoaded',()=>{
  const t=localStorage.getItem('crm_theme')||'dark';
  document.getElementById('theme-btn').textContent=t==='light'?'☀️ Light':'🌙 Dark';
  const sc=localStorage.getItem('crm_sidebar_collapsed')==='true';
  if(sc)document.querySelector('.sidebar').classList.add('collapsed');
  initParticles();
});

const CFG={
  get clientId(){return HARDCODED_CLIENT_ID;},
  get sheetIdMain(){return HARDCODED_SHEET_ID;},          // Main Tracker tab
  get sheetIdMeta(){return HARDCODED_SHEET_ID_META;},     // Audit + Users + Requests tabs
};

const SHEETS={main:'Main Tracker',audit:'Audit Log',users:'Users',requests:'Access Requests'};

const COLUMNS={
  regno:0,student_name:1,center:2,region:3,newpayment_checks:4,
  fees_amt:5,fees_paid:6,form_status:7,emi1_paid_date:8,emi2_paid_date:9,
  emi3_paid_date:10,emi4_paid_date:11,total_paid_date:12,attendance_15days:13,
  last_punch_date:14,ptp:15,connected_status:16,dialled_status:17,
  yesterday_disposition:18,todays_disposition:19,other_remarks:20,due_date:21
};
const NUM_COLS=22;
const MAIN_HEADERS=['Reg No','Student Name','Center','Region','Payment Check','Fees Amt','Fees Paid','Form Status','1st EMI Paid Date','2nd EMI Paid Date','3rd EMI Paid Date','4th EMI Paid Date','Total Paid Date','% 15 Days Attendance','Last Punch Date','PTP','Connected Status','Dialled Status','Yesterday Disposition',"Today's Disposition",'Other Remarks','Due Date'];

// State
let accessToken=null,currentUser=null;
let allData=[],filteredData=[],auditLog=[],users=[],accessRequests=[];
let selectedIdx=null,allRegions=[],allCenters=[],regionCenterMap={};
let selectedPayments=new Set();

// Cache
const CACHE_DURATION=5*60*1000;
const cache={
  data:null,timestamp:null,
  isValid(){return this.data&&this.timestamp&&(Date.now()-this.timestamp<CACHE_DURATION);},
  set(d){this.data=d;this.timestamp=Date.now();try{sessionStorage.setItem('tracker_cache',JSON.stringify({data:d,timestamp:this.timestamp}));}catch(e){}},
  get(){if(this.isValid())return this.data;try{const s=sessionStorage.getItem('tracker_cache');if(s){const p=JSON.parse(s);if(Date.now()-p.timestamp<CACHE_DURATION){this.data=p.data;this.timestamp=p.timestamp;return this.data;}}}catch(e){}return null;},
  clear(){this.data=null;this.timestamp=null;try{sessionStorage.removeItem('tracker_cache');}catch(e){}}
};

// ════════════════════════════════════════════════════════
// PARTICLES BACKGROUND
// ════════════════════════════════════════════════════════
function initParticles(){
  const canvas=document.getElementById('particle-canvas');
  if(!canvas)return;
  const ctx=canvas.getContext('2d');
  let W=canvas.width=window.innerWidth,H=canvas.height=window.innerHeight;
  const particles=Array.from({length:60},()=>({
    x:Math.random()*W,y:Math.random()*H,
    vx:(Math.random()-.5)*.3,vy:(Math.random()-.5)*.3,
    r:Math.random()*1.5+.5,o:Math.random()*.4+.1
  }));
  window.addEventListener('resize',()=>{W=canvas.width=window.innerWidth;H=canvas.height=window.innerHeight;});
  function draw(){
    ctx.clearRect(0,0,W,H);
    const isDark=!document.body.classList.contains('light');
    particles.forEach(p=>{
      p.x+=p.vx;p.y+=p.vy;
      if(p.x<0)p.x=W;if(p.x>W)p.x=0;
      if(p.y<0)p.y=H;if(p.y>H)p.y=0;
      ctx.beginPath();ctx.arc(p.x,p.y,p.r,0,Math.PI*2);
      ctx.fillStyle=isDark?`rgba(74,158,255,${p.o})`:`rgba(37,99,235,${p.o*0.5})`;
      ctx.fill();
    });
    // Draw connections
    for(let i=0;i<particles.length;i++){
      for(let j=i+1;j<particles.length;j++){
        const dx=particles[i].x-particles[j].x,dy=particles[i].y-particles[j].y;
        const dist=Math.sqrt(dx*dx+dy*dy);
        if(dist<120){
          ctx.beginPath();
          ctx.moveTo(particles[i].x,particles[i].y);
          ctx.lineTo(particles[j].x,particles[j].y);
          const alpha=(1-dist/120)*0.08*(isDark?1:0.4);
          ctx.strokeStyle=isDark?`rgba(74,158,255,${alpha})`:`rgba(37,99,235,${alpha})`;
          ctx.lineWidth=.5;ctx.stroke();
        }
      }
    }
    requestAnimationFrame(draw);
  }
  draw();
}

// ════════════════════════════════════════════════════════
// AUTH SCREENS
// ════════════════════════════════════════════════════════
function showLoginForm(){document.getElementById('auth-screen').style.display='flex';document.getElementById('signup-screen').style.display='none';}
function showSignupForm(){document.getElementById('auth-screen').style.display='none';document.getElementById('signup-screen').style.display='flex';}

// ════════════════════════════════════════════════════════
// SIGNUP
// ════════════════════════════════════════════════════════
async function submitSignupRequest(){
  const email=document.getElementById('signup-email').value.trim().toLowerCase();
  const name=document.getElementById('signup-name').value.trim();
  const manager=document.getElementById('signup-manager').value.trim().toLowerCase();
  if(!email||!email.includes('@')){showSignupErr('Please enter a valid email');return;}
  if(!name){showSignupErr('Please enter your name');return;}
  if(!manager||!manager.includes('@')){showSignupErr('Please enter manager email');return;}
  showLoader('Submitting request…');
  try{
    if(!accessToken){
      await new Promise((resolve,reject)=>{
        google.accounts.oauth2.initTokenClient({
          client_id:HARDCODED_CLIENT_ID,
          scope:'https://www.googleapis.com/auth/spreadsheets',
          callback:(r)=>{if(r.error)reject(r);else{accessToken=r.access_token;resolve();}},
          error_callback:reject
        }).requestAccessToken();
      });
    }
    await ensureMetaSheets();
    await sheetsAppend(CFG.sheetIdMeta,SHEETS.requests,[[email,name,manager,'pending',nowStr()]]);
    document.getElementById('signup-form').style.display='none';
    document.getElementById('signup-success').style.display='block';
    hideLoader();
  }catch(err){hideLoader();showSignupErr('Error: '+err.message);}
}

// ════════════════════════════════════════════════════════
// AUTH
// ════════════════════════════════════════════════════════
function startGoogleAuth(){
  if(!HARDCODED_CLIENT_ID||HARDCODED_CLIENT_ID==='YOUR_CLIENT_ID_HERE'){showAuthErr('⚙️ Set HARDCODED_CLIENT_ID');return;}
  if(!HARDCODED_SHEET_ID||HARDCODED_SHEET_ID==='YOUR_SHEET_ID_HERE'){showAuthErr('⚙️ Set HARDCODED_SHEET_ID');return;}
  if(!window.google){showAuthErr('Google auth failed to load');return;}
  google.accounts.oauth2.initTokenClient({
    client_id:HARDCODED_CLIENT_ID,
    scope:'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email',
    callback:onTokenReceived,
    error_callback:e=>showAuthErr('Auth error: '+(e.message||JSON.stringify(e)))
  }).requestAccessToken();
}

async function onTokenReceived(r){
  if(r.error){showAuthErr(r.error);return;}
  accessToken=r.access_token;showLoader('Verifying…');
  try{
    const uInfo=await gFetch('https://www.googleapis.com/oauth2/v2/userinfo').then(r=>r.json());
    await ensureMetaSheets();
    await loadUsers();
    const found=users.find(u=>u.email.toLowerCase()===uInfo.email.toLowerCase());
    if(!found){
      if(users.length===0){
        await sheetsAppend(CFG.sheetIdMeta,SHEETS.users,[[uInfo.email,'admin','system',nowStr(),'']]);
        await loadUsers();
        currentUser={...uInfo,role:'admin',manager:''};
        bootApp();
      }else{
        hideLoader();
        showAuthErr('❌ Access Denied: '+uInfo.email+' not registered.');
        accessToken=null;
      }
    }else{
      currentUser={...uInfo,role:found.role,manager:found.manager||''};
      bootApp();
    }
  }catch(e){hideLoader();showAuthErr('Error: '+e.message);}
}

function bootApp(){
  hideLoader();
  document.getElementById('auth-screen').style.display='none';
  document.getElementById('app').style.display='flex';
  document.getElementById('user-name').textContent=currentUser.name||currentUser.email;
  document.getElementById('user-role').textContent=currentUser.role.toUpperCase();
  const av=document.getElementById('user-avatar');
  if(currentUser.picture){av.innerHTML=`<img src="${currentUser.picture}" style="width:32px;height:32px;border-radius:50%;object-fit:cover"/>`;}
  else{av.textContent=(currentUser.name||currentUser.email)[0].toUpperCase();}

  if(currentUser.role==='admin'){
    document.querySelectorAll('.admin-nav').forEach(el=>el.style.display='');
    document.querySelectorAll('.manager-nav').forEach(el=>el.style.display='');
    document.querySelectorAll('.agent-hide').forEach(el=>el.style.display='');
    const ci=document.getElementById('info-client-id'),si=document.getElementById('info-sheet-id');
    if(ci)ci.textContent=HARDCODED_CLIENT_ID.substring(0,30)+'…';
    if(si)si.textContent=HARDCODED_SHEET_ID.substring(0,30)+'…';
  }else if(currentUser.role==='manager'){
    document.querySelectorAll('.admin-nav').forEach(el=>el.style.display='none');
    document.querySelectorAll('.manager-nav').forEach(el=>el.style.display='');
    document.querySelectorAll('.agent-hide').forEach(el=>el.style.display='none');
  }else{
    document.querySelectorAll('.admin-nav').forEach(el=>el.style.display='none');
    document.querySelectorAll('.manager-nav').forEach(el=>el.style.display='none');
    document.querySelectorAll('.agent-hide').forEach(el=>el.style.display='none');
  }

  // Animate stats on load
  setTimeout(()=>loadAll(),100);
}

function doLogout(){
  cache.clear();accessToken=null;currentUser=null;
  document.getElementById('app').style.display='none';
  document.getElementById('auth-screen').style.display='flex';
  document.getElementById('auth-err').classList.remove('show');
}

// ════════════════════════════════════════════════════════
// SHEETS API — DUAL SHEET SUPPORT
// ════════════════════════════════════════════════════════
const API_BASE='https://sheets.googleapis.com/v4/spreadsheets';
function gFetch(url,opts={}){return fetch(url,{...opts,headers:{'Authorization':'Bearer '+accessToken,'Content-Type':'application/json',...(opts.headers||{})}});}

async function sheetsGet(sheetId,range){
  const r=await gFetch(`${API_BASE}/${sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`);
  if(!r.ok){const e=await r.json();throw new Error(e.error?.message||'Read error');}
  return r.json();
}
async function sheetsUpdate(sheetId,range,values){
  const r=await gFetch(`${API_BASE}/${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{method:'PUT',body:JSON.stringify({range,majorDimension:'ROWS',values})});
  if(!r.ok){const e=await r.json();throw new Error(e.error?.message||'Write error');}
  return r.json();
}
async function sheetsAppend(sheetId,sheet,values){
  const r=await gFetch(`${API_BASE}/${sheetId}/values/${encodeURIComponent(sheet+'!A1')}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,{method:'POST',body:JSON.stringify({majorDimension:'ROWS',values})});
  if(!r.ok){const e=await r.json();throw new Error(e.error?.message||'Append error');}
  return r.json();
}
async function sheetsBatch(sheetId,reqs){
  const r=await gFetch(`${API_BASE}/${sheetId}:batchUpdate`,{method:'POST',body:JSON.stringify({requests:reqs})});
  if(!r.ok){const e=await r.json();throw new Error(e.error?.message||'Batch error');}
  return r.json();
}
async function getSheetInfo(sheetId){
  const r=await gFetch(`${API_BASE}/${sheetId}`);
  if(!r.ok)throw new Error('Cannot read sheet');
  return r.json();
}

async function ensureMetaSheets(){
  const info=await getSheetInfo(CFG.sheetIdMeta);
  const existing=info.sheets.map(s=>s.properties.title);
  const toCreate=[SHEETS.audit,SHEETS.users,SHEETS.requests].filter(n=>!existing.includes(n));
  if(toCreate.length){
    await sheetsBatch(CFG.sheetIdMeta,toCreate.map(t=>({addSheet:{properties:{title:t}}})));
    if(toCreate.includes(SHEETS.audit))await sheetsUpdate(CFG.sheetIdMeta,`${SHEETS.audit}!A1`,[['timestamp','agent_email','agent_name','regno','student_name','field','old_value','new_value']]);
    if(toCreate.includes(SHEETS.users))await sheetsUpdate(CFG.sheetIdMeta,`${SHEETS.users}!A1`,[['email','role','added_by','added_on','manager']]);
    if(toCreate.includes(SHEETS.requests))await sheetsUpdate(CFG.sheetIdMeta,`${SHEETS.requests}!A1`,[['email','name','manager_email','status','submitted_on']]);
  }
}

async function ensureMainSheet(){
  const info=await getSheetInfo(CFG.sheetIdMain);
  const existing=info.sheets.map(s=>s.properties.title);
  if(!existing.includes(SHEETS.main)){
    await sheetsBatch(CFG.sheetIdMain,[{addSheet:{properties:{title:SHEETS.main}}}]);
    await sheetsUpdate(CFG.sheetIdMain,`${SHEETS.main}!A1`,[MAIN_HEADERS]);
  }
}

// ════════════════════════════════════════════════════════
// MULTI-SELECT PAYMENT
// ════════════════════════════════════════════════════════
function buildPaymentDropdown(paymentOptions){
  const wrapper=document.getElementById('payment-multiselect-wrapper');if(!wrapper)return;
  const allPay=paymentOptions.filter(Boolean);
  wrapper.innerHTML=`<div class="multiselect-container"><div class="multiselect-trigger" onclick="togglePaymentDropdown(event)"><span id="payment-ms-label">All Payments</span><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg></div><div class="multiselect-dropdown" id="payment-ms-dropdown"><div class="ms-option ms-all" onclick="toggleAllPayments(event)"><span class="ms-check" id="ms-check-all">✓</span><span>All Payments</span></div><div class="ms-divider"></div>${allPay.map(p=>`<div class="ms-option" onclick="togglePayment(event,'${p}')"><span class="ms-check ms-pay-check" data-pay="${p}"></span><span>${p}</span></div>`).join('')}</div></div>`;
  selectedPayments.clear();updatePaymentLabel();
}
function togglePaymentDropdown(e){e.stopPropagation();const dd=document.getElementById('payment-ms-dropdown');const isOpen=dd.classList.toggle('open');if(isOpen)document.addEventListener('click',closePaymentDropdown,{once:true});}
function closePaymentDropdown(){const dd=document.getElementById('payment-ms-dropdown');if(dd)dd.classList.remove('open');}
function toggleAllPayments(e){e.stopPropagation();selectedPayments.clear();document.querySelectorAll('.ms-pay-check').forEach(el=>el.textContent='');const ac=document.getElementById('ms-check-all');if(ac)ac.textContent='✓';updatePaymentLabel();applyFilters();}
function togglePayment(e,val){e.stopPropagation();if(selectedPayments.has(val))selectedPayments.delete(val);else selectedPayments.add(val);document.querySelectorAll('.ms-pay-check').forEach(el=>{el.textContent=selectedPayments.has(el.dataset.pay)?'✓':'';});const ac=document.getElementById('ms-check-all');if(ac)ac.textContent=selectedPayments.size===0?'✓':'';updatePaymentLabel();applyFilters();}
function updatePaymentLabel(){const lbl=document.getElementById('payment-ms-label');if(!lbl)return;if(selectedPayments.size===0)lbl.textContent='All Payments';else if(selectedPayments.size===1)lbl.textContent=[...selectedPayments][0];else lbl.textContent=selectedPayments.size+' selected';}

// ════════════════════════════════════════════════════════
// LOAD DATA
// ════════════════════════════════════════════════════════
async function loadAll(){
  const btn=document.getElementById('sync-btn');btn.classList.add('spinning');showLoader('Syncing…');
  try{
    const cached=cache.get();
    if(cached&&!btn.classList.contains('force-refresh')){
      allData=cached.main||[];auditLog=cached.audit||[];users=cached.users||[];accessRequests=cached.requests||[];
      buildFilters();applyFilters();renderAudit();renderUsersList();renderTeamMembers();renderAccessRequests();
      hideLoader();showToast('✅ Loaded from cache','info');btn.classList.remove('spinning');return;
    }
    await Promise.all([loadMainData(),loadAuditData(),loadUsers(),loadAccessRequests()]);
    cache.set({main:allData,audit:auditLog,users:users,requests:accessRequests});
    buildFilters();applyFilters();renderAudit();renderUsersList();renderTeamMembers();renderAccessRequests();
    hideLoader();showToast('✅ Synced successfully','ok');
  }catch(e){hideLoader();showToast('❌ '+e.message,'err');}
  btn.classList.remove('spinning');btn.classList.remove('force-refresh');
}

async function loadMainData(){
  const res=await sheetsGet(CFG.sheetIdMain,`${SHEETS.main}!A1:V2000`);
  const rows=res.values||[];
  if(rows.length<=1){allData=[];return;}
  allData=rows.slice(1).map((row,i)=>{
    const r={};
    Object.entries(COLUMNS).forEach(([k,ci])=>r[k]=(row[ci]||'').toString().trim());
    r._rowIndex=i+2;return r;
  }).filter(r=>r.regno);
}

async function loadAuditData(){
  const res=await sheetsGet(CFG.sheetIdMeta,`${SHEETS.audit}!A1:H5000`);
  const rows=res.values||[];
  if(rows.length<=1){auditLog=[];return;}
  auditLog=rows.slice(1).map(r=>({
    timestamp:r[0]||'',agent_email:r[1]||'',agent_name:r[2]||'',
    regno:r[3]||'',student_name:r[4]||'',field:r[5]||'',old_value:r[6]||'',new_value:r[7]||''
  })).reverse();
  document.getElementById('audit-count').textContent=auditLog.length;
  const emails=[...new Set(auditLog.map(a=>a.agent_email).filter(Boolean))];
  const sel=document.getElementById('audit-filter-user');
  sel.innerHTML='<option value="">All Users</option>'+emails.map(e=>`<option value="${e}">${e}</option>`).join('');
}

async function loadUsers(){
  const res=await sheetsGet(CFG.sheetIdMeta,`${SHEETS.users}!A1:E500`);
  const rows=res.values||[];
  if(rows.length<=1){users=[];return;}
  users=rows.slice(1).map(r=>({email:r[0]||'',role:r[1]||'agent',added_by:r[2]||'',added_on:r[3]||'',manager:r[4]||''})).filter(u=>u.email);
}

async function loadAccessRequests(){
  const res=await sheetsGet(CFG.sheetIdMeta,`${SHEETS.requests}!A1:E500`);
  const rows=res.values||[];
  if(rows.length<=1){accessRequests=[];return;}
  accessRequests=rows.slice(1).map((r,i)=>({
    email:r[0]||'',name:r[1]||'',manager_email:r[2]||'',status:r[3]||'pending',submitted_on:r[4]||'',_rowIndex:i+2
  }));
  const pending=accessRequests.filter(r=>r.status==='pending').length;
  const badge=document.getElementById('requests-count');
  if(badge)badge.textContent=pending;
}

async function writeRowToSheet(rowIndex,record,changes){
  const row=Array(NUM_COLS).fill('');
  Object.entries(COLUMNS).forEach(([k,ci])=>row[ci]=record[k]||'');
  await sheetsUpdate(CFG.sheetIdMain,`${SHEETS.main}!A${rowIndex}:V${rowIndex}`,[row]);
  if(changes.length){
    const now=nowStr();
    await sheetsAppend(CFG.sheetIdMeta,SHEETS.audit,
      changes.map(c=>[now,currentUser.email,currentUser.name||currentUser.email,record.regno,record.student_name,c.field,c.old,c.new])
    );
  }
  cache.clear();
}

// ════════════════════════════════════════════════════════
// FILTERS
// ════════════════════════════════════════════════════════
function buildFilters(){
  allRegions=unique(allData.map(r=>r.region));
  allCenters=unique(allData.map(r=>r.center));
  regionCenterMap={};
  allData.forEach(r=>{if(r.region&&r.center){if(!regionCenterMap[r.region])regionCenterMap[r.region]=new Set();regionCenterMap[r.region].add(r.center);}});
  fillSelect('f-region',allRegions,'All Regions');
  fillSelect('f-center',allCenters,'All Centers');
  buildPaymentDropdown(unique(allData.map(r=>r.newpayment_checks)));
}
function fillSelect(id,opts,lbl){const el=document.getElementById(id);const cur=el.value;el.innerHTML=`<option value="">${lbl}</option>`+opts.filter(Boolean).map(o=>`<option value="${o}">${o}</option>`).join('');if(cur)el.value=cur;}
function onRegionChange(){const sr=getValue('f-region');if(!sr){fillSelect('f-center',allCenters,'All Centers');}else{const cir=regionCenterMap[sr]?Array.from(regionCenterMap[sr]):[];fillSelect('f-center',cir,'All Centers');}applyFilters();}
let filterTimeout;
function applyFilters(){
  clearTimeout(filterTimeout);
  filterTimeout=setTimeout(()=>{
    const region=getValue('f-region'),center=getValue('f-center'),form=getValue('f-form'),connected=getValue('f-connected'),dialled=getValue('f-dialled'),search=getValue('f-search').toLowerCase();
    filteredData=allData.filter(r=>{
      if(region&&r.region!==region)return false;
      if(center&&r.center!==center)return false;
      if(selectedPayments.size>0&&!selectedPayments.has(r.newpayment_checks))return false;
      if(form&&r.form_status!==form)return false;
      if(connected&&r.connected_status!==connected)return false;
      if(dialled&&r.dialled_status!==dialled)return false;
      if(search&&!r.regno.toLowerCase().includes(search)&&!r.student_name.toLowerCase().includes(search))return false;
      return true;
    });
    renderTable();updateStats();
  },150);
}
function clearFilters(){
  ['f-region','f-center','f-form','f-connected','f-dialled'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('f-search').value='';
  fillSelect('f-center',allCenters,'All Centers');
  selectedPayments.clear();
  document.querySelectorAll('.ms-pay-check').forEach(el=>el.textContent='');
  const ac=document.getElementById('ms-check-all');if(ac)ac.textContent='✓';
  updatePaymentLabel();applyFilters();
}
function updateStats(){
  const conn=filteredData.filter(r=>r.connected_status==='Connected').length;
  const ptp=filteredData.filter(r=>r.ptp).length;
  const dialled=filteredData.filter(r=>r.dialled_status==='Dialled').length;
  animateCount('s-total',filteredData.length);
  animateCount('s-conn',conn);
  animateCount('s-ptp',ptp);
  animateCount('s-dialled',dialled);
  document.getElementById('tw-meta').textContent=`${filteredData.length} / ${allData.length} records`;
}

function animateCount(id,target){
  const el=document.getElementById(id);if(!el)return;
  const start=parseInt(el.textContent)||0;
  if(start===target){el.textContent=target;return;}
  const step=Math.ceil(Math.abs(target-start)/12);
  let cur=start;
  const interval=setInterval(()=>{
    cur+=cur<target?Math.min(step,target-cur):-Math.min(step,cur-target);
    el.textContent=cur;
    if(cur===target)clearInterval(interval);
  },30);
}

// ════════════════════════════════════════════════════════
// DOWNLOAD CSV
// ════════════════════════════════════════════════════════
function downloadFilteredData(){
  if(!filteredData.length){showToast('No data to download','err');return;}
  const headers=['Reg No','Student Name','Center','Region','Payment Check','Fees Amt','Fees Paid','Form Status','1st EMI Paid','2nd EMI Paid','3rd EMI Paid','4th EMI Paid','Total Paid Date','15 Days Attendance','Last Punch Date','PTP','Connected Status','Dialled Status','Yesterday Disposition',"Today's Disposition",'Other Remarks','Due Date'];
  const csvRows=[headers.join(',')];
  filteredData.forEach(r=>{
    const row=[r.regno,r.student_name,r.center,r.region,r.newpayment_checks,r.fees_amt,r.fees_paid,r.form_status,r.emi1_paid_date,r.emi2_paid_date,r.emi3_paid_date,r.emi4_paid_date,r.total_paid_date,r.attendance_15days,r.last_punch_date,r.ptp,r.connected_status,r.dialled_status,r.yesterday_disposition,r.todays_disposition,r.other_remarks,r.due_date]
      .map(v=>{const val=(v||'').toString().replace(/"/g,'""');return val.includes(',')||val.includes('"')||val.includes('\n')?`"${val}"`:val;});
    csvRows.push(row.join(','));
  });
  const blob=new Blob(['\uFEFF'+csvRows.join('\r\n')],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');
  a.href=url;a.download=`tracker_export_${nowStr().replace(/[: ]/g,'_')}.csv`;
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  showToast(`✅ Downloaded ${filteredData.length} records`,'ok');
}

// ════════════════════════════════════════════════════════
// TABLE
// ════════════════════════════════════════════════════════
function renderTable(){
  const tbody=document.getElementById('tbl-body');
  if(!filteredData.length){
    tbody.innerHTML=`<tr><td colspan="22"><div class="empty"><div class="empty-icon">🔍</div><div class="empty-t">No records found</div><div class="empty-s">Try adjusting your filters</div></div></td></tr>`;
    return;
  }
  tbody.innerHTML=filteredData.map(r=>{
    const idx=allData.indexOf(r);
    const sel=selectedIdx===idx?'selected':'';
    const dueCls=getDueDateClass(r.due_date);
    return`<tr onclick="openPanel(${idx})" class="${sel}" style="animation:rowIn .2s ease backwards">
      <td class="regno-cell">${r.regno}</td>
      <td class="name-cell">${r.student_name}</td>
      <td>${r.center||'—'}</td>
      <td>${bBadge(r.region,'blue')}</td>
      <td>${payBadge(r.newpayment_checks)}</td>
      <td class="mono-cell">${r.fees_amt||'—'}</td>
      <td class="mono-cell">${r.fees_paid||'—'}</td>
      <td class="mono-cell">${r.emi1_paid_date||'—'}</td>
      <td class="mono-cell">${r.emi2_paid_date||'—'}</td>
      <td class="mono-cell">${r.emi3_paid_date||'—'}</td>
      <td class="mono-cell">${r.emi4_paid_date||'—'}</td>
      <td class="mono-cell">${r.total_paid_date||'—'}</td>
      <td>${formBadge(r.form_status)}</td>
      <td class="mono-cell">${r.attendance_15days||'—'}</td>
      <td class="mono-cell">${r.last_punch_date||'—'}</td>
      <td>${r.ptp?bBadge('✓','green'):'—'}</td>
      <td>${connBadge(r.connected_status)}</td>
      <td>${dialBadge(r.dialled_status)}</td>
      <td>${dispBadge(r.yesterday_disposition)}</td>
      <td>${dispBadge(r.todays_disposition)}</td>
      <td class="${dueCls}">${dueDateBadge(r.due_date)}</td>
      <td><button class="btn btn-ghost btn-sm edit-btn" onclick="event.stopPropagation();openPanel(${idx})">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        Edit
      </button></td>
    </tr>`;
  }).join('');
}

function getDueDateClass(d){
  if(!d)return'';
  try{const due=new Date(d),now=new Date();now.setHours(0,0,0,0);due.setHours(0,0,0,0);const diff=(due-now)/864e5;if(diff<0)return'due-overdue';if(diff<=3)return'due-soon';return'due-ok';}
  catch(e){return'';}
}

function dueDateBadge(d){
  if(!d)return'—';
  try{
    const due=new Date(d),now=new Date();
    now.setHours(0,0,0,0);due.setHours(0,0,0,0);
    const diff=(due-now)/864e5;
    let cls='gray',label=d,icon='📅';
    if(diff<0){cls='red';icon='🔴';label=`${d} (${Math.abs(Math.round(diff))}d overdue)`;}
    else if(diff===0){cls='orange';icon='🟠';label=`${d} (Today!)`;}
    else if(diff<=3){cls='yellow';icon='🟡';label=`${d} (${Math.round(diff)}d left)`;}
    else{cls='green';icon='🟢';label=d;}
    return`<span class="badge b-${cls}" style="font-size:10px">${icon} ${label}</span>`;
  }catch(e){return d;}
}

const bBadge=(v,c)=>v?`<span class="badge b-${c}">${v}</span>`:`<span class="null-val">—</span>`;
function payBadge(v){const m={'Full Paid':'green','50% Paid':'yellow','Token Paid':'blue','Less than Token':'orange'};return bBadge(v,m[v]||'gray');}
function connBadge(v){const m={'Connected':'green','Not Connected':'red','Busy':'yellow','No Answer':'orange','Switched Off':'red','Invalid Number':'red'};return bBadge(v,m[v]||'gray');}
function dialBadge(v){return bBadge(v,v==='Dialled'?'green':'gray');}
function formBadge(v){const m={'Stage 1':'blue','Stage 2':'yellow','Stage 3':'orange','Completed':'green'};return bBadge(v,m[v]||'gray');}
function dispBadge(v){const m={'Interested':'green','Enrolled':'green','Not Interested':'red','Drop':'red','Follow Up':'yellow','Callback Requested':'yellow','RNR':'orange','Wrong Number':'red','Future Prospect':'violet'};return bBadge(v,m[v]||'gray');}

// ════════════════════════════════════════════════════════
// PANEL — DUE DATE READ-ONLY
// ════════════════════════════════════════════════════════
function openPanel(idx){
  selectedIdx=idx;
  const r=allData[idx];

  document.getElementById('p-name').textContent=r.student_name;
  document.getElementById('p-regno').textContent='#'+r.regno;

  // Build due date display
  const dueDateHtml=r.due_date?`<div class="due-display">${dueDateBadge(r.due_date)}</div>`:`<span class="null-val">Not set</span>`;

  document.getElementById('p-info').innerHTML=`
    <div class="info-cell"><div class="info-k">Reg No</div><div class="info-v">${r.regno}</div></div>
    <div class="info-cell"><div class="info-k">Center</div><div class="info-v">${r.center||'—'}</div></div>
    <div class="info-cell"><div class="info-k">Region</div><div class="info-v">${r.region||'—'}</div></div>
    <div class="info-cell"><div class="info-k">Payment</div><div class="info-v">${r.newpayment_checks||'—'}</div></div>
    <div class="info-cell"><div class="info-k">Fees Amount</div><div class="info-v mono">₹${r.fees_amt||'0'}</div></div>
    <div class="info-cell"><div class="info-k">Fees Paid</div><div class="info-v mono">₹${r.fees_paid||'0'}</div></div>
    <div class="info-cell"><div class="info-k">Form Status</div><div class="info-v">${formBadge(r.form_status)}</div></div>
    <div class="info-cell"><div class="info-k">Attendance</div><div class="info-v">${r.attendance_15days||'—'}</div></div>
    <div class="info-cell"><div class="info-k">Last Punch</div><div class="info-v mono">${r.last_punch_date||'—'}</div></div>
    <div class="info-cell"><div class="info-k">1st EMI</div><div class="info-v mono">${r.emi1_paid_date||'—'}</div></div>
    <div class="info-cell"><div class="info-k">2nd EMI</div><div class="info-v mono">${r.emi2_paid_date||'—'}</div></div>
    <div class="info-cell"><div class="info-k">3rd EMI</div><div class="info-v mono">${r.emi3_paid_date||'—'}</div></div>
    <div class="info-cell"><div class="info-k">4th EMI</div><div class="info-v mono">${r.emi4_paid_date||'—'}</div></div>
    <div class="info-cell"><div class="info-k">Total Paid Date</div><div class="info-v mono">${r.total_paid_date||'—'}</div></div>
    <div class="info-cell" style="grid-column:span 2"><div class="info-k">📅 Due Date</div><div class="info-v">${dueDateHtml}</div></div>`;

  setValue('e-ptp',r.ptp);
  setValue('e-connected',r.connected_status);
  setValue('e-dialled',r.dialled_status);
  setValue('e-today',r.todays_disposition);
  document.getElementById('e-remarks').value=r.other_remarks||'';

  const hist=auditLog.filter(a=>a.regno===r.regno);
  document.getElementById('p-history').innerHTML=!hist.length
    ?'<div class="no-history">No changes recorded yet</div>'
    :hist.slice(0,10).map(a=>`
      <div class="audit-item">
        <div class="audit-head">
          <span class="audit-who">👤 ${a.agent_name||a.agent_email}</span>
          <span class="audit-when">${a.timestamp}</span>
        </div>
        <div class="audit-row">
          <span class="audit-field">${a.field}:</span>
          <span class="audit-from">${a.old_value||'(empty)'}</span>
          <span class="audit-arr">→</span>
          <span class="audit-to">${a.new_value}</span>
        </div>
      </div>`).join('');

  document.getElementById('panel-ov').classList.add('open');
  document.getElementById('panel').classList.add('open');
  renderTable();
}

function closePanel(){
  document.getElementById('panel-ov').classList.remove('open');
  document.getElementById('panel').classList.remove('open');
  selectedIdx=null;
}

// ════════════════════════════════════════════════════════
// SAVE (no due_date field — it's read-only now)
// ════════════════════════════════════════════════════════
async function saveChanges(){
  if(selectedIdx===null)return;
  const r=allData[selectedIdx];
  const newVals={
    ptp:document.getElementById('e-ptp').value.trim(),
    connected_status:getValue('e-connected'),
    dialled_status:getValue('e-dialled'),
    todays_disposition:getValue('e-today'),
    other_remarks:document.getElementById('e-remarks').value.trim()
    // due_date intentionally omitted — read-only
  };
  const labels={ptp:'PTP',connected_status:'Connected Status',dialled_status:'Dialled Status',todays_disposition:"Today's Disposition",other_remarks:'Other Remarks'};
  const changes=Object.keys(newVals).filter(k=>(r[k]||'')!==(newVals[k]||'')).map(k=>({key:k,field:labels[k],old:r[k]||'',new:newVals[k]}));

  if(!changes.length){showToast('No changes to save','info');return;}

  const updated={...r};
  if(newVals.todays_disposition&&newVals.todays_disposition!==r.todays_disposition&&r.todays_disposition){
    updated.yesterday_disposition=r.todays_disposition;
    changes.push({key:'yesterday_disposition',field:"Yesterday's Disposition (auto)",old:r.yesterday_disposition||'',new:r.todays_disposition});
  }
  Object.assign(updated,newVals);

  const btn=document.getElementById('save-btn');
  btn.innerHTML='<span class="btn-spinner"></span> Saving…';
  btn.disabled=true;
  showLoader('Writing to sheet…');

  try{
    await writeRowToSheet(r._rowIndex,updated,changes);
    Object.assign(allData[selectedIdx],updated);
    const now=nowStr();
    changes.forEach(c=>auditLog.unshift({timestamp:now,agent_email:currentUser.email,agent_name:currentUser.name||currentUser.email,regno:r.regno,student_name:r.student_name,field:c.field,old_value:c.old,new_value:c.new}));
    document.getElementById('audit-count').textContent=auditLog.length;
    renderTable();renderAudit();openPanel(selectedIdx);
    hideLoader();showToast(`✅ Saved ${changes.length} change${changes.length>1?'s':''}!`,'ok');
  }catch(e){hideLoader();showToast('❌ '+e.message,'err');}
  btn.innerHTML='💾 Save to Sheet';btn.disabled=false;
}

// ════════════════════════════════════════════════════════
// AUDIT, TEAM, REQUESTS, USERS
// ════════════════════════════════════════════════════════
function renderAudit(){
  const fu=getValue('audit-filter-user');
  const data=fu?auditLog.filter(a=>a.agent_email===fu):auditLog;
  const el=document.getElementById('audit-list');
  if(!data.length){el.innerHTML='<div class="empty"><div class="empty-icon">📭</div><div class="empty-t">No audit entries</div></div>';return;}
  el.innerHTML=data.slice(0,200).map(a=>`
    <div class="audit-item">
      <div class="audit-head">
        <span class="audit-who">👤 ${a.agent_name||a.agent_email}</span>
        <span class="audit-when">${a.timestamp}</span>
      </div>
      <div class="audit-regno">📋 Reg: ${a.regno} — ${a.student_name}</div>
      <div class="audit-row">
        <span class="audit-field">${a.field}:</span>
        <span class="audit-from">${a.old_value||'(empty)'}</span>
        <span class="audit-arr">→</span>
        <span class="audit-to">${a.new_value}</span>
      </div>
    </div>`).join('');
}

function renderTeamMembers(){
  const el=document.getElementById('team-members-list');if(!el)return;
  const myEmail=currentUser.email;
  const team=users.filter(u=>u.manager&&u.manager.toLowerCase()===myEmail.toLowerCase());
  if(!team.length){el.innerHTML='<div style="font-size:13px;color:var(--ink3);padding:12px 0">No team members assigned yet</div>';return;}
  el.innerHTML=team.map(u=>`
    <div class="u-row">
      <div class="u-avatar">${(u.email[0]).toUpperCase()}</div>
      <span class="u-email">${u.email}</span>
      <span class="u-role-tag ur-${u.role}">${u.role.toUpperCase()}</span>
      <button class="btn btn-danger btn-sm" onclick="removeTeamMember('${u.email}')">Remove</button>
    </div>`).join('');
}

async function addTeamMember(){
  const email=document.getElementById('new-member-email').value.trim().toLowerCase();
  if(!email||!email.includes('@')){showToast('Valid email required','err');return;}
  if(users.find(u=>u.email===email)){showToast('User already exists','err');return;}
  showLoader('Adding member…');
  try{
    await sheetsAppend(CFG.sheetIdMeta,SHEETS.users,[[email,'agent',currentUser.email,nowStr(),currentUser.email]]);
    await loadUsers();renderTeamMembers();
    document.getElementById('new-member-email').value='';
    cache.clear();hideLoader();showToast('✅ Member added','ok');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

async function removeTeamMember(email){
  if(!confirm('Remove '+email+' from team?'))return;
  const res=await sheetsGet(CFG.sheetIdMeta,`${SHEETS.users}!A:A`);
  const rows=res.values||[];let rowNum=-1;
  for(let i=1;i<rows.length;i++){if((rows[i][0]||'').toLowerCase()===email.toLowerCase()){rowNum=i+1;break;}}
  if(rowNum<0){showToast('User not found','err');return;}
  const info=await getSheetInfo(CFG.sheetIdMeta);
  const sh=info.sheets.find(s=>s.properties.title===SHEETS.users);
  showLoader('Removing…');
  try{
    await sheetsBatch(CFG.sheetIdMeta,[{deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:rowNum-1,endIndex:rowNum}}}]);
    await loadUsers();renderTeamMembers();cache.clear();hideLoader();showToast('Removed','ok');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

function renderAccessRequests(){
  const el=document.getElementById('requests-list');if(!el)return;
  let reqs=accessRequests;
  if(currentUser.role==='manager')reqs=reqs.filter(r=>r.manager_email.toLowerCase()===currentUser.email.toLowerCase());
  const pending=reqs.filter(r=>r.status==='pending');
  if(!pending.length){el.innerHTML='<div class="empty"><div class="empty-icon">📭</div><div class="empty-t">No pending requests</div><div class="empty-s">All requests have been handled</div></div>';return;}
  el.innerHTML=pending.map(req=>`
    <div class="request-card">
      <div class="request-header">
        <div class="request-info">
          <div class="request-avatar">${(req.name||req.email)[0].toUpperCase()}</div>
          <div>
            <h4>${req.name}</h4>
            <p>${req.email}</p>
          </div>
        </div>
        <span class="request-status status-pending">⏳ Pending</span>
      </div>
      <div class="request-details">
        <div class="request-detail"><strong>Manager</strong><span>${req.manager_email}</span></div>
        <div class="request-detail"><strong>Submitted</strong><span>${req.submitted_on}</span></div>
      </div>
      <div class="request-actions">
        <button class="btn btn-success" onclick="approveRequest('${req.email}','${req.name}','${req.manager_email}')">✓ Approve Access</button>
        <button class="btn btn-danger" onclick="declineRequest('${req.email}','${req.name}')">✗ Decline</button>
      </div>
    </div>`).join('');
}

async function approveRequest(email,name,managerEmail){
  showLoader('Approving request…');
  try{
    const req=accessRequests.find(r=>r.email.toLowerCase()===email.toLowerCase()&&r.status==='pending');
    if(!req){showToast('Request not found','err');hideLoader();return;}
    await sheetsUpdate(CFG.sheetIdMeta,`${SHEETS.requests}!D${req._rowIndex}`,[['approved']]);
    await sheetsAppend(CFG.sheetIdMeta,SHEETS.users,[[email,'agent',currentUser.email,nowStr(),managerEmail]]);
    await loadUsers();await loadAccessRequests();renderAccessRequests();
    cache.clear();hideLoader();
    showToast(`✅ ${name||email} approved!`,'ok');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

async function declineRequest(email,name){
  if(!confirm('Decline access request from '+email+'?'))return;
  showLoader('Declining…');
  try{
    const req=accessRequests.find(r=>r.email.toLowerCase()===email.toLowerCase()&&r.status==='pending');
    if(!req){showToast('Request not found','err');hideLoader();return;}
    await sheetsUpdate(CFG.sheetIdMeta,`${SHEETS.requests}!D${req._rowIndex}`,[['declined']]);
    await loadAccessRequests();renderAccessRequests();
    hideLoader();showToast(`Declined request from ${name||email}`,'info');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

function renderUsersList(){
  const el=document.getElementById('users-list');if(!el)return;
  if(!users.length){el.innerHTML='<div style="font-size:13px;color:var(--ink3);padding:8px 0">No users yet</div>';return;}
  el.innerHTML=users.map(u=>`
    <div class="u-row">
      <div class="u-avatar">${(u.email[0]).toUpperCase()}</div>
      <span class="u-email">${u.email}</span>
      <span class="u-role-tag ur-${u.role}">${u.role.toUpperCase()}</span>
      ${currentUser?.role==='admin'?`<button class="btn btn-danger btn-sm" onclick="removeUser('${u.email}')">Remove</button>`:''}
    </div>`).join('');
}

async function addUser(){
  const email=document.getElementById('new-email').value.trim().toLowerCase();
  const role=document.getElementById('new-role').value;
  if(!email||!email.includes('@')){showToast('Valid email required','err');return;}
  if(users.find(u=>u.email===email)){showToast('User already exists','err');return;}
  showLoader('Adding user…');
  try{
    await sheetsAppend(CFG.sheetIdMeta,SHEETS.users,[[email,role,currentUser.email,nowStr(),'']]);
    await loadUsers();renderUsersList();
    document.getElementById('new-email').value='';
    cache.clear();hideLoader();showToast('✅ Added: '+email,'ok');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

async function removeUser(email){
  if(!confirm('Remove '+email+' from system?'))return;
  const res=await sheetsGet(CFG.sheetIdMeta,`${SHEETS.users}!A:A`);
  const rows=res.values||[];let rowNum=-1;
  for(let i=1;i<rows.length;i++){if((rows[i][0]||'').toLowerCase()===email.toLowerCase()){rowNum=i+1;break;}}
  if(rowNum<0){showToast('User not found','err');return;}
  const info=await getSheetInfo(CFG.sheetIdMeta);
  const sh=info.sheets.find(s=>s.properties.title===SHEETS.users);
  showLoader('Removing…');
  try{
    await sheetsBatch(CFG.sheetIdMeta,[{deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:rowNum-1,endIndex:rowNum}}}]);
    await loadUsers();renderUsersList();cache.clear();hideLoader();showToast('User removed','ok');
  }catch(e){hideLoader();showToast('Error: '+e.message,'err');}
}

// ════════════════════════════════════════════════════════
// NAV & UI
// ════════════════════════════════════════════════════════
function toggleSidebar(){
  const s=document.querySelector('.sidebar');
  const c=s.classList.toggle('collapsed');
  localStorage.setItem('crm_sidebar_collapsed',c?'true':'false');
}

function showPage(name,btn){
  ['tracker','audit','members','requests','settings'].forEach(p=>{
    const el=document.getElementById('page-'+p);
    if(el)el.style.display=p===name?'':'none';
  });
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  if(btn)btn.classList.add('active');
  const titles={tracker:'📊 Tracker',audit:'🕐 Audit Log',members:'👥 My Team',requests:'📬 Access Requests',settings:'⚙️ Settings'};
  document.getElementById('page-title').textContent=titles[name]||name;
  if(name==='audit')renderAudit();
  if(name==='members')renderTeamMembers();
  if(name==='requests')renderAccessRequests();
  if(name==='settings')renderUsersList();
}

// ════════════════════════════════════════════════════════
// UTILS
// ════════════════════════════════════════════════════════
function getValue(id){return document.getElementById(id)?.value||'';}
function setValue(id,val){const el=document.getElementById(id);if(el)el.value=val||'';}
function unique(arr){return[...new Set(arr.filter(Boolean))];}
function nowStr(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')+' '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+':'+String(d.getSeconds()).padStart(2,'0');}
function showLoader(msg){document.getElementById('loader-msg').textContent=msg||'Loading…';document.getElementById('loader').classList.add('on');}
function hideLoader(){document.getElementById('loader').classList.remove('on');}
function showAuthErr(msg){const el=document.getElementById('auth-err');el.textContent=msg;el.classList.add('show');}
function showSignupErr(msg){const el=document.getElementById('signup-err');el.textContent=msg;el.classList.add('show');}

let toastQueue=[];let toastShowing=false;
function showToast(msg,type='info'){
  const el=document.getElementById('toast');
  el.textContent=msg;
  el.className='toast on '+(type||'');
  clearTimeout(window._toastTimeout);
  window._toastTimeout=setTimeout(()=>el.classList.remove('on'),3400);
}

// Auto-refresh every 5 minutes
setInterval(()=>{
  if(currentUser&&document.getElementById('app').style.display==='flex'){
    const btn=document.getElementById('sync-btn');
    btn.classList.add('force-refresh');
    loadAll();
  }
},5*60*1000);
