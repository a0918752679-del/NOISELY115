(function(){
  const TOKEN_KEY='ntpc_noise_dashboard_management_token_v1';
  const ROLE_LEVEL={viewer:0,editor:1,admin:2};
  let currentRole='viewer';
  let aiRequireRole='editor';
  let aiStatusLoaded=false;
  function getToken(){return (localStorage.getItem(TOKEN_KEY)||'').trim()}
  function h(base={}){const token=getToken();return token?{...base,'X-Dashboard-Token':token}:base}
  function esc(s){return String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
  function roleLabel(role){return role==='admin'?'系統管理者':role==='editor'?'資料編輯者':'前台瀏覽'}
  function roleOk(role=currentRole){return ROLE_LEVEL[role] >= ROLE_LEVEL[aiRequireRole || 'editor']}
  function css(){
    const st=document.createElement('style');
    st.textContent=`
    .ai-fab{position:fixed;right:18px;bottom:18px;z-index:9998;border:0;border-radius:999px;background:#0f4c81;color:white;font-weight:900;padding:13px 16px;box-shadow:0 14px 36px rgba(15,76,129,.28);cursor:pointer;letter-spacing:.02em}
    .ai-panel{position:fixed;right:18px;bottom:76px;width:min(430px,calc(100vw - 28px));max-height:min(720px,calc(100vh - 100px));z-index:9999;background:#f8fbff;border:1px solid rgba(99,139,178,.35);border-radius:20px;box-shadow:0 24px 70px rgba(22,45,78,.28);display:none;overflow:hidden;font-family:inherit;color:#10243f}
    .ai-panel.show{display:flex;flex-direction:column}.ai-head{padding:14px 16px;background:linear-gradient(135deg,#0f4c81,#1f78b4);color:#fff;display:flex;justify-content:space-between;gap:12px;align-items:center}.ai-head b{font-size:16px}.ai-head small{display:block;opacity:.9;margin-top:2px}.ai-close{border:0;background:rgba(255,255,255,.18);color:#fff;border-radius:10px;padding:5px 9px;cursor:pointer}.ai-body{padding:14px;overflow:auto}
    .ai-auth{background:#fff;border:1px solid rgba(93,128,165,.24);border-radius:14px;padding:10px;margin-bottom:10px}.ai-auth-title{font-size:13px;font-weight:900;color:#0f4c81;margin-bottom:7px}.ai-auth-row{display:grid;grid-template-columns:1fr auto auto;gap:7px;align-items:center}.ai-auth-row input{min-width:0;border:1px solid rgba(93,128,165,.35);border-radius:11px;padding:10px;background:#fff}.ai-auth-row button{border:0;border-radius:11px;padding:10px 10px;font-weight:900;cursor:pointer}.ai-login-btn{background:#0f4c81;color:#fff}.ai-logout-btn{background:#eef2f7;color:#334155}.ai-status{font-size:12px;color:#64748b;margin-top:7px;line-height:1.45}.ai-status.ok{color:#166534}.ai-status.warn{color:#b45309}.ai-status.err{color:#b91c1c}
    .ai-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}.ai-btn{border:1px solid rgba(93,128,165,.35);background:#fff;border-radius:12px;padding:10px 8px;font-weight:800;color:#0f4c81;cursor:pointer}.ai-btn:hover{background:#eef7ff}.ai-btn:disabled{opacity:.5;cursor:not-allowed;background:#f1f5f9}.ai-q{display:flex;gap:8px;margin:10px 0}.ai-q input{flex:1;min-width:0;border:1px solid rgba(93,128,165,.35);border-radius:12px;padding:10px}.ai-q button{border:0;background:#0f4c81;color:#fff;border-radius:12px;padding:0 12px;font-weight:900;cursor:pointer}.ai-q button:disabled{opacity:.5;cursor:not-allowed}.ai-out{white-space:pre-wrap;background:#fff;border:1px solid rgba(93,128,165,.24);border-radius:14px;padding:12px;line-height:1.55;font-size:13px;max-height:430px;overflow:auto}.ai-note{font-size:12px;color:#64748b;margin:8px 0 0}.ai-loading{opacity:.7}
    @media(max-width:520px){.ai-auth-row{grid-template-columns:1fr}.ai-auth-row button{width:100%}.ai-q{display:grid;grid-template-columns:1fr auto}.ai-panel{right:10px;bottom:72px;width:calc(100vw - 20px)}}`;
    document.head.appendChild(st);
  }
  function mount(){
    if(document.getElementById('aiFab'))return;
    css();
    const fab=document.createElement('button');fab.id='aiFab';fab.className='ai-fab';fab.textContent='AI 助理';
    const panel=document.createElement('section');panel.id='aiPanel';panel.className='ai-panel';
    panel.innerHTML=`<div class="ai-head"><div><b>AI 成效分析助理</b><small>摘要、健檢、排場建議、自然語言查詢</small></div><button class="ai-close" type="button">×</button></div><div class="ai-body"><div class="ai-auth"><div class="ai-auth-title">AI 功能登入</div><div class="ai-auth-row"><input id="aiToken" type="password" autocomplete="current-password" placeholder="請輸入 AI 使用密碼 / 管理密碼"><button id="aiLogin" class="ai-login-btn" type="button">登入</button><button id="aiLogout" class="ai-logout-btn" type="button">登出</button></div><div id="aiAuthStatus" class="ai-status">尚未登入。請先輸入 AI 使用密碼後再執行 AI 功能。</div></div><div class="ai-grid"><button class="ai-btn" data-task="report">主管摘要</button><button class="ai-btn" data-task="recommend">排場建議</button><button class="ai-btn" data-task="validate">資料健檢</button><button class="ai-btn" data-task="anomaly">異常清單</button></div><div class="ai-q"><input id="aiQuestion" placeholder="例如：幫我比較淡水區和板橋區成效"><button id="aiAsk" type="button">送出</button></div><div id="aiOutput" class="ai-out">請先在上方輸入 AI 使用密碼並登入，再使用 AI 功能。若環境變數 AI_REQUIRE_ROLE=viewer，前台也可直接使用查詢。</div><div class="ai-note">AI 回覆依目前平台/Google Sheet 資料產生，正式告發、通檢與裁罰仍以承辦複核為準。</div></div>`;
    document.body.appendChild(fab);document.body.appendChild(panel);
    const out=panel.querySelector('#aiOutput');
    const tokenInput=panel.querySelector('#aiToken');
    const loginBtn=panel.querySelector('#aiLogin');
    const logoutBtn=panel.querySelector('#aiLogout');
    const statusEl=panel.querySelector('#aiAuthStatus');
    const actionBtns=Array.from(panel.querySelectorAll('.ai-btn'));
    const askBtn=panel.querySelector('#aiAsk');
    function setStatus(text,type='warn'){statusEl.className=`ai-status ${type}`;statusEl.textContent=text;}
    function setControls(){
      const needsLogin=aiRequireRole!=='viewer';
      const usable=!needsLogin || roleOk();
      actionBtns.forEach(btn=>{btn.disabled=!usable});
      askBtn.disabled=!usable;
      logoutBtn.disabled=!Boolean(getToken());
    }
    function renderAuthState(){
      const token=getToken();
      if(token && !tokenInput.value) tokenInput.value=token;
      if(aiRequireRole==='viewer'){
        setStatus('AI 目前設定為前台可用；管理密碼可留空。','ok');
      }else if(token && roleOk()){
        setStatus(`已登入：${roleLabel(currentRole)}。AI 所需權限：${roleLabel(aiRequireRole)}。`,'ok');
      }else if(token){
        setStatus(`已輸入密碼，但目前權限不足或尚未驗證。AI 所需權限：${roleLabel(aiRequireRole)}。`,'warn');
      }else{
        setStatus(`尚未登入。AI 所需權限：${roleLabel(aiRequireRole)}，請輸入 AI 使用密碼。`,'warn');
      }
      setControls();
    }
    async function loadAIStatus(){
      try{
        const res=await fetch('/api/ai/status',{headers:{'Accept':'application/json'},cache:'no-store'});
        const obj=await res.json().catch(()=>({}));
        if(res.ok && obj.ok){aiRequireRole=obj.requireRole||aiRequireRole;aiStatusLoaded=true;}
      }catch(e){}
      renderAuthState();
    }
    async function verifyToken(token,showOk=true){
      const res=await fetch('/api/auth/check',{headers:{'Accept':'application/json','X-Dashboard-Token':token},cache:'no-store'});
      const obj=await res.json().catch(()=>({}));
      if(!res.ok)throw new Error(obj.message||`HTTP ${res.status}`);
      currentRole=obj.role||'viewer';
      localStorage.setItem(TOKEN_KEY,token);
      if(showOk) out.textContent=`登入成功：${roleLabel(currentRole)}。可以開始使用 AI 功能。`;
      renderAuthState();
      return obj;
    }
    async function login(){
      const token=String(tokenInput.value||'').trim();
      if(!token){setStatus('請先輸入 AI 使用密碼。','err');tokenInput.focus();return;}
      loginBtn.disabled=true;setStatus('正在驗證密碼…','warn');
      try{await verifyToken(token,true)}
      catch(e){currentRole='viewer';localStorage.removeItem(TOKEN_KEY);setStatus(`登入失敗：${e.message||e}`,'err');out.textContent='AI 登入失敗，請確認密碼是否與 Zeabur 環境變數 DASHBOARD_ADMIN_TOKEN 或 DASHBOARD_EDITOR_TOKEN 一致。'}
      finally{loginBtn.disabled=false;renderAuthState();}
    }
    function logout(){
      currentRole='viewer';localStorage.removeItem(TOKEN_KEY);tokenInput.value='';out.textContent='已登出 AI 功能。請重新輸入 AI 使用密碼後再執行。';renderAuthState();
    }
    function ensureCanUseAI(){
      if(aiRequireRole==='viewer')return true;
      if(getToken() && roleOk())return true;
      out.textContent='請先在上方輸入 AI 使用密碼並登入，再執行 AI 功能。';
      setStatus(`尚未登入或權限不足。AI 所需權限：${roleLabel(aiRequireRole)}。`,'err');
      tokenInput.focus();
      return false;
    }
    function toggle(){panel.classList.toggle('show');if(panel.classList.contains('show')){if(!aiStatusLoaded)loadAIStatus();renderAuthState();}}
    fab.addEventListener('click',toggle);panel.querySelector('.ai-close').addEventListener('click',toggle);
    loginBtn.addEventListener('click',login);logoutBtn.addEventListener('click',logout);
    tokenInput.addEventListener('keydown',e=>{if(e.key==='Enter')loginBtn.click()});
    async function call(task,body={}){
      if(!ensureCanUseAI())return;
      out.classList.add('ai-loading');out.textContent='AI 分析中…';
      const map={report:'/api/ai/report',recommend:'/api/ai/recommend-locations',validate:'/api/ai/validate-import',anomaly:'/api/ai/anomaly-check',query:'/api/ai/query'};
      try{
        const res=await fetch(map[task],{method:'POST',headers:h({'Content-Type':'application/json','Accept':'application/json'}),body:JSON.stringify(body),cache:'no-store'});
        const json=await res.json().catch(()=>({}));
        if(!res.ok||json.ok===false)throw new Error(json.message||`HTTP ${res.status}`);
        out.textContent=json.report||json.answer||json.message||(json.anomalies?JSON.stringify(json.anomalies,null,2):JSON.stringify(json,null,2));
      }catch(e){
        if(/權限不足|Token|401|403/.test(String(e.message||e))){currentRole='viewer';renderAuthState();}
        out.innerHTML=`<b>AI 功能執行失敗</b>\n${esc(e.message||e)}\n\n請確認：\n1. 上方 AI 使用密碼已登入。\n2. Zeabur 環境變數 AI_ENABLED=true。\n3. DASHBOARD_ADMIN_TOKEN / DASHBOARD_EDITOR_TOKEN 已設定。\n4. 使用 OpenAI 時，OPENAI_API_KEY 已設定；未設定則可用 AI_PROVIDER=heuristic。`;
      }
      finally{out.classList.remove('ai-loading')}
    }
    actionBtns.forEach(btn=>btn.addEventListener('click',()=>call(btn.dataset.task,{style:'主管簡報版'})));
    askBtn.addEventListener('click',()=>{const q=panel.querySelector('#aiQuestion').value.trim();if(!q){out.textContent='請輸入問題。';return}call('query',{question:q})});
    panel.querySelector('#aiQuestion').addEventListener('keydown',e=>{if(e.key==='Enter')askBtn.click()});
    const saved=getToken();
    if(saved){tokenInput.value=saved;loadAIStatus().then(()=>verifyToken(saved,false).catch(()=>{currentRole='viewer';localStorage.removeItem(TOKEN_KEY);tokenInput.value='';renderAuthState();}));}
    else loadAIStatus();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();
})();
