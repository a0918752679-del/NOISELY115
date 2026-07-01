(function(){
  const TOKEN_KEY='ntpc_noise_dashboard_management_token_v1';
  function h(base={}){const token=localStorage.getItem(TOKEN_KEY)||'';return token?{...base,'X-Dashboard-Token':token}:base}
  function esc(s){return String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
  function css(){
    const st=document.createElement('style');
    st.textContent=`
    .ai-fab{position:fixed;right:18px;bottom:18px;z-index:9998;border:0;border-radius:999px;background:#0f4c81;color:white;font-weight:900;padding:13px 16px;box-shadow:0 14px 36px rgba(15,76,129,.28);cursor:pointer;letter-spacing:.02em}
    .ai-panel{position:fixed;right:18px;bottom:76px;width:min(430px,calc(100vw - 28px));max-height:min(680px,calc(100vh - 100px));z-index:9999;background:#f8fbff;border:1px solid rgba(99,139,178,.35);border-radius:20px;box-shadow:0 24px 70px rgba(22,45,78,.28);display:none;overflow:hidden;font-family:inherit;color:#10243f}
    .ai-panel.show{display:flex;flex-direction:column}.ai-head{padding:14px 16px;background:linear-gradient(135deg,#0f4c81,#1f78b4);color:#fff;display:flex;justify-content:space-between;gap:12px;align-items:center}.ai-head b{font-size:16px}.ai-head small{display:block;opacity:.9;margin-top:2px}.ai-close{border:0;background:rgba(255,255,255,.18);color:#fff;border-radius:10px;padding:5px 9px;cursor:pointer}.ai-body{padding:14px;overflow:auto}.ai-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}.ai-btn{border:1px solid rgba(93,128,165,.35);background:#fff;border-radius:12px;padding:10px 8px;font-weight:800;color:#0f4c81;cursor:pointer}.ai-btn:hover{background:#eef7ff}.ai-q{display:flex;gap:8px;margin:10px 0}.ai-q input{flex:1;border:1px solid rgba(93,128,165,.35);border-radius:12px;padding:10px}.ai-q button{border:0;background:#0f4c81;color:#fff;border-radius:12px;padding:0 12px;font-weight:900;cursor:pointer}.ai-out{white-space:pre-wrap;background:#fff;border:1px solid rgba(93,128,165,.24);border-radius:14px;padding:12px;line-height:1.55;font-size:13px;max-height:430px;overflow:auto}.ai-note{font-size:12px;color:#64748b;margin:8px 0 0}.ai-loading{opacity:.7}`;
    document.head.appendChild(st);
  }
  function mount(){
    if(document.getElementById('aiFab'))return;
    css();
    const fab=document.createElement('button');fab.id='aiFab';fab.className='ai-fab';fab.textContent='AI 助理';
    const panel=document.createElement('section');panel.id='aiPanel';panel.className='ai-panel';
    panel.innerHTML=`<div class="ai-head"><div><b>AI 成效分析助理</b><small>摘要、健檢、排場建議、自然語言查詢</small></div><button class="ai-close" type="button">×</button></div><div class="ai-body"><div class="ai-grid"><button class="ai-btn" data-task="report">主管摘要</button><button class="ai-btn" data-task="recommend">排場建議</button><button class="ai-btn" data-task="validate">資料健檢</button><button class="ai-btn" data-task="anomaly">異常清單</button></div><div class="ai-q"><input id="aiQuestion" placeholder="例如：幫我比較淡水區和板橋區成效"><button id="aiAsk" type="button">送出</button></div><div id="aiOutput" class="ai-out">請先登入管理 Token，再使用 AI 功能。若環境變數 AI_REQUIRE_ROLE=viewer，前台也可使用查詢。</div><div class="ai-note">AI 回覆依目前平台/Google Sheet 資料產生，正式告發、通檢與裁罰仍以承辦複核為準。</div></div>`;
    document.body.appendChild(fab);document.body.appendChild(panel);
    const out=panel.querySelector('#aiOutput');
    function toggle(){panel.classList.toggle('show')}
    fab.addEventListener('click',toggle);panel.querySelector('.ai-close').addEventListener('click',toggle);
    async function call(task,body={}){
      out.classList.add('ai-loading');out.textContent='AI 分析中…';
      const map={report:'/api/ai/report',recommend:'/api/ai/recommend-locations',validate:'/api/ai/validate-import',anomaly:'/api/ai/anomaly-check',query:'/api/ai/query'};
      try{
        const res=await fetch(map[task],{method:'POST',headers:h({'Content-Type':'application/json','Accept':'application/json'}),body:JSON.stringify(body),cache:'no-store'});
        const json=await res.json().catch(()=>({}));
        if(!res.ok||json.ok===false)throw new Error(json.message||`HTTP ${res.status}`);
        out.textContent=json.report||json.answer||json.message||(json.anomalies?JSON.stringify(json.anomalies,null,2):JSON.stringify(json,null,2));
      }catch(e){out.innerHTML=`<b>AI 功能執行失敗</b>\n${esc(e.message||e)}\n\n請確認 Zeabur 環境變數：AI_ENABLED、OPENAI_API_KEY 或 AI_PROVIDER=heuristic，以及 DASHBOARD_ADMIN_TOKEN / DASHBOARD_EDITOR_TOKEN。`}
      finally{out.classList.remove('ai-loading')}
    }
    panel.querySelectorAll('.ai-btn').forEach(btn=>btn.addEventListener('click',()=>call(btn.dataset.task,{style:'主管簡報版'})));
    panel.querySelector('#aiAsk').addEventListener('click',()=>{const q=panel.querySelector('#aiQuestion').value.trim();if(!q){out.textContent='請輸入問題。';return}call('query',{question:q})});
    panel.querySelector('#aiQuestion').addEventListener('keydown',e=>{if(e.key==='Enter')panel.querySelector('#aiAsk').click()});
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();
})();
