(function () {
  const NOTICE = "この金額は入力内容に基づく概算です。正式な報酬額は必要資料・申請先・業務範囲確認後に確定します。";
  const BASE = {
    "建設業許可|法人|新規|知事": 150000, "建設業許可|個人|新規|知事": 120000,
    "建設業許可|法人|更新|知事": 75000, "建設業許可|個人|更新|知事": 65000,
    "業種追加": 75000, "決算変更届": 40000, "各種変更届": 30000,
    "宅建業免許|新規": 120000, "宅建業免許|更新": 80000,
    "株式会社設立": 80000, "合同会社設立": 60000,
    "創業融資|事業計画書作成": 80000, "創業融資|面談対策込み": 120000,
    "車庫証明": 10000,
  };
  function n(v){const x=Number(v||0);return Number.isFinite(x)?x:0}
  function yen(v){return `${Math.floor(v).toLocaleString("ja-JP")} 円`;}
  function init(){
    const root=document.getElementById("estimate-calculator-root"); if(!root)return;
    root.innerHTML=`<form id="estimate-calc-form" class="form"><div class="grid cols-2">
      <label>顧客<select name="clientId"></select></label><label>案件名<input name="projectName" /></label>
      <label>業務種別<select name="workType"><option>建設業許可</option><option>業種追加</option><option>決算変更届</option><option>各種変更届</option><option>宅建業免許</option><option>株式会社設立</option><option>合同会社設立</option><option>創業融資</option><option>車庫証明</option></select></label>
      <label>申請区分<select name="applicationType"><option>新規</option><option>更新</option><option>事業計画書作成</option><option>面談対策込み</option></select></label>
      <label>法人/個人<select name="corporateType"><option>法人</option><option>個人</option></select></label><label>知事/大臣<select name="governorType"><option>知事</option><option>大臣</option></select></label>
      <label>一般/特定<select name="generalSpecific"><option>一般</option><option>特定</option></select></label><label>業種数<input name="industryCount" type="number" value="1" min="1" /></label>
      <label>役員数<input name="officerCount" type="number" value="2" min="0" /></label><label>営業所数<input name="officeCount" type="number" value="1" min="1" /></label>
      <label>書類不足レベル<select name="documentLevel"><option>低</option><option>中</option><option>高</option></select></label><label>急ぎ対応<select name="urgent"><option value="0">なし</option><option value="1">あり</option></select></label>
      <label>経管確認難易度<select name="keikan"><option>低</option><option>高</option></select></label><label>専技確認難易度<select name="sengi"><option>低</option><option>高</option></select></label>
      <label>財産要件確認難易度<select name="zaisan"><option>低</option><option>高</option></select></label><label>訪問対応<select name="visit"><option value="0">なし</option><option value="1">あり</option></select></label>
      <label>代理取得<select name="agent"><option value="0">なし</option><option value="1">あり</option></select></label><label>実費<input name="expense" type="number" value="0" min="0" /></label>
      <label>値引き<input name="discount" type="number" value="0" min="0" /></label>
    </div><label>メモ<textarea name="memo"></textarea></label>
    <div class="row-actions"><button type="button" id="calc-run" class="secondary-btn">再計算</button><button type="button" id="calc-save">保存</button><button type="button" id="calc-apply">見積へ反映</button></div>
    <div id="calc-result" class="panel"></div></form>`;
    const app=window.GyoseiApp; const form=root.querySelector('#estimate-calc-form'); const cs=form.elements.clientId;
    (app?.getClients?.()||[]).forEach(c=>{const o=document.createElement('option');o.value=c.id;o.textContent=c.name||c.companyName||c.contactName||'未設定';cs.appendChild(o)});
    const run=()=>{
      const f=form.elements; let base=0; const wt=f.workType.value;
      if(wt==='建設業許可') base=BASE[`建設業許可|${f.corporateType.value}|${f.applicationType.value}|${f.governorType.value}`]||0;
      else if(wt==='宅建業免許') base=BASE[`宅建業免許|${f.applicationType.value}`]||0;
      else if(wt==='創業融資') base=BASE[`創業融資|${f.applicationType.value}`]||0;
      else base=BASE[wt]||0;
      const addons=[]; const add=(n,a)=>{if(a>0)addons.push({name:n,amount:a})};
      add('業種加算', Math.max(0,n(f.industryCount.value)-1)*10000); add('役員加算',Math.max(0,n(f.officerCount.value)-2)*5000); add('営業所加算',Math.max(0,n(f.officeCount.value)-1)*30000);
      add('急ぎ対応', n(f.urgent.value)?30000:0); add('書類不足', f.documentLevel.value==='中'?20000:(f.documentLevel.value==='高'?40000:0)); add('経管確認', f.keikan.value==='高'?30000:0);
      add('専技確認', f.sengi.value==='高'?30000:0); add('財産要件', f.zaisan.value==='高'?20000:0); add('訪問対応', n(f.visit.value)?10000:0); add('代理取得', n(f.agent.value)?10000:0);
      const addon=addons.reduce((s,x)=>s+x.amount,0), discount=n(f.discount.value), expense=n(f.expense.value), taxable=base+addon-discount;
      const tax=Math.floor(taxable*(app?.getTaxRate?.()??0.1)), total=taxable+tax+expense;
      form.dataset.result=JSON.stringify({base,addons,addon,discount,expense,tax,taxable,total});
      root.querySelector('#calc-result').innerHTML=`<p>基本報酬: ${yen(base)}</p><p>加算明細: ${(addons.map(a=>`${a.name} ${yen(a.amount)}`).join(' / ')||'なし')}</p><p>値引き: ${yen(discount)}</p><p>実費: ${yen(expense)}</p><p>消費税: ${yen(tax)}</p><p><strong>合計: ${yen(total)}</strong></p><p class='meta'>${NOTICE}</p>`;
    };
    root.querySelector('#calc-run').addEventListener('click',run); run();
    root.querySelector('#calc-save').addEventListener('click', async()=>{ const app=window.GyoseiApp; const sb=app?.getSupabaseClient?.(); const u=app?.getCurrentUser?.(); if(!sb||!u)return;
      run(); const r=JSON.parse(form.dataset.result||'{}'); const f=form.elements;
      const payload={user_id:u.id,client_id:f.clientId.value||null,project_name:f.projectName.value||null,work_type:f.workType.value,application_type:f.applicationType.value,corporate_type:f.corporateType.value,governor_type:f.governorType.value,general_specific:f.generalSpecific.value,industry_count:n(f.industryCount.value),officer_count:n(f.officerCount.value),office_count:n(f.officeCount.value),document_level:f.documentLevel.value,urgent:n(f.urgent.value)>0,keikan_level:f.keikan.value,sengi_level:f.sengi.value,zaisan_level:f.zaisan.value,visit_required:n(f.visit.value)>0,agent_required:n(f.agent.value)>0,expense_amount:r.expense||0,discount_amount:r.discount||0,memo:f.memo.value||null,base_fee:r.base||0,addon_fee:r.addon||0,taxable_subtotal:r.taxable||0,tax:r.tax||0,total:r.total||0,addon_breakdown:JSON.stringify(r.addons||[])};
      const res=await sb.from('estimate_calculations').insert(payload); if(res.error){app.showMessage('見積自動算出の保存に失敗しました。'+res.error.message,true);return;} app.showMessage('見積自動算出を保存しました。'); });
    root.querySelector('#calc-apply').addEventListener('click', async()=>{const app=window.GyoseiApp; const sb=app?.getSupabaseClient?.(); const u=app?.getCurrentUser?.(); if(!sb||!u)return; run(); const r=JSON.parse(form.dataset.result||'{}'); const f=form.elements;
      const clients=app?.getClients?.()||[]; const selected=clients.find(c=>String(c.id)===String(f.clientId.value||''));
      const customerName=(selected?.name||selected?.companyName||selected?.contactName||f.projectName.value||'自動算出');
      const est={user_id:u.id,client_id:f.clientId.value||null,customer_name:customerName,estimate_title:f.projectName.value||'見積自動算出',estimate_date:new Date().toISOString().slice(0,10),status:'未回答',memo:f.memo.value||null,subtotal:(r.taxable||0)+(r.expense||0),tax:r.tax||0,total:(r.taxable||0)+(r.tax||0)+(r.expense||0)};
      est.estimate_number='M-AUTO-'+Date.now();
      const e=await sb.from('estimates').insert(est).select('id').single(); if(e.error){app.showMessage('見積登録に失敗しました。'+e.error.message,true);return;}
      const items=[{item_name:'基本報酬',quantity:1,unit_price:r.base||0,amount:r.base||0,sort_order:1},{item_name:'加算',quantity:1,unit_price:r.addon||0,amount:r.addon||0,sort_order:2},{item_name:'値引き',quantity:1,unit_price:-(r.discount||0),amount:-(r.discount||0),sort_order:3},{item_name:'実費(非課税)',quantity:1,unit_price:r.expense||0,amount:r.expense||0,sort_order:4}].map(i=>({...i,user_id:u.id,estimate_id:e.data.id}));
      const ir=await sb.from('estimate_items').insert(items); if(ir.error){app.showMessage('見積明細登録に失敗しました。'+ir.error.message,true);return;}
      await app.reloadAllData(); app.showMessage('見積へ反映しました。');
    });
  }
  document.addEventListener('DOMContentLoaded', init);
})();
