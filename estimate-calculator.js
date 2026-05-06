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
  function h(s){return String(s ?? "").replace(/[&<>\"']/g,(ch)=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));}
  function fmtDate(v){if(!v) return "-"; const d=new Date(v); return Number.isNaN(d.getTime())?String(v):d.toLocaleDateString("ja-JP");}

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
    <div id="calc-result" class="panel"></div></form>
    <section class="panel" id="calc-saved-list-wrap"><h3>保存済み概算見積</h3><div id="calc-saved-list"></div></section>`;

    const app=window.GyoseiApp; const form=root.querySelector('#estimate-calc-form'); const cs=form.elements.clientId;
    const savedList=root.querySelector('#calc-saved-list');
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

    const fillForm=(calc)=>{
      const f=form.elements;
      f.clientId.value=calc.client_id||""; f.projectName.value=calc.project_name||""; f.workType.value=calc.work_type||"建設業許可";
      f.applicationType.value=calc.application_type||"新規"; f.corporateType.value=calc.corporate_type||"法人"; f.governorType.value=calc.governor_type||"知事";
      f.generalSpecific.value=calc.general_specific||"一般"; f.industryCount.value=n(calc.industry_count)||1; f.officerCount.value=n(calc.officer_count)||2;
      f.officeCount.value=n(calc.office_count)||1; f.documentLevel.value=calc.document_level||"低"; f.urgent.value=calc.urgent?"1":"0";
      f.keikan.value=calc.keikan_level||"低"; f.sengi.value=calc.sengi_level||"低"; f.zaisan.value=calc.zaisan_level||"低";
      f.visit.value=calc.visit_required?"1":"0"; f.agent.value=calc.agent_required?"1":"0"; f.expense.value=n(calc.expense_amount)||0;
      f.discount.value=n(calc.discount_amount)||0; f.memo.value=calc.memo||"";
      run();
    };

    const reflectToEstimate=async(calc)=>{
      const sb=app?.getSupabaseClient?.(); const u=app?.getCurrentUser?.(); if(!sb||!u)return;
      const clients=app?.getClients?.()||[]; const selected=clients.find(c=>String(c.id)===String(calc.client_id||''));
      const customerName=(selected?.name||selected?.companyName||selected?.contactName||calc.project_name||'自動算出');
      const est={user_id:u.id,client_id:calc.client_id||null,customer_name:customerName,estimate_title:calc.project_name||'見積自動算出',estimate_date:new Date().toISOString().slice(0,10),status:'未回答',memo:calc.memo||null,subtotal:n(calc.taxable_subtotal)+n(calc.expense_amount),tax:n(calc.tax),total:n(calc.total)};
      est.estimate_number='M-AUTO-'+Date.now();
      const e=await sb.from('estimates').insert(est).select('id').single(); if(e.error){app.showMessage('見積登録に失敗しました。'+e.error.message,true);return false;}
      const items=[{item_name:'基本報酬',quantity:1,unit_price:n(calc.base_fee),amount:n(calc.base_fee),sort_order:1},{item_name:'加算',quantity:1,unit_price:n(calc.addon_fee),amount:n(calc.addon_fee),sort_order:2},{item_name:'値引き',quantity:1,unit_price:-n(calc.discount_amount),amount:-n(calc.discount_amount),sort_order:3},{item_name:'実費(非課税)',quantity:1,unit_price:n(calc.expense_amount),amount:n(calc.expense_amount),sort_order:4}].map(i=>({...i,user_id:u.id,estimate_id:e.data.id}));
      const ir=await sb.from('estimate_items').insert(items); if(ir.error){app.showMessage('見積明細登録に失敗しました。'+ir.error.message,true);return false;}
      await app.reloadAllData();
      renderSavedList();
      app.showMessage('見積へ反映しました。');
      return true;
    };

    const renderSavedList=()=>{
      const calcs=(app?.getEstimateCalculations?.()||[]).slice().sort((a,b)=>new Date(b.created_at||0)-new Date(a.created_at||0));
      if(!calcs.length){savedList.innerHTML='<p class="meta">保存済みデータはありません。</p>'; return;}
      savedList.innerHTML=`<table><thead><tr><th>作成日</th><th>顧客名</th><th>案件名</th><th>業務種別</th><th>申請区分</th><th>基本報酬</th><th>加算</th><th>実費</th><th>消費税</th><th>合計</th><th>メモ</th><th>操作</th></tr></thead><tbody>${calcs.map(c=>`<tr>
      <td>${h(fmtDate(c.created_at))}</td><td>${h(c.client_name||c.customer_name||"-")}</td><td>${h(c.project_name||"-")}</td><td>${h(c.work_type||"-")}</td><td>${h(c.application_type||"-")}</td>
      <td>${h(yen(c.base_fee||0))}</td><td>${h(yen(c.addon_fee||0))}</td><td>${h(yen(c.expense_amount||0))}</td><td>${h(yen(c.tax||0))}</td><td>${h(yen(c.total||0))}</td><td>${h(c.memo||"-")}</td>
      <td><div class="row-actions"><button type="button" data-action="detail" data-id="${h(c.id)}">詳細表示</button><button type="button" data-action="reload" data-id="${h(c.id)}">フォームに再読込</button><button type="button" data-action="reflect" data-id="${h(c.id)}">見積へ反映</button><button type="button" data-action="delete" data-id="${h(c.id)}" class="danger-btn">削除</button></div></td>
      </tr>`).join('')}</tbody></table>`;
    };

    root.querySelector('#calc-run').addEventListener('click',run); run();
    root.querySelector('#calc-save').addEventListener('click', async()=>{ const sb=app?.getSupabaseClient?.(); const u=app?.getCurrentUser?.(); if(!sb||!u)return;
      run(); const r=JSON.parse(form.dataset.result||'{}'); const f=form.elements;
      const payload={user_id:u.id,client_id:f.clientId.value||null,project_name:f.projectName.value||null,work_type:f.workType.value,application_type:f.applicationType.value,corporate_type:f.corporateType.value,governor_type:f.governorType.value,general_specific:f.generalSpecific.value,industry_count:n(f.industryCount.value),officer_count:n(f.officerCount.value),office_count:n(f.officeCount.value),document_level:f.documentLevel.value,urgent:n(f.urgent.value)>0,keikan_level:f.keikan.value,sengi_level:f.sengi.value,zaisan_level:f.zaisan.value,visit_required:n(f.visit.value)>0,agent_required:n(f.agent.value)>0,expense_amount:r.expense||0,discount_amount:r.discount||0,memo:f.memo.value||null,base_fee:r.base||0,addon_fee:r.addon||0,taxable_subtotal:r.taxable||0,tax:r.tax||0,total:r.total||0,addon_breakdown:JSON.stringify(r.addons||[])};
      const res=await sb.from('estimate_calculations').insert(payload); if(res.error){app.showMessage('見積自動算出の保存に失敗しました。'+res.error.message,true);return;}
      await app.reloadAllData(); renderSavedList(); app.showMessage('見積自動算出を保存しました。');
    });
    root.querySelector('#calc-apply').addEventListener('click', async()=>{run(); const r=JSON.parse(form.dataset.result||'{}'); const f=form.elements;
      await reflectToEstimate({client_id:f.clientId.value||null,project_name:f.projectName.value||null,memo:f.memo.value||null,base_fee:r.base||0,addon_fee:r.addon||0,discount_amount:r.discount||0,expense_amount:r.expense||0,taxable_subtotal:r.taxable||0,tax:r.tax||0,total:r.total||0});
    });

    savedList.addEventListener('click', async(event)=>{
      const btn=event.target.closest('button[data-action]'); if(!btn) return;
      const id=btn.dataset.id; const action=btn.dataset.action;
      const calc=(app?.getEstimateCalculations?.()||[]).find((x)=>String(x.id)===String(id)); if(!calc) return;
      if(action==='detail'){
        const addon=(()=>{try{return JSON.stringify(JSON.parse(calc.addon_breakdown||'[]'));}catch{return calc.addon_breakdown||'[]';}})();
        alert(`作成日: ${fmtDate(calc.created_at)}\n顧客: ${calc.client_name||'-'}\n案件: ${calc.project_name||'-'}\n業務種別: ${calc.work_type||'-'}\n申請区分: ${calc.application_type||'-'}\n基本報酬: ${yen(calc.base_fee||0)}\n加算: ${yen(calc.addon_fee||0)}\n実費: ${yen(calc.expense_amount||0)}\n消費税: ${yen(calc.tax||0)}\n合計: ${yen(calc.total||0)}\n加算明細: ${addon}\nメモ: ${calc.memo||'-'}`);
      } else if(action==='reload'){
        fillForm(calc); app.showMessage('保存済みデータをフォームへ再読込しました。');
      } else if(action==='reflect'){
        await reflectToEstimate(calc);
      } else if(action==='delete'){
        const ok=confirm('この保存済み算出データを削除しますか？'); if(!ok) return;
        const sb=app?.getSupabaseClient?.(); if(!sb) return;
        const del=await sb.from('estimate_calculations').delete().eq('id', calc.id);
        if(del.error){app.showMessage('保存済み算出データの削除に失敗しました。'+del.error.message,true); return;}
        await app.reloadAllData(); renderSavedList(); app.showMessage('保存済み算出データを削除しました。');
      }
    });

    renderSavedList();
  }
  document.addEventListener('DOMContentLoaded', init);
})();
