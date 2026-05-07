(function() {
  'use strict';
  if (typeof XLSX === 'undefined' || typeof Chart === 'undefined') {
    document.body.innerHTML = '<div style="color:white;text-align:center;padding:60px;font-family:sans-serif;">Ошибка: библиотеки не загружены.</div>';
    return;
  }

  const state = {
    rawData:[], columns:[],
    segmentField:'', employeeField:'', phoneField:'',
    selectedSegments: new Set(),
    candidateRows:[], finalRows:[],
    employeeCap:100, employees:[],
    loaded:false,
    filtersApplied: false, // флаг: применялись ли фильтры
    sortCriteria: [],
    virtualRowsLimit: 80
  };

  const EMP_COLORS   = ['#f37021','#d95c22','#8d5a4a','#f3a86b','#5a3d36','#ff8a4a'];
  const CHART_COLORS = ['#f37021','#ff8a4a','#f3a86b','#d95c22','#8d5a4a','#5a3d36','#f58d4d','#c44d1a'];

  const fmt       = n  => n != null ? new Intl.NumberFormat('ru-RU').format(n) : '—';
  const esc       = v  => String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const normPhone = v  => String(v??'').replace(/\D/g,'');
  const q         = s  => document.querySelector(s);
  const qa        = s  => document.querySelectorAll(s);
  const getPanel  = n  => document.getElementById('panel-'+n);

  /* ===== TOAST ===== */
  function toast(msg, type='success') {
    const icons = {success:'✓', warning:'⚠', error:'✕'};
    const el = document.createElement('div');
    el.className = 'toast '+type;
    el.innerHTML = '<span style="font-size:14px;">'+icons[type]+'</span><span class="toast-msg">'+esc(msg)+'</span>';
    q('#toastWrap').appendChild(el);
    setTimeout(()=>{ el.style.opacity='0'; setTimeout(()=>el.remove(),200); }, 3000);
  }

  const showLoading = () => q('#loadingOverlay').classList.add('active');
  const hideLoading = () => q('#loadingOverlay').classList.remove('active');

  function findField(cands) {
    const norm = s => String(s).toLowerCase().replace(/[\s_.,-]/g,'');
    for (const c of cands) {
      const found = state.columns.find(col => norm(col)===norm(c));
      if (found) return found;
    }
    return '';
  }

  function updateAll() { updateMetrics(); updatePreview(); updateCharts(); }

  /* ===== ПЕРЕСЧЁТ candidateRows по текущим сегментам ===== */
  function recomputeCandidates() {
    // Применяем только сегментный фильтр (без условий из вкладки Фильтры)
    // Если фильтры были применены — учитываем их тоже
    if (state.filtersApplied) {
      // пересчитываем с учётом активных filter-rule
      const filters = collectFilters();
      state.candidateRows = state.rawData.filter(row => {
        if (!state.selectedSegments.has(String(row[state.segmentField]??'').trim()||'—')) return false;
        return filters.every(f => matchFilter(row, f));
      });
    } else {
      // только сегментный фильтр
      if (state.selectedSegments.size) {
        state.candidateRows = state.rawData.filter(row =>
          state.selectedSegments.has(String(row[state.segmentField]??'').trim()||'—')
        );
      } else {
        state.candidateRows = [];
      }
    }
    // Если уже были применены сотрудники — сбрасываем финал
    state.finalRows = [];
    buildEmployeesPanel();
    updateAll();
  }

  function updateMetrics() {
    q('#metricTotal').textContent = fmt(state.rawData.length)||'—';
    const afterSeg = state.segmentField && state.selectedSegments.size
      ? state.rawData.filter(r=>state.selectedSegments.has(String(r[state.segmentField]??'').trim()||'—')).length : 0;
    q('#metricSegments').textContent = fmt(afterSeg||state.rawData.length)||'—';
    q('#metricFilters').textContent = state.candidateRows.length ? fmt(state.candidateRows.length) : '—';
    q('#metricFinal').textContent = state.finalRows.length ? fmt(state.finalRows.length) : '—';
    const activeEmps = state.employees.filter(e=>!e.excluded&&e.quota>0).length;
    q('#metricEmployees').textContent = state.employees.length ? activeEmps+'/'+state.employees.length : '—';
    const phones = new Set(state.finalRows.map(r=>normPhone(r[state.phoneField])).filter(v=>v.length>=10));
    q('#expPhones').textContent = phones.size ? fmt(phones.size) : '—';
    q('#expEmployees').textContent = activeEmps||'—';
    const can = state.finalRows.length>0 || state.candidateRows.length>0;
    q('#btnExport').disabled = !can;
    q('#btnExportTop').disabled = !can;
    q('#btnExportXlsx').disabled = !can;
  }

  function updatePreview() {
    const rows = state.finalRows.length ? state.finalRows : state.candidateRows.length ? state.candidateRows : state.rawData;
    const cols = state.columns.length ? state.columns : (rows[0]?Object.keys(rows[0]):[]);
    q('#previewHead').innerHTML = '<tr>'+cols.map(c=>'<th>'+esc(c)+'</th>').join('')+'</tr>';
    q('#previewCount').textContent = fmt(rows.length)+' строк (показаны первые '+fmt(Math.min(rows.length,state.virtualRowsLimit))+')';
    if (!rows.length) {
      q('#previewBody').innerHTML = '<tr><td colspan="'+(cols.length||1)+'" style="text-align:center;color:var(--text-muted);padding:18px;">Нет данных</td></tr>';
      return;
    }
    q('#previewBody').innerHTML = rows.slice(0,state.virtualRowsLimit).map(r=>'<tr>'+cols.map(c=>'<td>'+esc(String(r[c]??''))+'</td>').join('')+'</tr>').join('');
  }

  /* ===== CHARTS ===== */
  let charts={seg:null,emp:null};

  function makeSegData(rows,field,topN) {
    const map=new Map();
    rows.forEach(r=>{const k=String(r[field]??'').trim()||'—';map.set(k,(map.get(k)||0)+1);});
    const sorted=Array.from(map).sort((a,b)=>b[1]-a[1]);
    const top=sorted.slice(0,topN), other=sorted.slice(topN).reduce((s,p)=>s+p[1],0);
    const labels=top.map(x=>x[0]), values=top.map(x=>x[1]);
    if(other>0){labels.push('Прочее');values.push(other);}
    return{labels,values};
  }
  function makeEmpData(rows,field) {
    const map=new Map();
    rows.forEach(r=>{const k=String(r[field]??'').trim()||'—';map.set(k,(map.get(k)||0)+1);});
    const sorted=Array.from(map).sort((a,b)=>b[1]-a[1]);
    return{labels:sorted.map(x=>x[0]),values:sorted.map(x=>x[1])};
  }

  function updateCharts() {
    const base = state.candidateRows.length ? state.candidateRows : state.rawData;
    const sData = makeSegData(base, state.segmentField, 6);
    if(charts.seg) charts.seg.destroy();
    charts.seg = new Chart(q('#segmentChart').getContext('2d'),{
      type:'doughnut',
      data:{labels:sData.labels,datasets:[{data:sData.values,backgroundColor:CHART_COLORS,borderWidth:0,borderRadius:3}]},
      options:{animation:false,responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{color:'rgba(255,255,255,0.45)',font:{size:10},boxWidth:9,padding:9}}}}
    });
    const eBase = state.finalRows.length ? state.finalRows : (state.candidateRows.length ? state.candidateRows : state.rawData);
    const eData = makeEmpData(eBase, state.employeeField);
    if(!eData.labels.length){eData.labels=['Нет данных'];eData.values=[0];}
    const ew = q('#empChartWrap');
    ew.style.height = Math.max(150,eData.labels.length*27+38)+'px';
    if(charts.emp) charts.emp.destroy();
    charts.emp = new Chart(q('#employeeChart').getContext('2d'),{
      type:'bar',
      data:{labels:eData.labels,datasets:[{data:eData.values,backgroundColor:eData.labels.map((_,i)=>EMP_COLORS[i%EMP_COLORS.length]),borderRadius:4,borderSkipped:false}]},
      options:{animation:false,parsing:false,indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'rgba(255,255,255,0.35)',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'}},y:{ticks:{color:'rgba(255,255,255,0.65)',font:{size:10}},grid:{display:false}}}}
    });
  }

  /* ===== SEGMENTS ===== */
  function buildSegmentsPanel() {
    const field = state.segmentField; if(!field) return;
    const map = new Map();
    state.rawData.forEach(r=>{const v=String(r[field]??'').trim()||'—';map.set(v,(map.get(v)||0)+1);});
    const groups = new Map();
    map.forEach((count,val)=>{const fw=val.split(/[\s_-]/)[0]||'—';if(!groups.has(fw))groups.set(fw,[]);groups.get(fw).push({val,count});});
    const sortedGroups = Array.from(groups).sort((a,b)=>a[0].localeCompare(b[0],'ru'));
    state.selectedSegments = new Set(map.keys());

    let gh = '';
    sortedGroups.forEach(([gname,items])=>{
      const total = items.reduce((s,i)=>s+i.count,0);
      gh += '<div class="seg-group"><div class="seg-group-hdr"><span class="seg-group-name">'+esc(gname)+'</span><span class="seg-group-badge">'+fmt(total)+'</span></div><div class="seg-list">';
      items.sort((a,b)=>a.val.localeCompare(b.val,'ru')).forEach(({val,count})=>{
        gh += '<div class="seg-item active" data-seg="'+esc(val)+'">'
          +'<div class="seg-check"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg></div>'
          +'<span class="seg-name">'+esc(val)+'</span><span class="seg-count">'+count+'</span></div>';
      });
      gh += '</div></div>';
    });

    getPanel('segments').innerHTML =
      '<div class="bcard"><div class="bcard-hdr">'
      +'<div class="flex items-center gap-2"><div class="step-badge">1</div><span class="font-bold" style="font-size:14px;">Сегменты</span></div>'
      +'<div class="flex gap-2"><button class="btn btn-ghost btn-sm" id="btnSelAll">Все</button><button class="btn btn-ghost btn-sm" id="btnClrAll">Снять</button></div>'
      +'</div><div class="bcard-body">'
      +'<div class="seg-search-wrap"><span class="search-icon">🔍</span><input class="input" id="segSearch" placeholder="Поиск сегментов..."></div>'
      +'<div id="segGroups">'+gh+'</div>'
      +'</div></div>';

    qa('#panel-segments .seg-item').forEach(item=>{
      item.addEventListener('click',()=>{
        item.classList.toggle('active');
        const seg = item.dataset.seg;
        if(state.selectedSegments.has(seg)) state.selectedSegments.delete(seg);
        else state.selectedSegments.add(seg);
        recomputeCandidates(); // ← МГНОВЕННЫЙ ПЕРЕСЧЁТ
      });
    });

    const si = document.getElementById('segSearch');
    if(si) si.addEventListener('input',e=>{
      const t=e.target.value.toLowerCase();
      qa('#panel-segments .seg-item').forEach(i=>{i.style.display=i.dataset.seg.toLowerCase().includes(t)?'':'none';});
    });

    document.getElementById('btnSelAll').addEventListener('click',()=>{
      state.selectedSegments = new Set(state.rawData.map(r=>String(r[state.segmentField]??'').trim()||'—'));
      qa('#panel-segments .seg-item').forEach(i=>i.classList.add('active'));
      recomputeCandidates();
    });
    document.getElementById('btnClrAll').addEventListener('click',()=>{
      state.selectedSegments = new Set();
      qa('#panel-segments .seg-item').forEach(i=>i.classList.remove('active'));
      recomputeCandidates();
    });
  }

  /* ===== FILTERS ===== */
  const FCFG = {
    'Телефон':                    {op:'value',  ph:'7999...'},
    'Общ.сумма':                  {op:'number', ph:'Сумма'},
    'Посл.покупка':               {op:'date',   ph:'дд.мм.гггг'},
    'Сумма покупок за период':    {op:'number', ph:'Сумма'},
    'Из них сумма Прямых продаж': {op:'number', ph:'Сумма'},
    'Кол-во штук товаров':        {op:'number', ph:'Кол-во'},
    'Город':                      {op:'value',  ph:'Москва'},
    'Дата коммуникации':          {op:'date',   ph:'дд.мм.гггг'},
    'ДР':                         {op:'date',   ph:'дд.мм.гггг'}
  };
  const EXCL = ['Номер карты','ФИО','Заметки'];
  let fId = 0;

  function collectFilters() {
    const filters = [];
    qa('#filterRules .filter-rule').forEach(r=>{
      const field = r.querySelector('.ff-field')?.value;
      if(!field) return;
      const op = r.querySelector('.ff-op')?.value;
      if(op==='date_range'){
        const d1=(r.querySelector('.ff-d1')||{}).value||'';
        const d2=(r.querySelector('.ff-d2')||{}).value||'';
        if(d1||d2) filters.push({field,op,value:d1,value2:d2});
      } else {
        const val=(r.querySelector('.ff-val')||{}).value||'';
        if(val) filters.push({field,op,value:val});
      }
    });
    return filters;
  }

  function buildFiltersPanel() {
    const cols = state.columns.filter(c=>!EXCL.includes(c)&&c!==state.segmentField&&c!==state.employeeField);
    getPanel('filters').innerHTML =
      '<div class="bcard"><div class="bcard-hdr">'
      +'<div class="flex items-center gap-2"><div class="step-badge">2</div><span class="font-bold" style="font-size:14px;">Фильтры</span></div>'
      +'</div><div class="bcard-body">'
      +'<div style="margin-bottom:11px;"><button class="btn btn-ghost btn-sm" id="btnAddF">+ Добавить условие</button></div>'
      +'<div id="filterRules"></div>'
      +'<div class="apply-bar"><span class="text-sm text-muted">Фильтры необязательны</span><button class="btn btn-primary btn-sm" id="btnApplyF">Применить</button></div>'
      +'</div></div>';

    document.getElementById('btnAddF').addEventListener('click',()=>{
      fId++;
      const rule = document.createElement('div');
      rule.className = 'filter-rule';
      rule.innerHTML =
        '<div class="filter-rule-hdr"><span class="font-semibold text-sm" style="color:var(--text-secondary);">Условие '+fId+'</span>'
        +'<button class="btn btn-ghost btn-sm" data-rm style="padding:4px 8px;">✕</button></div>'
        +'<div class="filter-fields">'
        +'<div><label class="field-label">Поле</label><select class="select ff-field">'+cols.map(c=>'<option value="'+esc(c)+'">'+esc(c)+'</option>').join('')+'</select></div>'
        +'<div><label class="field-label">Условие</label><select class="select ff-op"></select></div>'
        +'<div class="ff-val-wrap"><label class="field-label">Значение</label><input class="input ff-val" placeholder="..."></div>'
        +'</div>';
      document.getElementById('filterRules').appendChild(rule);

      const fs=rule.querySelector('.ff-field'), ops=rule.querySelector('.ff-op'), vw=rule.querySelector('.ff-val-wrap');

      function updUI(){
        const cfg=FCFG[fs.value]||{op:'value',ph:'Значение'};
        ops.innerHTML='';
        if(cfg.op==='value'){
          ops.innerHTML='<option value="value">Содержит</option>';
          vw.innerHTML='<label class="field-label">Значение</label><input class="input ff-val" placeholder="'+esc(cfg.ph)+'">';
        } else if(cfg.op==='number'){
          ops.innerHTML='<option value="gt">Больше</option><option value="gte">≥</option><option value="lt">Меньше</option><option value="lte">≤</option><option value="eq">Равно</option>';
          vw.innerHTML='<label class="field-label">Значение</label><input type="number" class="input ff-val" placeholder="'+esc(cfg.ph)+'">';
        } else if(cfg.op==='date'){
          ops.innerHTML='<option value="date_eq">В дату</option><option value="date_after">После</option><option value="date_before">До</option><option value="date_range">Диапазон</option>';
          setDateInput(vw,'date_eq');
        }
      }
      function setDateInput(wrap,op){
        if(op==='date_range') wrap.innerHTML='<label class="field-label">Диапазон</label><div class="date-range"><input class="input ff-d1" placeholder="дд.мм.гггг"><span>—</span><input class="input ff-d2" placeholder="дд.мм.гггг"></div>';
        else wrap.innerHTML='<label class="field-label">Дата</label><input class="input ff-val" placeholder="дд.мм.гггг">';
      }
      fs.addEventListener('change',updUI);
      ops.addEventListener('change',()=>{ if((FCFG[fs.value]||{}).op==='date') setDateInput(vw,ops.value); });
      updUI();
      rule.querySelector('[data-rm]').addEventListener('click',()=>rule.remove());
    });

    document.getElementById('btnApplyF').addEventListener('click',()=>{
      if(!state.selectedSegments.size) return toast('Выберите сегменты','warning');
      showLoading();
      setTimeout(async ()=>{
        const filters = collectFilters();
        state.candidateRows = await filterLargeData(state.rawData, row=>{
          if(!state.selectedSegments.has(String(row[state.segmentField]??'').trim()||'—')) return false;
          return filters.every(f=>matchFilter(row,f));
        });
        state.filtersApplied = true;
        buildEmployeesPanel();
        hideLoading();
        toast('Отобрано: '+fmt(state.candidateRows.length));
        updateAll();
      },30);
    });
  }


  async function filterLargeData(rows, predicate, chunk=5000){
    const out=[];
    for(let i=0;i<rows.length;i+=chunk){
      const part=rows.slice(i,i+chunk);
      for(const r of part){ if(predicate(r)) out.push(r); }
      await new Promise(requestAnimationFrame);
    }
    return out;
  }

  function parseDate(str){
    if(!str) return null;
    const p=String(str).split('.');
    if(p.length===3) return new Date(+p[2],+p[1]-1,+p[0]);
    return new Date(str);
  }
  function matchFilter(row,f){
    const raw=String(row[f.field]??'');
    const cfg=FCFG[f.field];
    if(cfg&&cfg.op==='date'){
      const rd=parseDate(raw);if(!rd)return false;
      if(f.op==='date_eq'){const d=parseDate(f.value);return d&&rd.toDateString()===d.toDateString();}
      if(f.op==='date_after'){const d=parseDate(f.value);return d&&rd>=d;}
      if(f.op==='date_before'){const d=parseDate(f.value);return d&&rd<=d;}
      if(f.op==='date_range'){const d1=parseDate(f.value),d2=parseDate(f.value2);if(d1&&d2)return rd>=d1&&rd<=d2;if(d1)return rd>=d1;if(d2)return rd<=d2;return true;}
      return true;
    }
    if(f.op==='value') return raw.toLowerCase().includes(String(f.value).toLowerCase());
    const num=parseFloat(raw.replace(/[\s]/g,'').replace(',','.'));
    const fnum=parseFloat(String(f.value).replace(',','.'));
    if(isNaN(num)||isNaN(fnum)) return false;
    switch(f.op){case'gt':return num>fnum;case'gte':return num>=fnum;case'lt':return num<fnum;case'lte':return num<=fnum;case'eq':return num===fnum;default:return true;}
  }

  /* ===== PRIORITY ===== */
  function buildPriorityPanel(){
    const avail = state.columns.filter(c=>c!=='Номер карты'&&c!=='Заметки');
    getPanel('priority').innerHTML =
      '<div class="bcard"><div class="bcard-hdr">'
      +'<div class="flex items-center gap-2"><div class="step-badge">3</div><span class="font-bold" style="font-size:14px;">Приоритет сортировки</span></div>'
      +'<div class="flex gap-2 items-center">'
      +'<select class="select" id="prioSel" style="width:auto;min-width:130px;font-size:12px;">'+avail.map(c=>'<option value="'+esc(c)+'">'+esc(c)+'</option>').join('')+'</select>'
      +'<button class="btn btn-outline btn-sm" id="btnAddPrio">+ Добавить</button>'
      +'</div></div>'
      +'<div class="bcard-body">'
      +'<div class="text-xs text-muted" style="margin-bottom:9px;">Перетаскивайте для порядка. ▼ убывание, ▲ возрастание.</div>'
      +'<div class="text-xs text-muted" id="prioExplain" style="margin-bottom:8px;">Порядок пока не задан</div>'+'<div class="priority-list" id="prioList"><div class="empty-state" style="padding:18px;"><div class="empty-state-text">Добавьте поля сортировки</div></div></div>'
      +'<div class="apply-bar"><span class="text-sm text-muted">Применяется к отобранным строкам</span><button class="btn btn-primary btn-sm" id="btnApplyPrio">Применить</button></div>'
      +'</div></div>';

    document.getElementById('btnAddPrio').addEventListener('click',()=>{
      const field=document.getElementById('prioSel').value;if(!field)return;
      const used=new Set(Array.from(qa('#prioList .priority-item')).map(el=>el.dataset.field));
      if(used.has(field)) return toast('Поле уже добавлено','warning');
      renderPrio(getPrio().concat([{field,dir:'desc'}]));
    });

    document.getElementById('btnApplyPrio').addEventListener('click',()=>{
      const prio=getPrio();
      if(!prio.length) return toast('Добавьте поля приоритета','warning');
      if(!state.candidateRows.length) return toast('Сначала примените фильтры','warning');
      state.sortCriteria = prio;
      sortByCriteria(prio);
      toast('Приоритет применён: '+prio.map(p=>p.field+' '+dirLbl(p.dir)).join(', '));
      updateAll();
    });
  }

  function getPrio(){ return Array.from(qa('#prioList .priority-item')).map(el=>({field:el.dataset.field,dir:el.dataset.dir})); }

  function sortByCriteria(prio){
    state.candidateRows.sort((a,b)=>{
      for(const p of prio){
        if(p.dir==='custom') continue;
        const va=a[p.field],vb=b[p.field];
        const na=parseFloat(String(va||'').replace(/[\s]/g,'').replace(',','.'));
        const nb=parseFloat(String(vb||'').replace(/[\s]/g,'').replace(',','.'));
        let cmp;
        if(!isNaN(na)&&!isNaN(nb)){cmp=na-nb;}
        else{const da=parseDate(String(va||'')),db=parseDate(String(vb||''));if(da&&db)cmp=da-db;else cmp=String(va||'').localeCompare(String(vb||''),'ru');}
        if(cmp!==0) return p.dir==='desc'?-cmp:cmp;
      }
      return 0;
    });
  }

  function renderPrio(items){
    const list=document.getElementById('prioList');if(!list)return;
    if(!items.length){list.innerHTML='<div class="empty-state" style="padding:18px;"><div class="empty-state-text">Добавьте поля сортировки</div></div>'; const ex=document.getElementById('prioExplain'); if(ex) ex.textContent='Порядок пока не задан'; return;}
    const ex=document.getElementById('prioExplain'); if(ex) ex.textContent='Сортировка: '+items.map((i,idx)=>`${idx+1}) ${i.field} ${dirLbl(i.dir)}`).join(' → ');
    list.innerHTML=items.map(i=>
      '<div class="priority-item" draggable="true" data-field="'+esc(i.field)+'" data-dir="'+i.dir+'">'
      +'<span class="drag-handle">⠿</span><span class="priority-field">'+esc(i.field)+'</span>'
      +'<button class="priority-dir">'+dirLbl(i.dir)+'</button><button class="priority-rm">✕</button></div>'
    ).join('');

    list.querySelectorAll('.priority-dir').forEach(btn=>{
      btn.addEventListener('click',e=>{
        e.stopPropagation();
        const item=btn.closest('.priority-item');
        const dirs=['desc','asc','custom'];
        const next=dirs[(dirs.indexOf(item.dataset.dir)+1)%dirs.length];
        item.dataset.dir=next;btn.textContent=dirLbl(next);
      });
    });
    list.querySelectorAll('.priority-rm').forEach(btn=>{
      btn.addEventListener('click',e=>{
        e.stopPropagation();btn.closest('.priority-item').remove();
        if(!list.querySelector('.priority-item')) list.innerHTML='<div class="empty-state" style="padding:18px;"><div class="empty-state-text">Добавьте поля сортировки</div></div>';
      });
    });

    let dragEl=null;
    list.querySelectorAll('.priority-item').forEach(item=>{
      item.addEventListener('dragstart',()=>{dragEl=item;item.classList.add('dragging');});
      item.addEventListener('dragend',()=>{item.classList.remove('dragging');dragEl=null;});
      item.addEventListener('dragover',e=>{
        e.preventDefault();if(!dragEl||dragEl===item)return;
        const rect=item.getBoundingClientRect();
        if(e.clientY<rect.top+rect.height/2) item.parentNode.insertBefore(dragEl,item);
        else item.parentNode.insertBefore(dragEl,item.nextSibling);
      });
    });
  }

  function dirLbl(d){ return d==='desc'?'▼ Убыв':d==='asc'?'▲ Возр':'★ Своя'; }

  /* ===== EMPLOYEES ===== */
  function buildEmployeesPanel(){
    if(!state.candidateRows.length){
      getPanel('employees').innerHTML='<div class="bcard"><div class="bcard-body"><div class="empty-state"><div class="empty-state-icon">🔍</div><div class="empty-state-text">Сначала примените фильтры</div></div></div></div>';
      return;
    }
    const map=new Map();
    state.candidateRows.forEach(r=>{const n=String(r[state.employeeField]??'').trim()||'—';if(!map.has(n))map.set(n,[]);map.get(n).push(r);});
    state.employees=Array.from(map).map(([name,rows])=>({name,found:rows.length,quota:Math.min(rows.length,state.employeeCap),excluded:false})).sort((a,b)=>a.name.localeCompare(b.name,'ru'));
    renderEmployeesPanel();
    updateAll();
  }

  function renderEmployeesPanel(){
    let html='<div class="bcard"><div class="bcard-hdr"><div class="flex items-center gap-2"><div class="step-badge">4</div><span class="font-bold" style="font-size:14px;">Распределение</span></div></div>'
      +'<div class="bcard-body"><div class="emp-controls">'
      +'<div class="cap-wrap"><span class="text-sm text-muted">Лимит:</span><input type="range" min="1" max="500" value="'+state.employeeCap+'" id="capSlider" class="cap-slider"><span class="cap-val" id="capVal">'+state.employeeCap+'</span></div>'
      +'<div class="flex gap-2"><button class="btn btn-ghost btn-sm" id="btnEq">⚖️ Поровну</button><button class="btn btn-ghost btn-sm" id="btnFill">🎯 Заполнить</button></div>'
      +'</div><div class="emp-list" id="empList">';

    state.employees.forEach((emp,i)=>{
      const pct=Math.min(100,Math.round((emp.quota/emp.found)*100));
      html+='<div class="emp-item" data-idx="'+i+'">'
        +'<div><div class="emp-name">'+esc(emp.name)+'</div><div class="emp-found">'+fmt(emp.found)+' найдено</div><div class="emp-bar-wrap"><div class="emp-bar-fill" style="width:'+pct+'%"></div></div></div>'
        +'<div class="emp-quota-lbl">'+emp.quota+'</div>'
        +'<div class="emp-quota"><input type="number" value="'+Math.min(emp.found,state.employeeCap)+'" min="0" max="'+emp.found+'" '+(emp.excluded?'disabled':'')+' class="eq-inp"></div>'
        +'<button class="emp-exclude '+(emp.excluded?'active':'')+'">✕</button>'
        +'</div>';
    });

    html+='</div><div class="apply-bar" style="margin-top:13px;"><span class="text-sm text-muted">Настройте квоты</span><button class="btn btn-primary btn-sm" id="btnApplyEmps">Применить</button></div></div></div>';
    getPanel('employees').innerHTML=html;

    const slider=document.getElementById('capSlider'),capValEl=document.getElementById('capVal');
    slider.addEventListener('input',()=>{
      state.employeeCap=parseInt(slider.value);capValEl.textContent=state.employeeCap;
      state.employees.forEach(e=>{if(!e.excluded)e.quota=Math.min(e.found,state.employeeCap);});
      refreshEmpUI();updateAll();
    });
    document.getElementById('btnEq').addEventListener('click',()=>{
      const active=state.employees.filter(e=>!e.excluded);if(!active.length)return;
      const per=Math.min(state.employeeCap,Math.floor(state.candidateRows.length/active.length));
      active.forEach(e=>e.quota=Math.min(e.found,per));refreshEmpUI();updateAll();
    });
    document.getElementById('btnFill').addEventListener('click',()=>{
      state.employees.forEach(e=>{if(!e.excluded)e.quota=Math.min(e.found,state.employeeCap);});refreshEmpUI();updateAll();
    });

    function refreshEmpUI(){
      qa('#empList .emp-item').forEach(item=>{
        const idx=parseInt(item.dataset.idx);const emp=state.employees[idx];if(emp.excluded)return;
        const inp=item.querySelector('.eq-inp'),bar=item.querySelector('.emp-bar-fill'),lbl=item.querySelector('.emp-quota-lbl');
        if(inp){inp.value=emp.quota;bar.style.width=Math.min(100,Math.round((emp.quota/emp.found)*100))+'%';lbl.textContent=emp.quota;}
      });
    }

    qa('#empList .emp-item').forEach(item=>{
      const idx=parseInt(item.dataset.idx);
      const inp=item.querySelector('.eq-inp'),ex=item.querySelector('.emp-exclude');
      const bar=item.querySelector('.emp-bar-fill'),lbl=item.querySelector('.emp-quota-lbl');
      if(inp){
        inp.addEventListener('input',()=>{
          let v=parseInt(inp.value)||0;
          v=Math.max(0,Math.min(v,state.employees[idx].found,state.employeeCap));
          inp.value=v;state.employees[idx].quota=v;
          bar.style.width=Math.min(100,Math.round((v/state.employees[idx].found)*100))+'%';
          lbl.textContent=v;updateAll();
        });
      }
      if(ex){
        ex.addEventListener('click',()=>{
          state.employees[idx].excluded=!state.employees[idx].excluded;
          ex.classList.toggle('active');
          if(state.employees[idx].excluded){
            inp.value=0;inp.disabled=true;state.employees[idx].quota=0;bar.style.width='0%';lbl.textContent='0';
          } else {
            inp.disabled=false;const nq=Math.min(state.employees[idx].found,state.employeeCap);
            inp.value=nq;state.employees[idx].quota=nq;
            bar.style.width=Math.min(100,Math.round((nq/state.employees[idx].found)*100))+'%';lbl.textContent=nq;
          }
          updateAll();
        });
      }
    });

    document.getElementById('btnApplyEmps').addEventListener('click',()=>{
      showLoading();
      setTimeout(()=>{
        const groups=new Map();
        state.candidateRows.forEach(r=>{const n=String(r[state.employeeField]??'').trim()||'—';if(!groups.has(n))groups.set(n,[]);groups.get(n).push(r);});
        state.finalRows=[];
        state.employees.forEach(emp=>{
          if(emp.excluded||emp.quota<=0)return;
          state.finalRows=state.finalRows.concat((groups.get(emp.name)||[]).slice(0,emp.quota));
        });
        hideLoading();toast('Итог: '+fmt(state.finalRows.length));updateAll();
      },80);
    });
  }

  /* ===== MODAL ===== */
  function openModal(title,content){
    q('#modalTitle').textContent=title;
    q('#modalBody').innerHTML=content;
    q('#modalOverlay').classList.add('active');
  }
  q('#modalClose').addEventListener('click',()=>q('#modalOverlay').classList.remove('active'));
  q('#modalOverlay').addEventListener('click',e=>{if(e.target===q('#modalOverlay'))q('#modalOverlay').classList.remove('active');});

  q('#cardSegChart').addEventListener('click',()=>{
    openModal('📊 Сегменты','<div style="height:420px;"><canvas id="mSC"></canvas></div>');
    setTimeout(()=>{
      const ctx=document.getElementById('mSC')?.getContext('2d');if(!ctx)return;
      const d=makeSegData(state.candidateRows.length?state.candidateRows:state.rawData,state.segmentField,10);
      new Chart(ctx,{type:'doughnut',data:{labels:d.labels,datasets:[{data:d.values,backgroundColor:CHART_COLORS,borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'right',labels:{color:'rgba(255,255,255,0.6)',font:{size:12},boxWidth:12}}}}});
    },100);
  });
  q('#cardEmpChart').addEventListener('click',()=>{
    openModal('👥 Сотрудники','<div style="overflow-y:auto;max-height:60vh;"><div id="mEW" style="height:280px;"><canvas id="mEC"></canvas></div></div>');
    setTimeout(()=>{
      const ctx=document.getElementById('mEC')?.getContext('2d');const wrap=document.getElementById('mEW');if(!ctx||!wrap)return;
      const d=makeEmpData(state.finalRows.length?state.finalRows:(state.candidateRows.length?state.candidateRows:state.rawData),state.employeeField);
      wrap.style.height=Math.max(280,d.labels.length*30+48)+'px';
      new Chart(ctx,{type:'bar',data:{labels:d.labels,datasets:[{data:d.values,backgroundColor:d.labels.map((_,i)=>EMP_COLORS[i%EMP_COLORS.length]),borderRadius:4}]},options:{animation:false,parsing:false,indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'rgba(255,255,255,0.4)'},grid:{color:'rgba(255,255,255,0.05)'}},y:{ticks:{color:'rgba(255,255,255,0.7)'},grid:{display:false}}}}});
    },100);
  });
  q('#cardPreview').addEventListener('click',()=>{
    const rows=state.finalRows.length?state.finalRows:state.candidateRows.length?state.candidateRows:state.rawData;
    const cols=state.columns.length?state.columns:(rows[0]?Object.keys(rows[0]):[]);
    openModal('👁️ Превью ('+fmt(rows.length)+' строк)',
      '<table class="modal-table"><thead><tr>'+cols.map(c=>'<th>'+esc(c)+'</th>').join('')+'</tr></thead><tbody>'
      +(rows.length?rows.map(r=>'<tr>'+cols.map(c=>'<td>'+esc(String(r[c]??''))+'</td>').join('')+'</tr>').join(''):'<tr><td colspan="'+cols.length+'" style="text-align:center;padding:18px;color:var(--text-muted);">Нет данных</td></tr>')
      +'</tbody></table>'
    );
  });

  /* ===== FILE UPLOAD ===== */
  const uploadZone=document.getElementById('uploadZone');
  const fileInput=document.getElementById('fileInput');

  uploadZone.addEventListener('click',function(e){
    if(e.target===fileInput) return;
    e.stopPropagation();
    setTimeout(()=>fileInput.click(),0);
  });
  uploadZone.addEventListener('dragover',e=>{e.preventDefault();e.stopPropagation();uploadZone.classList.add('dragover');});
  uploadZone.addEventListener('dragleave',e=>{e.stopPropagation();uploadZone.classList.remove('dragover');});
  uploadZone.addEventListener('drop',e=>{
    e.preventDefault();e.stopPropagation();uploadZone.classList.remove('dragover');
    const file=e.dataTransfer.files[0];if(file)handleFile(file);
  });
  fileInput.addEventListener('change',e=>{
    const file=e.target.files[0];if(file)handleFile(file);
    e.target.value='';
  });

  function handleFile(file){
    showLoading();
    const reader=new FileReader();
    reader.onerror=()=>{hideLoading();toast('Ошибка чтения файла','error');};
    reader.onload=function(evt){
      try{
        let wb;
        if(/\.csv$/i.test(file.name)) wb=XLSX.read(evt.target.result,{type:'string',codepage:65001});
        else wb=XLSX.read(new Uint8Array(evt.target.result),{type:'array'});
        const sheet=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(sheet,{defval:''});
        if(!rows.length){hideLoading();return toast('Файл пустой','error');}

        state.rawData=rows;
        state.columns=Object.keys(rows[0]);
        state.segmentField=findField(['Категория','Category','Segment','Сегмент']);
        state.employeeField=findField(['Сотрудник','Employee','Менеджер','Manager']);
        state.phoneField=findField(['Телефон','Phone','Tel','Мобильный']);
        state.filtersApplied=false;

        if(!state.segmentField||!state.employeeField||!state.phoneField){
          hideLoading();
          return toast('Не найдены поля. Доступные: '+state.columns.join(', '),'error');
        }

        state.loaded=true;
        uploadZone.classList.add('loaded');
        q('#uploadText').textContent='✅ '+file.name;
        q('#uploadHint').textContent=fmt(rows.length)+' строк загружено';

        buildSegmentsPanel();
        buildFiltersPanel();
        buildPriorityPanel();
        getPanel('employees').innerHTML='<div class="bcard"><div class="bcard-body"><div class="empty-state"><div class="empty-state-icon">🔍</div><div class="empty-state-text">Сначала примените фильтры</div></div></div></div>';

        qa('.tab').forEach(t=>t.classList.remove('disabled'));

        // Инициализируем candidateRows всеми строками
        state.candidateRows = [...state.rawData];

        hideLoading();
        toast('Загружено: '+fmt(rows.length)+' строк');
        updateAll();
      }catch(err){
        hideLoading();console.error(err);
        toast('Ошибка: '+err.message,'error');
      }
    };
    if(/\.csv$/i.test(file.name)) reader.readAsText(file,'utf-8');
    else reader.readAsArrayBuffer(file);
  }

  /* ===== TABS ===== */
  qa('.tab').forEach(tab=>{
    tab.addEventListener('click',function(){
      if(this.classList.contains('disabled')) return;
      qa('.tab').forEach(t=>t.classList.remove('active'));
      qa('.panel').forEach(p=>p.classList.remove('active'));
      this.classList.add('active');
      const panel=document.getElementById('panel-'+this.dataset.tab);
      if(panel) panel.classList.add('active');
    });
  });

  /* ===== EXPORT ===== */
  qa('#btnExport, #btnExportTop').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const selectedRows = state.finalRows.length ? state.finalRows : state.candidateRows;
      if(!selectedRows.length) return toast('Сначала сделайте отбор по сегментам/фильтрам','warning');
      const phones=[...new Set(selectedRows.map(r=>normPhone(r[state.phoneField])).filter(v=>v.length>=10))];
      const blob=new Blob([phones.join('\n')],{type:'text/plain;charset=utf-8'});
      const a=document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.download='compas_'+new Date().toISOString().slice(0,10)+'.txt';
      document.body.appendChild(a);a.click();document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
      toast('Готово: '+fmt(phones.length)+' номеров');
    });
  });


  q('#btnExportXlsx').addEventListener('click',()=>{
    const selectedRows = state.finalRows.length ? state.finalRows : state.candidateRows;
    if(!selectedRows.length) return toast('Сначала сделайте отбор по сегментам/фильтрам','warning');
    const ws=XLSX.utils.json_to_sheet(selectedRows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Selection');
    XLSX.writeFile(wb,'compas_'+new Date().toISOString().slice(0,10)+'.xlsx');
    toast('XLSX выгружен: '+fmt(selectedRows.length)+' строк');
  });

  /* ===== SAVE / RESET ===== */
  q('#btnSave').addEventListener('click',()=>{
    try{
      localStorage.setItem('compas_cfg',JSON.stringify({selected:Array.from(state.selectedSegments),cap:state.employeeCap,time:Date.now()}));
      toast('Сохранено');
    }catch(e){toast('Ошибка сохранения','error');}
  });

  q('#btnReset').addEventListener('click',()=>{
    if(!confirm('Сбросить все данные?')) return;
    state.rawData=[];state.candidateRows=[];state.finalRows=[];
    state.selectedSegments=new Set();state.employees=[];
    state.loaded=false;state.filtersApplied=false;

    uploadZone.classList.remove('loaded');
    q('#uploadText').textContent='Перетащите файл или нажмите для выбора';
    q('#uploadHint').textContent='.xlsx или .csv';

    getPanel('segments').innerHTML='<div class="bcard"><div class="bcard-body"><div class="empty-state"><div class="empty-state-icon">📁</div><div class="empty-state-text">Загрузите файл для начала работы</div></div></div></div>';
    getPanel('filters').innerHTML='';getPanel('priority').innerHTML='';getPanel('employees').innerHTML='';

    qa('.tab').forEach((t,i)=>{t.classList.remove('active');if(i===0)t.classList.add('active');else t.classList.add('disabled');});
    qa('.panel').forEach(p=>p.classList.remove('active'));
    document.getElementById('panel-segments').classList.add('active');

    ['metricTotal','metricSegments','metricFilters','metricFinal','metricEmployees'].forEach(id=>{const el=document.getElementById(id);if(el)el.textContent='—';});
    q('#expPhones').textContent='—';q('#expEmployees').textContent='—';
    q('#previewBody').innerHTML='<tr><td style="text-align:center;color:var(--text-muted);padding:18px;">Нет данных</td></tr>';
    q('#previewHead').innerHTML='';q('#previewCount').textContent='0 строк';
    qa('#btnExport, #btnExportTop, #btnExportXlsx').forEach(b=>b.disabled=true);

    updateCharts();toast('Сброшено','warning');
  });

  updateCharts();
})();
