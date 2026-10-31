(function () {
  'use strict';
  const core = window.SharePointCore;
  const storageKey = 'mightora.sharepoint-maker.v1';
  const $ = id => document.getElementById(id);
  const example = {
    name: 'Project Tracker', description: 'Track delivery work and progress',
    columns: [
      { name: 'Status', displayName: 'Status', type: 'choice', choices: 'Planned, In progress, Done, Blocked', required: true, indexed: true },
      { name: 'Priority', displayName: 'Priority', type: 'choice', choices: 'Low, Medium, High', required: false, indexed: true },
      { name: 'DueDate', displayName: 'Due Date', type: 'dateTime', required: false, indexed: false },
      { name: 'Complete', displayName: 'Complete', type: 'boolean', required: false, indexed: false }
    ]
  };
  let project = structuredCloneSafe(example);
  let toastTimeout;
  function structuredCloneSafe(value) { return JSON.parse(JSON.stringify(value)); }
  function safeLoad() {
    try {
      const value = window.localStorage.getItem(storageKey);
      if (value) project = core.validateProject(JSON.parse(value));
    } catch (_) { project = structuredCloneSafe(example); }
  }
  function save() {
    try { window.localStorage.setItem(storageKey, JSON.stringify(project)); }
    catch (_) { /* Private browsing or storage disabled; exports still work. */ }
  }
  function toast(message) {
    const element = $('toast'); element.textContent = message; element.classList.add('visible');
    clearTimeout(toastTimeout); toastTimeout = setTimeout(() => element.classList.remove('visible'), 3300);
  }
  function node(tag, text, klass) {
    const el = document.createElement(tag);
    if (text !== undefined && text !== null) el.textContent = text;
    if (klass) el.className = klass;
    return el;
  }
  function typeLabel(type) { return {text:'Text',number:'Number',boolean:'Yes / No',dateTime:'Date',choice:'Choice'}[type] || type; }
  function render() {
    $('list-name').value = project.name;
    $('list-description').value = project.description;
    $('column-count').textContent = project.columns.length;
    const list = $('column-list'); list.replaceChildren();
    project.columns.forEach((column, index) => {
      const card = node('div', undefined, 'column-card');
      card.append(node('span', String(index + 1).padStart(2, '0'), 'column-number'));
      const info = node('div', undefined, 'column-info');
      info.append(node('strong', column.displayName), node('small', column.name + ' · ' + typeLabel(column.type)));
      card.append(info);
      const flags = node('div', undefined, 'column-flags');
      if (column.required) flags.append(node('span', 'Required', 'flag'));
      if (column.indexed) flags.append(node('span', 'Indexed', 'flag'));
      card.append(flags);
      const button = node('button', '×', 'remove-field'); button.type = 'button';
      button.setAttribute('aria-label', 'Remove ' + column.displayName);
      button.addEventListener('click', () => { project.columns.splice(index, 1); save(); render(); toast('Column removed'); });
      card.append(button); list.append(card);
    });
    const body = $('schema-body'); body.replaceChildren();
    [{ displayName:'Title',type:'text',required:false,indexed:false }, ...project.columns].forEach(col => {
      const tr = node('tr'); const first = node('td', col.displayName); const second = node('td'); second.append(node('span', typeLabel(col.type), 'type-chip'));
      const flags = [col.required ? 'Required' : '',col.indexed ? 'Indexed' : ''].filter(Boolean).join(' · ') || '—';
      tr.append(first, second, node('td', flags)); body.append(tr);
    });
    const warningArea = $('warnings'); warningArea.replaceChildren();
    try {
      const warnings = core.listWarnings(project);
      if (warnings.length) warnings.forEach(w => warningArea.append(node('div', w, 'guidance-item')));
      else warningArea.append(node('div', 'Your basic schema is ready to export. Remember to verify permissions and usage patterns.', 'guidance-item good'));
    } catch (error) { warningArea.append(node('div', error.message, 'guidance-item')); }
  }
  function safeProject() {
    const current = { ...project, name:$('list-name').value, description:$('list-description').value };
    return core.validateProject(current);
  }
  function saveFieldInputs() {
    try { project = safeProject(); save(); render(); }
    catch (err) { toast(err.message); }
  }
  function download(filename, data, type='application/json') {
    const url = URL.createObjectURL(new Blob([data], {type:type+';charset=utf-8'}));
    const a = document.createElement('a'); a.href = url; a.download = filename;
    document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function reportDownload(action) {
    try { action(); $('export-status').textContent = 'File created locally. Check your downloads.'; toast('Download prepared'); }
    catch(err) { $('export-status').textContent = err.message; toast(err.message); }
  }
  function renderFormatting() {
    const preset = $('preset').value;
    const accent = $('accent').value;
    $('accent').disabled = preset !== 'progress' && preset !== 'link';
    $('format-code').textContent = JSON.stringify(core.formattingPayload(preset, accent), null, 2);
    const area = $('format-examples'); area.replaceChildren();
    const examples = { status:['Done','In progress','Blocked'], priority:['High','Medium','Low'],progress:['25','65','100'],link:['Project overview','Documentation','Support portal'] }[preset];
    examples.forEach((value,index) => {
      const row = node('div', undefined, 'preview-value'); row.append(node('strong', 'Example ' + (index + 1)));
      const display = node('div');
      if (preset === 'status') {
        const span = node('span', '●  ' + value, 'preview-pill');
        const colours = { Done:['#d1fae5','#065f46'], 'In progress':['#dbeafe','#1d4ed8'], Blocked:['#ffe4e6','#9f1239'] }[value];
        span.style.backgroundColor = colours[0];span.style.color = colours[1];display.append(span);
      } else if (preset === 'priority') {
        const dot = node('i',undefined,'preview-dot');dot.style.background = {High:'#e11d48',Medium:'#f59e0b',Low:'#059669'}[value];display.append(dot,node('span',value));
      } else if(preset === 'progress') {
        const wrap = node('div',undefined,'progress-preview'); const rail = node('div',undefined,'progress-rail');const bar=node('div',undefined,'progress-fill');bar.style.width=value+'%';bar.style.background=core.colours[accent];rail.append(bar);wrap.append(rail,node('span',value+'%'));display.append(wrap);
      } else {
        const a = node('span','Open link ↗');a.style.color=core.colours[accent];a.style.textDecoration='underline';a.style.fontWeight='700';display.append(a);
      }
      row.append(display);area.append(row);
    });
  }
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); toast('JSON copied to clipboard'); }
    catch(_) {
      const t=document.createElement('textarea');t.value=text;t.style.position='fixed';t.style.opacity='0';document.body.append(t);t.select();
      const success=document.execCommand('copy');t.remove();toast(success?'JSON copied to clipboard':'Clipboard unavailable. Select and copy the JSON manually.');
    }
  }
  function updateHeaderHeight() {
    const header = $('brand-header');
    const height = header ? header.getBoundingClientRect().height : 72;
    if (height > 20 && height < 300) {
      document.documentElement.style.setProperty('--header-height',height+'px');
      document.documentElement.style.scrollPaddingTop=(height+55)+'px';
    }
  }
  function setup() {
    safeLoad(); render(); renderFormatting();
    $('list-name').addEventListener('change', saveFieldInputs);
    $('list-description').addEventListener('change', saveFieldInputs);
    $('new-type').addEventListener('change', () => $('choices-wrap').classList.toggle('hidden', $('new-type').value !== 'choice'));
    $('add-column-form').addEventListener('submit', event => {
      event.preventDefault();
      try {
        const displayName=$('new-display').value.trim(); const type=$('new-type').value;
        const column = {displayName,name:core.internalName(displayName),type,required:$('new-required').checked,indexed:$('new-indexed').checked};
        if(type==='choice') column.choices=$('new-choices').value;
        const next={...safeProject(),columns:[...project.columns,column]};
        project=core.validateProject(next);save();render();
        $('new-display').value=''; $('new-choices').value='';$('new-required').checked=false;$('new-indexed').checked=false;
        $('new-display').focus();toast('Column added');
      } catch (err) { toast(err.message); }
    });
    $('preset').addEventListener('change',renderFormatting);$('accent').addEventListener('change',renderFormatting);
    $('copy-format').addEventListener('click',()=>copyText($('format-code').textContent));
    $('download-format').addEventListener('click',()=>download('sharepoint-column-formatting.json',$('format-code').textContent));
    $('download-graph').addEventListener('click',()=>reportDownload(()=>download('sharepoint-create-list.json',JSON.stringify(core.graphPayload(safeProject()),null,2))));
    $('download-csv').addEventListener('click',()=>reportDownload(()=>download('sharepoint-columns.csv',core.csvSchema(safeProject()),'text/csv')));
    $('download-project').addEventListener('click',()=>reportDownload(()=>download('sharepoint-maker-project.json',JSON.stringify({...safeProject(),version:1},null,2))));
    $('import-project').addEventListener('change',async event=>{
      const file=event.target.files[0];if(!file)return;
      try { if(file.size>1024*1024)throw Error('Maximum project file size is 1 MB.');
        const contents=await file.text();project=core.validateProject(JSON.parse(contents));save();render();toast('Project imported'); }
      catch(err){toast('Import failed: '+err.message);}finally{event.target.value='';}
    });
    $('reset-project').addEventListener('click',()=>{if(!window.confirm('Replace the current design with the sample? Download a backup first if needed.'))return;project=structuredCloneSafe(example);save();render();toast('Sample restored');});
    $('show-resources').addEventListener('click',()=>{
      const hosts=new Set(performance.getEntriesByType('resource').map(entry=>{try{return new URL(entry.name).hostname;}catch{return null;}}).filter(Boolean));
      $('resources-output').textContent=hosts.size?'Resources observed from: '+[...hosts].join(', '):'No resource hosts reported. Browser timing entries may be incomplete.';
    });
    updateHeaderHeight();window.addEventListener('resize',updateHeaderHeight);window.addEventListener('load',updateHeaderHeight);
    if(typeof ResizeObserver!=='undefined')new ResizeObserver(updateHeaderHeight).observe($('brand-header'));
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',setup);else setup();
})();
