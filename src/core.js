/* SharePoint Maker Toolkit: deterministic, browser-safe generation logic. */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.SharePointCore = api;
})(typeof window !== 'undefined' ? window : null, function () {
  const TYPES = ['text', 'number', 'boolean', 'dateTime', 'choice'];
  const SCHEMA = 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json';
  const colours = { green: '#0f766e', blue: '#2563eb', violet: '#7c3aed', amber: '#b45309', rose: '#be123c' };

  function internalName(value) {
    const clean = String(value || '').trim().replace(/[^a-zA-Z0-9_ ]/g, '').replace(/\s+/g, '');
    const result = /^[a-zA-Z]/.test(clean) ? clean : 'Field' + clean;
    return result.slice(0, 32) || 'Field';
  }

  function validateProject(project) {
    if (!project || typeof project !== 'object') throw Error('Invalid project format.');
    const name = String(project.name || '').trim();
    if (!name || name.length > 120) throw Error('Enter a list name (up to 120 characters).');
    if (!Array.isArray(project.columns) || project.columns.length > 100) throw Error('Column limit exceeded.');
    const seen = new Set(['title']);
    const columns = project.columns.map((item) => {
      if (!item || !TYPES.includes(item.type)) throw Error('Unsupported column type.');
      const displayName = String(item.displayName || '').trim();
      if (!displayName || displayName.length > 120) throw Error('Column names must be 1–120 characters.');
      const name = internalName(item.name || displayName);
      if (seen.has(name.toLowerCase())) throw Error('Duplicate internal column name: ' + name);
      seen.add(name.toLowerCase());
      const column = { name, displayName, type: item.type, required: !!item.required, indexed: !!item.indexed };
      if (item.type === 'choice') {
        column.choices = String(item.choices || '').split(',').map(s => s.trim()).filter(Boolean).slice(0, 100);
        if (!column.choices.length) throw Error('Choice columns need at least one option.');
      }
      return column;
    });
    return { name, description: String(project.description || '').slice(0, 1000), columns };
  }

  function graphPayload(project) {
    const p = validateProject(project);
    return {
      displayName: p.name,
      description: p.description,
      list: { template: 'genericList' },
      columns: p.columns.map(c => {
        const base = { name: c.name, displayName: c.displayName, required: c.required, indexed: c.indexed };
        if (c.type === 'choice') base.choice = { choices: c.choices, allowTextEntry: false, displayAs: 'dropDownMenu' };
        else if (c.type === 'dateTime') base.dateTime = { format: 'dateOnly' };
        else base[c.type] = {};
        return base;
      })
    };
  }

  function csvSchema(project) {
    const p = validateProject(project);
    const cells = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    return [
      ['Display name', 'Internal name', 'Type', 'Required', 'Indexed', 'Choices'].map(cells).join(','),
      ...p.columns.map(c => [c.displayName, c.name, c.type, c.required, c.indexed, (c.choices || []).join(' | ')].map(cells).join(','))
    ].join('\r\n');
  }

  function listWarnings(project) {
    const p = validateProject(project);
    const messages = [];
    if (!p.columns.some(c => c.indexed)) messages.push('No custom columns are indexed. Plan filtered views before this list becomes large.');
    if (p.columns.some(c => c.type === 'text' && c.displayName.toLowerCase().includes('status'))) messages.push('Consider Choice instead of Text for controlled status values.');
    if (p.columns.length > 18) messages.push('This list has many columns. Check usability and whether the data should be split into related lists.');
    return messages;
  }

  function formattingPayload(preset, colourName) {
    const accent = colours[colourName] || colours.green;
    if (preset === 'status') return {
      $schema: SCHEMA, elmType: 'span',
      style: {
        display: 'inline-flex', 'align-items': 'center', padding: '5px 11px',
        'border-radius': '999px', 'font-weight': '600', 'font-size': '12px',
        color: "=if(@currentField == 'Done', '#065f46', if(@currentField == 'Blocked', '#9f1239', '#1d4ed8'))",
        'background-color': "=if(@currentField == 'Done', '#d1fae5', if(@currentField == 'Blocked', '#ffe4e6', '#dbeafe'))"
      }, txtContent: '@currentField'
    };
    if (preset === 'priority') return {
      $schema: SCHEMA, elmType: 'div',
      style: { display: 'flex', 'align-items': 'center', gap: '7px', 'font-weight': '600' },
      children: [
        { elmType: 'span', style: { height: '9px', width: '9px', 'border-radius': '50%', 'background-color': "=if(@currentField == 'High', '#e11d48', if(@currentField == 'Medium', '#f59e0b', '#059669'))" } },
        { elmType: 'span', txtContent: '@currentField' }
      ]
    };
    if (preset === 'progress') return {
      $schema: SCHEMA, elmType: 'div', style: { display: 'flex', 'align-items': 'center', gap: '8px', width: '100%' },
      children: [
        { elmType: 'div', style: { height: '10px', width: '110px', 'border-radius': '999px', 'background-color': '#e2e8f0', overflow: 'hidden' }, children: [
          { elmType: 'div', style: { height: '100%', 'background-color': accent, width: "=if(@currentField < 0, '0%', if(@currentField > 100, '100%', @currentField + '%'))", 'border-radius': '999px' } }
        ] },
        { elmType: 'span', style: { 'font-weight': '600' }, txtContent: "= @currentField + '%'" }
      ]
    };
    if (preset === 'link') return {
      $schema: SCHEMA, elmType: 'a', attributes: { href: '@currentField', target: '_blank', rel: 'noopener noreferrer' },
      style: { color: accent, 'font-weight': '600', 'text-decoration': 'underline' }, txtContent: 'Open link'
    };
    throw Error('Unknown formatting preset.');
  }

  return { TYPES, colours, internalName, validateProject, graphPayload, csvSchema, listWarnings, formattingPayload };
});
