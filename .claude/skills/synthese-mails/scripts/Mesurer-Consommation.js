// Consommation reelle des lancements de Panoramix, d'apres les journaux de Claude Code. Usage : node Mesurer-Consommation.js [AAAA-MM-JJ]
const fs = require('fs'), path = require('path');
const root = process.env.USERPROFILE + '/.claude/projects';
const since = new Date(process.argv[2] || Date.now() - 7 * 864e5);
const kinds = [
  [/^Nous sommes le /, 'chat'], [/^Travail Panoramix/, 'preparation'],
  [/^Avec le connecteur Atlassian, en lecture seule\. Obtiens le cloudId/, 'commentaires Jira'],
  [/^Avec le connecteur Atlassian/, 'liste Jira'],
  [/synthèse incrémentale|« Synthèse incrémentale »/i, 'synthese incrementale'],
  [/mode « Actualisation »/, 'actualisation'],
  [/^Utilise le skill synthese-mails pour produire la synthèse matinale/, 'synthese complete'],
  [/squad/i, 'squad'],
  [/brouillon|prise de contact|Rédige/i, 'redaction contact'],
];
const rows = [];
for (const dir of fs.readdirSync(root)) {
  if (dir.startsWith('ssh-')) continue;
  const full = path.join(root, dir);
  if (!fs.statSync(full).isDirectory()) continue;
  for (const f of fs.readdirSync(full).filter(x => x.endsWith('.jsonl'))) {
    const p = path.join(full, f);
    if (fs.statSync(p).mtime < since) continue;
    const lines = fs.readFileSync(p, 'utf8').split('\n');
    let first = null, start = null, ep = null;
    const seen = new Map();
    const tools = {}; const reads = [];
    for (const l of lines) {
      if (!l.trim()) continue;
      let o; try { o = JSON.parse(l); } catch { continue; }
      if (ep === null && o.entrypoint) ep = o.entrypoint;
      if (!first && o.type === 'user' && o.message) {
        const c = o.message.content;
        first = typeof c === 'string' ? c : (Array.isArray(c) ? (c.find(x => x.type === 'text') || {}).text : '') || '';
        start = o.timestamp;
      }
      if (o.type === 'assistant' && o.message && o.message.usage) {
        seen.set(o.message.id, { u: o.message.usage, m: o.message.model });
        for (const c of (o.message.content || [])) if (c.type === 'tool_use') {
          tools[c.name] = (tools[c.name] || 0) + 1;
          if (c.name === 'Read' && c.input && c.input.file_path) reads.push(path.basename(c.input.file_path));
        }
      }
    }
    if (!first) continue;
    const k = kinds.find(([r]) => r.test(first));
    if (!k) continue;
    const t = { kind: k[1], start, tours: seen.size, inp: 0, cw: 0, cr: 0, out: 0, models: {}, tools, reads };
    for (const { u, m } of seen.values()) {
      if (m === '<synthetic>') continue;
      t.inp += u.input_tokens || 0; t.cw += u.cache_creation_input_tokens || 0; t.cr += u.cache_read_input_tokens || 0; t.out += u.output_tokens || 0;
      t.models[m] = (t.models[m] || 0) + 1;
    }
    rows.push(t);
  }
}
rows.sort((a, b) => String(a.start).localeCompare(String(b.start)));
const k = n => (n / 1000).toFixed(0).padStart(6) + 'k';
console.log('date        heure  type                      tours  entree  cache+   cacheLu  sortie  modele');
for (const r of rows) console.log(String(r.start).slice(0, 10), String(r.start).slice(11, 16), r.kind.padEnd(24), String(r.tours).padStart(5), k(r.inp), k(r.cw), k(r.cr), k(r.out), Object.keys(r.models).map(m => m.replace('claude-', '')).join(','));
console.log('\nPar type (moyenne par lancement)');
const by = {};
for (const r of rows) { const b = by[r.kind] = by[r.kind] || { n: 0, tours: 0, inp: 0, cw: 0, cr: 0, out: 0 }; b.n++; for (const f of ['tours', 'inp', 'cw', 'cr', 'out']) b[f] += r[f]; }
for (const [n, b] of Object.entries(by)) console.log(n.padEnd(24), 'n=' + String(b.n).padStart(3), 'tours', (b.tours / b.n).toFixed(0).padStart(4), 'cache+', k(b.cw / b.n), 'cacheLu', k(b.cr / b.n), 'sortie', k(b.out / b.n));
const last = rows.filter(r => ['actualisation', 'synthese incrementale', 'synthese complete'].includes(r.kind)).slice(-3);
for (const r of last) { console.log('\n', r.start, r.kind, JSON.stringify(r.tools)); const c = {}; r.reads.forEach(x => c[x] = (c[x] || 0) + 1); console.log(JSON.stringify(c)); }

// Par jour et par poste, en entrees equivalentes Opus (lecture cache 0,05x ; ecriture 1 h 2x ; sortie 5x)
const eq = r => r.inp + r.cw * 2 + r.cr * 0.05 + r.out * 5;
const days = {};
for (const r of rows) { const d = String(r.start).slice(0, 10); (days[d] = days[d] || {})[r.kind] = (days[d][r.kind] || 0) + eq(r) ; }
console.log('\nEquivalent entree Opus (milliers), par jour');
for (const [d, v] of Object.entries(days)) console.log(d, Object.entries(v).map(([k, x]) => k + ' ' + Math.round(x / 1000)).join(' | '), '=> total', Math.round(Object.values(v).reduce((a, b) => a + b, 0) / 1000));
const tot = {}; rows.forEach(r => tot[r.kind] = (tot[r.kind] || 0) + eq(r)); const all = Object.values(tot).reduce((a, b) => a + b, 0);
console.log('\nParts :', Object.entries(tot).sort((a, b) => b[1] - a[1]).map(([k, x]) => k + ' ' + Math.round(x / all * 100) + '%').join(' | '));
// Equivalent tarif API par lancement, au tarif de chaque modele (entree, cache ecrit 1 h, cache lu, sortie, $/M).
// Haiku : prix sous 100 k d'entree, lecture de cache supposee a 0,1 x.
const tarifs = { opus: [4, 8, 0.2, 20], sonnet: [2, 4, 0.2, 10], haiku: [0.1, 0.2, 0.01, 0.5] };
const cout = r => {
  const m = Object.keys(r.models)[0] || '';
  const t = tarifs[Object.keys(tarifs).find(k => m.includes(k))] || tarifs.opus;
  return (r.inp * t[0] + r.cw * t[1] + r.cr * t[2] + r.out * t[3]) / 1e6;
};
console.log('\nCout moyen par lancement (5 derniers), equivalent API');
for (const kind of Object.keys(by)) {
  const l = rows.filter(r => r.kind === kind).slice(-5);
  console.log(kind.padEnd(24), '$' + (l.reduce((s, r) => s + cout(r), 0) / l.length).toFixed(2), '(' + Object.keys(l[l.length - 1].models).join(',') + ')');
}
