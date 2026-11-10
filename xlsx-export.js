// Scouting do Banco · criação de ficheiros Excel (.xlsx) com o aspeto da aplicação.
// Usa a biblioteca ExcelJS (vendor/exceljs.min.js, licença MIT), carregada só quando é precisa.
(function () {
  const NAVY = 'FF102F4D', BLUE = 'FF2D86CC', SKY = 'FFE8F6FC', MUTED = 'FF617789', PAPER = 'FFF5FAFC', LINE = 'FFD9E8F1';
  const GREEN = 'FF2F9E62', RED = 'FFD9443B', GKF = 'FFE6F6EC', REDF = 'FFFBE3E1', YEL = 'FFFCEFC0', REDC = 'FFF8D3D0', WHITE = 'FFFFFFFF';
  const ZN = ['Alto esq.', 'Alto centro', 'Alto dir.', 'Baixo esq.', 'Baixo centro', 'Baixo dir.'];
  const font = (o = {}) => Object.assign({ name: 'Arial', size: 10, color: { argb: 'FF16324A' } }, o, o.color ? { color: { argb: o.color } } : {});
  const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
  const center = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const thin = { bottom: { style: 'thin', color: { argb: LINE } } };
  const mmss = s => { s = Math.max(0, Math.round(s || 0)); return String(Math.floor(s / 60)).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0'); };
  const mins = s => Math.round((s || 0) / 60);
  const pct = (a, b) => b ? Math.round(a * 100 / b) + '%' : '–';
  const sign = n => n > 0 ? '+' + n : String(n);
  const dstr = ms => new Date(ms).toLocaleDateString('pt-PT');
  const venue = h => h === true ? 'Casa' : h === false ? 'Fora' : '';
  const sum = a => (a || []).reduce((x, y) => x + (y || 0), 0);
  const sheetName = (used, s) => { s = String(s).replace(/[\[\]:*?\/\\]/g, '').slice(0, 31) || 'Folha'; let n = s, k = 2; while (used.has(n)) n = s.slice(0, 28) + ' ' + k++; used.add(n); return n; };
  // Link interno para outra folha: escrito como 'Folha'!A1 (sem #), que a ExcelJS grava como "location".
  // Link para outra folha com a fórmula HIPERLIGAÇÃO, que o Excel segue sempre (os links "internos" da ExcelJS nem sempre funcionam).
  const link = (ws, addr, text, target) => {
    const c = ws.getCell(addr), q = s => String(s).replace(/"/g, '""');
    c.value = { formula: `HYPERLINK("#'${q(target).replace(/'/g, "''")}'!A1","${q(text)}")`, result: text };
    return c;
  };
  // Números guardados como números (sem o aviso "número guardado como texto"), com formato.
  const PCT = '0%', PM = '+0;-0;0', DEC = '0.0', MMSS = '[mm]:ss';
  const FRAC = (a, b) => b ? a / b : '–';
  const T = s => (s || 0) / 86400;
  const fmtCols = (row, map) => { for (const [col, f] of Object.entries(map)) { const c = row.getCell(+col); if (typeof c.value === 'number') c.numFmt = f; } };

  function band(ws, range, text, size = 18) {
    ws.mergeCells(range); const c = ws.getCell(range.split(':')[0]);
    c.value = text; c.font = font({ bold: true, size, color: WHITE }); c.fill = fill(NAVY); c.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(c.row).height = 34;
  }
  // Quadrado com valor grande e rótulo pequeno (ocupa 2 colunas × 2 linhas).
  function tile(ws, row, col, label, value, bg = SKY, color = NAVY, sub, numFmt) {
    for (const r of [row, row + 1]) { ws.mergeCells(r, col, r, col + 1); for (const c of [col, col + 1]) ws.getCell(r, c).fill = fill(bg); }
    const v = ws.getCell(row, col); v.value = value; v.font = font({ bold: true, size: 20, color }); v.alignment = center;
    if (numFmt && typeof value === 'number') v.numFmt = numFmt;
    const l = ws.getCell(row + 1, col); l.value = sub ? label + '\n' + sub : label; l.font = font({ bold: true, size: 9, color: MUTED }); l.alignment = center;
    ws.getRow(row).height = 38; ws.getRow(row + 1).height = sub ? 28 : 18;
  }
  function subBand(ws, row, text) {
    ws.mergeCells(row, 2, row, 12); const c = ws.getCell(row, 2);
    c.value = text; c.font = font({ bold: true, size: 12, color: WHITE }); c.fill = fill(BLUE); c.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 24;
  }
  function note(ws, addr, text) { const c = ws.getCell(addr); c.value = text; c.font = font({ size: 9, italic: true, color: MUTED }); }
  function title(ws, addr, text) { const c = ws.getCell(addr); c.value = text; c.font = font({ bold: true, size: 11, color: NAVY }); }

  // Mapa da baliza: 3 colunas × 2 alturas, cada zona ocupa 2 colunas × 2 linhas.
  // step = distância entre zonas (3 nas fichas, que têm uma coluna estreita de intervalo).
  function goalMap(ws, row, col, cells, step = 2) {
    cells.forEach((z, i) => {
      const r = row + Math.floor(i / 3) * 2, c = col + (i % 3) * step;
      ws.mergeCells(r, c, r + 1, c + 1);
      const x = ws.getCell(r, c); x.value = (i + 1) + ' · ' + ZN[i] + '\n' + z.text;
      x.fill = fill(z.bg); x.font = font({ bold: true, size: 9, color: z.dark ? WHITE : 'FF102F4D' }); x.alignment = center;
      x.border = { top: { style: 'thin', color: { argb: WHITE } }, left: { style: 'thin', color: { argb: WHITE } }, bottom: { style: 'thin', color: { argb: WHITE } }, right: { style: 'thin', color: { argb: WHITE } } };
      ws.getRow(r).height = 26; ws.getRow(r + 1).height = 26;
    });
    note(ws, ws.getCell(row + 4, col).address, '← poste esquerdo · baliza vista de frente, como quem remata · poste direito →');
    return row + 6;
  }
  const blend = (a) => { const t = [0x2D, 0x86, 0xCC].map(x => Math.round(255 - (255 - x) * a)); return 'FF' + t.map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase(); };
  function shootCells(zs, zg) {
    const mx = Math.max(...zs);
    return zs.map((v, i) => { const a = v ? 0.18 + 0.82 * v / mx : 0.06; return { text: v + ' rem. · ' + (zg[i] || 0) + ' golo' + ((zg[i] || 0) === 1 ? '' : 's'), bg: blend(a), dark: a > 0.55 }; });
  }
  function gkCells(zsv, zga) {
    return zsv.map((s, i) => {
      const n = s + (zga[i] || 0), ga = zga[i] || 0;
      if (!n) return { text: 'sem remates', bg: 'FFEEF2F5', dark: false };
      return { text: ga + ' golo' + (ga === 1 ? '' : 's') + ' de ' + n + ' · ' + Math.round((n - ga) * 100 / n) + '% def.', bg: ga ? (ga / n > 0.4 ? 'FFE07A72' : 'FFF1AFA9') : 'FF8FD0AC', dark: ga / n > 0.4 };
    });
  }
  function readShoot(zs, zg, who) {
    const t = sum(zs); if (!t) return '';
    const L = zs[0] + zs[3], C = zs[1] + zs[4], R = zs[2] + zs[5], H = zs[0] + zs[1] + zs[2], s = Math.max(L, C, R);
    const side = s === R ? 'lado direito' : s === L ? 'lado esquerdo' : 'centro';
    let txt = `${who} mais para o ${side} (${s} de ${t}) e ` + (H * 2 === t ? 'tanto alto como baixo.' : H * 2 > t ? `mais alto (${H} de ${t}).` : `mais baixo (${t - H} de ${t}).`);
    const gt = sum(zg); if (gt) { const b = zg.indexOf(Math.max(...zg)); txt += ` Golos sobretudo no ${ZN[b].toLowerCase()} (${zg[b]} de ${gt}).`; }
    return txt;
  }

  // Folha com o resumo de um jogo e a tabela dos jogadores. nameOf(p) = folha da ficha do jogador (ou nada).
  // Bloco de um jogo a começar na linha o+1 (o = 0 numa folha só deste jogo). Devolve a linha seguinte livre.
  function gameSection(ws, g, ps, nameOf, o, opt = {}) {
    band(ws, `A${o + 1}:T${o + 1}`, (opt.label || '') + `${g.team}  ${g.us}–${g.them}  ${g.opp}`);
    note(ws, 'A' + (o + 2), [venue(g.home), dstr(g.date), 'tempo de jogo ' + mmss(g.dur)].filter(Boolean).join(' · '));
    title(ws, 'A' + (o + 3), 'A equipa neste jogo');
    const tot = k => sum(ps.map(p => p[k]));
    const sOn = tot('son'), sOff = tot('soff'), sv = tot('sv');
    const tiles = [['Golos', g.us, g.opp + ': ' + g.them], ['Remates', g.shUs ?? sOn + sOff, g.shThem != null ? g.opp + ': ' + g.shThem : ''],
      ['À baliza', g.shOnUs ?? sOn, g.shOnThem != null ? g.opp + ': ' + g.shOnThem : ''], ['Faltas', g.fcUs ?? tot('fc'), g.fcThem != null ? g.opp + ': ' + g.fcThem : ''],
      ['Defesas GR', sv, ''], ['% defesas GR', FRAC(sv, g.shOnThem), g.shOnThem != null ? 'de ' + g.shOnThem + ' remates à baliza' : '']];
    tiles.forEach(([l, v, s], i) => {
      const c = 1 + i * 3;
      for (const r of [o + 4, o + 5, o + 6]) { ws.mergeCells(r, c, r, c + 2); ws.getCell(r, c).fill = fill(SKY); }
      const a = ws.getCell(o + 4, c); a.value = v; a.font = font({ bold: true, size: 22, color: NAVY }); a.alignment = center; if (l.startsWith('%') && typeof v === 'number') a.numFmt = PCT;
      const b = ws.getCell(o + 5, c); b.value = l; b.font = font({ bold: true, size: 9, color: MUTED }); b.alignment = center;
      const d = ws.getCell(o + 6, c); d.value = s; d.font = font({ size: 9, color: MUTED }); d.alignment = center;
    });
    ws.getRow(o + 4).height = 40;
    title(ws, 'A' + (o + 8), opt.season ? 'Jogadores · clica no nome para abrir a ficha da época' : 'Jogadores · clica no nome para abrir a ficha do jogador');
    const hdr = ['N.º', 'Jogador', 'GR', 'Titular', 'Minutos', 'Golos', 'GM em campo', 'GS em campo', 'Saldo +/-', 'Remates', 'À baliza', 'Fora', '% à baliza', 'Zonas', 'Defesas GR', 'Faltas c.', 'Faltas s.', 'Amarelos', 'Vermelhos', 'Momentos altos'];
    const H = o + 9; const hr = ws.getRow(H); hr.values = hdr; hr.height = 30;
    hr.eachCell(c => { c.font = font({ bold: true, size: 9, color: NAVY }); c.fill = fill(SKY); c.alignment = center; c.border = thin; });
    ps.forEach((p, k) => {
      const r = H + 1 + k, row = ws.getRow(r), sh = (p.son || 0) + (p.soff || 0), pm = (p.gm || 0) - (p.gs || 0);
      row.values = [p.num, '', p.gk ? 'GR' : '', p.st ? 'Sim' : '', T(p.min), p.g, p.gm, p.gs, pm, sh, p.son || 0, p.soff || 0, FRAC(p.son || 0, sh),
        p.zs && sum(p.zs) ? zoneLabel(p.zs) : '–', p.sv || 0, p.fc, p.fs, p.y || 0, p.r || 0, p.hl];
      fmtCols(row, { 5: MMSS, 9: PM, 13: PCT });
      row.eachCell({ includeEmpty: true }, (c, n) => { if (n > 20) return; c.font = font(); c.alignment = { horizontal: n === 2 ? 'left' : 'center', vertical: 'middle' }; c.border = thin; });
      const target = nameOf(p);
      if (target) link(ws, 'B' + r, p.name, target).font = font({ bold: true, underline: true, color: BLUE });
      else { ws.getCell('B' + r).value = p.name; ws.getCell('B' + r).font = font({ bold: true }); }
      if (p.g) { ws.getCell(r, 6).fill = fill('FFC9EBD7'); ws.getCell(r, 6).font = font({ bold: true, color: GREEN }); }
      ws.getCell(r, 9).font = font({ bold: true, color: pm > 0 ? GREEN : pm < 0 ? RED : MUTED });
      if (p.y) { ws.getCell(r, 18).fill = fill('FFF2C230'); ws.getCell(r, 18).font = font({ bold: true }); }
      if (p.r) { ws.getCell(r, 19).fill = fill(RED); ws.getCell(r, 19).font = font({ bold: true, color: WHITE }); }
    });
    const rt = H + 1 + ps.length, tr = ws.getRow(rt);
    tr.values = ['', 'Total', '', '', T(tot('min')), tot('g'), '', '', '', sOn + sOff, sOn, sOff, FRAC(sOn, sOn + sOff), '', sv, tot('fc'), tot('fs'), tot('y'), tot('r'), tot('hl')];
    tr.eachCell({ includeEmpty: true }, (c, n) => { if (n > 20) return; c.font = font({ bold: true }); c.fill = fill(PAPER); c.alignment = { horizontal: n === 2 ? 'left' : 'center' }; });
    fmtCols(tr, { 5: '[h]:mm:ss', 13: PCT });
    note(ws, 'A' + (rt + 2), 'GM/GS = golos marcados/sofridos com o jogador em campo. Zonas: 1-3 em cima, 4-6 em baixo, da esquerda para a direita (baliza vista de frente).');
    if (opt.back) link(ws, 'P' + (o + 2), opt.back.text, opt.back.target).font = font({ bold: true, underline: true, color: BLUE });
    return rt + 3;
  }
  function fillGameSheet(ws, g, ps, nameOf, back) {
    ws.columns = Array.from({ length: 20 }, (_, i) => ({ width: i === 1 ? 20 : i === 13 ? 19 : 9 }));
    gameSection(ws, g, ps, nameOf, 0, { back, season: !!back });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 9, showGridLines: false }];
  }

  // ---------- livro de um jogo ----------
  function gameBook(ExcelJS, g) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'Scouting do Banco';
    const used = new Set(), ws = wb.addWorksheet(sheetName(used, 'Jogo'), { views: [{ showGridLines: false }] });
    const ps = g.players.slice().sort((a, b) => a.num - b.num);
    const names = new Map(ps.map(p => [p, sheetName(used, p.num + ' ' + p.name)]));
    fillGameSheet(ws, g, ps, p => names.get(p));

    // fichas dos jogadores
    ps.forEach(p => {
      const sh = wb.addWorksheet(names.get(p), { views: [{ showGridLines: false }] });
      sh.columns = [2, 13, 13, 2, 13, 13, 2, 13, 13, 2, 13, 13].map(w => ({ width: w }));
      link(sh, 'B1', '← Voltar ao jogo', ws.name).font = font({ bold: true, underline: true, color: BLUE });
      band(sh, 'B2:L2', `#${p.num} ${p.name}` + (p.gk ? '  ·  GR' : ''));
      note(sh, 'B3', `${p.st ? 'Titular' : 'Suplente'} · ${mmss(p.min)} jogados · ${g.team} ${g.us}–${g.them} ${g.opp} · ${dstr(g.date)}`);
      title(sh, 'B5', 'Neste jogo');
      playerGameBlock(sh, p, g, 5);
    });
    return wb;
  }
  // Dados de um jogador num jogo, a começar na linha r0 (título já escrito antes). Devolve a linha seguinte livre.
  function playerGameBlock(sh, p, g, r0) {
    const shots = (p.son || 0) + (p.soff || 0), pm = (p.gm || 0) - (p.gs || 0), share = g.dur ? (p.min || 0) / g.dur : 0;
    tile(sh, r0 + 1, 2, 'Minutos', T(p.min), SKY, NAVY, '', MMSS); tile(sh, r0 + 1, 5, '% do tempo de jogo', share, SKY, NAVY, '', PCT);
    tile(sh, r0 + 1, 8, 'Titular', p.st ? 'Sim' : 'Não'); tile(sh, r0 + 1, 11, 'Momentos altos', p.hl || 0);
    sh.mergeCells(r0 + 3, 2, r0 + 3, 12); const bar = sh.getCell(r0 + 3, 2); const n = Math.round(share * 40);
    bar.value = '█'.repeat(n) + '░'.repeat(40 - n); bar.font = font({ color: BLUE }); sh.getRow(r0 + 3).height = 16;
    tile(sh, r0 + 5, 2, 'Golos', p.g || 0); tile(sh, r0 + 5, 5, 'Remates', shots);
    tile(sh, r0 + 5, 8, 'À baliza', p.son || 0, SKY, NAVY, shots ? pct(p.son || 0, shots) : ''); tile(sh, r0 + 5, 11, 'Eficácia (golos/remate)', FRAC(p.g || 0, shots), SKY, NAVY, '', PCT);
    tile(sh, r0 + 8, 2, 'GM em campo', p.gm || 0, GKF, GREEN); tile(sh, r0 + 8, 5, 'GS em campo', p.gs || 0, REDF, RED);
    tile(sh, r0 + 8, 8, 'Saldo +/-', pm, SKY, pm > 0 ? GREEN : pm < 0 ? RED : NAVY, '', PM); tile(sh, r0 + 8, 11, 'Remates fora', p.soff || 0);
    tile(sh, r0 + 11, 2, 'Faltas cometidas', p.fc || 0); tile(sh, r0 + 11, 5, 'Faltas sofridas', p.fs || 0);
    tile(sh, r0 + 11, 8, 'Amarelos', p.y || 0, YEL); tile(sh, r0 + 11, 11, 'Vermelhos', p.r || 0, REDC);
    let row = r0 + 14;
    if (p.gk) {
      title(sh, 'B' + row, 'Guarda-redes'); row++;
      tile(sh, row, 2, 'Defesas', p.sv || 0, GKF, GREEN); tile(sh, row, 5, 'Golos sofridos', p.gs || 0, REDF, RED);
      tile(sh, row, 8, '% defesas', FRAC(p.sv || 0, (p.sv || 0) + (p.gs || 0)), GKF, GREEN, '', PCT); row += 3;
    }
    if (p.zs && sum(p.zs)) { title(sh, 'B' + row, 'Onde rematou'); row = goalMap(sh, row + 1, 2, shootCells(p.zs, p.zg || [0, 0, 0, 0, 0, 0]), 3); note(sh, 'B' + (row - 1), readShoot(p.zs, p.zg || [], 'Rematou')); row++; }
    if (p.gk && p.zsv && (sum(p.zsv) + sum(p.zga))) { title(sh, 'B' + row, 'Remates à baliza que recebeu'); row = goalMap(sh, row + 1, 2, gkCells(p.zsv, p.zga || [0, 0, 0, 0, 0, 0]), 3); }
    if (!(p.zs && sum(p.zs)) && !(p.gk && p.zsv && sum(p.zsv))) { note(sh, 'B' + row, 'Sem zonas da baliza registadas neste jogo.'); row += 2; }
    return row;
  }
  function zoneLabel(z) {
    const mx = Math.max(...z), mn = Math.min(...z); if (!mx) return '–';
    const nums = v => { const n = z.map((x, i) => x === v ? i + 1 : 0).filter(Boolean); return n.length > 3 ? n.slice(0, 3).join(',') + '…' : n.join(','); };
    return mx === mn ? 'igual ' + nums(mx) : 'mais ' + nums(mx) + ' · menos ' + nums(mn);
  }

  // ---------- livro da época ----------
  function seasonBook(ExcelJS, H, A, team, focus) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'Scouting do Banco';
    const used = new Set(), ws = wb.addWorksheet(sheetName(used, 'Época'), { views: [{ showGridLines: false }] });
    const ps = A.slice().sort((a, b) => a.num - b.num), names = new Map(ps.map(a => [a, sheetName(used, a.num + ' ' + a.name)]));
    const pidName = new Map(ps.map(a => [a.pid, names.get(a)]));
    ws.columns = Array.from({ length: 20 }, (_, i) => ({ width: i === 1 ? 20 : i === 12 || i === 13 ? 20 : 9 }));
    // em cima: o último jogo (ou o jogo pedido), só com os dados desse jogo
    if (!focus && H.length) focus = H.reduce((x, y) => (y.date > x.date ? y : x));
    let o = 0;
    if (focus) o = gameSection(ws, focus, focus.players.slice().sort((a, b) => a.num - b.num), p => pidName.get(p.pid), 0, { label: 'Último jogo · ', season: true }) + 1;
    band(ws, `A${o + 1}:T${o + 1}`, `Época ${team || ''} · ${H.length} jogo${H.length === 1 ? '' : 's'} · acumulado`);
    const w = H.filter(g => g.us > g.them).length, d = H.filter(g => g.us === g.them).length, l = H.length - w - d;
    const gf = sum(H.map(g => g.us)), ga = sum(H.map(g => g.them)), HS = H.filter(g => g.shUs != null);
    const tiles = [['Jogos', H.length], ['V-E-D', `${w}-${d}-${l}`], ['Golos', `${gf}–${ga}`], ['Pontos', w * 3 + d],
      ['Remates/jogo', HS.length ? sum(HS.map(g => g.shUs)) / HS.length : '–'],
      ['% defesas GR', HS.length ? FRAC(sum(HS.map(g => g.sv)), sum(HS.map(g => g.shOnThem))) : '–']];
    tiles.forEach(([lab, v], i) => {
      const c = 1 + i * 3; for (const r of [o + 3, o + 4]) { ws.mergeCells(r, c, r, c + 2); ws.getCell(r, c).fill = fill(SKY); }
      const a = ws.getCell(o + 3, c); a.value = v; a.font = font({ bold: true, size: 22, color: NAVY }); a.alignment = center; if (typeof v === 'number') a.numFmt = lab.startsWith('%') ? PCT : lab === 'Remates/jogo' ? DEC : '0';
      const b = ws.getCell(o + 4, c); b.value = lab; b.font = font({ bold: true, size: 9, color: MUTED }); b.alignment = center;
    });
    ws.getRow(o + 3).height = 40;
    // mapas da equipa
    const us = { zs: [0, 0, 0, 0, 0, 0], zg: [0, 0, 0, 0, 0, 0] }, th = { zsv: [0, 0, 0, 0, 0, 0], zga: [0, 0, 0, 0, 0, 0] };
    const add = (t, z) => { if (z) z.forEach((v, i) => t[i] += v || 0); };
    H.forEach(g => { if (g.zUs) { add(us.zs, g.zUs); add(us.zg, g.zUsG); add(th.zsv, g.zSv); add(th.zga, g.zGa); } else g.players.forEach(r => { add(us.zs, r.zs); add(us.zg, r.zg); if (r.gk) { add(th.zsv, r.zsv); add(th.zga, r.zga); } }); });
    let row = o + 6;
    if (sum(us.zs) || sum(th.zsv) + sum(th.zga)) {
      title(ws, 'A' + row, 'Onde rematamos'); title(ws, 'J' + row, 'Onde sofremos');
      if (sum(us.zs)) goalMap(ws, row + 1, 1, shootCells(us.zs, us.zg)); else note(ws, 'A' + (row + 1), 'Sem zonas registadas.');
      if (sum(th.zsv) + sum(th.zga)) goalMap(ws, row + 1, 10, gkCells(th.zsv, th.zga)); else note(ws, 'J' + (row + 1), 'Sem zonas registadas.');
      if (sum(us.zs)) note(ws, 'A' + (row + 6), readShoot(us.zs, us.zg, 'Rematamos'));
      row += 8;
    }
    title(ws, 'A' + row, 'Jogadores · clica no nome para abrir a ficha da época'); row++;
    const hdr = ['N.º', 'Jogador', 'GR', 'Jogos', 'Titular', 'Minutos', 'Min/jogo', 'Golos', 'Golos/40', 'Saldo +/-', 'Remates', '% à baliza', 'Zonas', 'Defesas', '% defesas', 'Faltas c.', 'Faltas s.', 'Amarelos', 'Vermelhos'];
    const hr = ws.getRow(row); hr.values = hdr; hr.height = 30;
    hr.eachCell(c => { c.font = font({ bold: true, size: 9, color: NAVY }); c.fill = fill(SKY); c.alignment = center; c.border = thin; });
    const H0 = row;
    ps.forEach((a, k) => {
      const r = H0 + 1 + k, rw = ws.getRow(r);
      rw.values = [a.num, '', a.gk ? 'GR' : '', a.called, a.st, mins(a.min), a.played ? mins(a.mpg) : '–', a.g, a.g40 < 0 ? '–' : a.g40, a.pm,
        a.sht, a.acc < 0 ? '–' : a.acc, sum(a.zs) ? zoneLabel(a.zs) : '–', a.sv, a.svp < 0 ? '–' : a.svp, a.fc, a.fs, a.y, a.r];
      fmtCols(rw, { 9: DEC, 10: PM, 12: PCT, 15: PCT });
      rw.eachCell({ includeEmpty: true }, (c, n) => { if (n > 19) return; c.font = font(); c.alignment = { horizontal: n === 2 ? 'left' : 'center', vertical: 'middle' }; c.border = thin; });
      link(ws, 'B' + r, a.name, names.get(a)).font = font({ bold: true, underline: true, color: BLUE });
      ws.getCell(r, 10).font = font({ bold: true, color: a.pm > 0 ? GREEN : a.pm < 0 ? RED : MUTED });
      if (a.g) { ws.getCell(r, 8).fill = fill('FFC9EBD7'); ws.getCell(r, 8).font = font({ bold: true, color: GREEN }); }
      if (a.y) { ws.getCell(r, 18).fill = fill('FFF2C230'); ws.getCell(r, 18).font = font({ bold: true }); }
      if (a.r) { ws.getCell(r, 19).fill = fill(RED); ws.getCell(r, 19).font = font({ bold: true, color: WHITE }); }
    });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 0, showGridLines: false }];
    // folha com os jogos
    const wj = wb.addWorksheet(sheetName(used, 'Jogos'), { views: [{ showGridLines: false }] });
    wj.columns = [{ width: 12 }, { width: 22 }, { width: 10 }, { width: 11 }, { width: 11 }, { width: 11 }, { width: 11 }, { width: 11 }];
    band(wj, 'A1:H1', 'Jogos da época', 16);
    note(wj, 'A2', 'Clica no adversário para abrir esse jogo com os dados de cada jogador.');
    const HJ = H.slice().sort((a, b) => a.date - b.date), dm = ms => { const d = new Date(ms); return String(d.getDate()).padStart(2, '0') + '-' + String(d.getMonth() + 1).padStart(2, '0'); };
    const gName = new Map(HJ.map(g => [g, sheetName(used, dm(g.date) + ' ' + g.opp)]));
    const jh = wj.getRow(3); jh.values = ['Data', 'Adversário', 'Casa/Fora', 'Resultado', 'Remates', 'À baliza', 'Remates eles', 'Defesas GR'];
    jh.eachCell(c => { c.font = font({ bold: true, size: 9, color: NAVY }); c.fill = fill(SKY); c.alignment = center; });
    HJ.forEach((g, i) => {
      const r = wj.getRow(4 + i); r.values = [dstr(g.date), g.opp, venue(g.home), g.us + '–' + g.them, g.shUs ?? '', g.shOnUs ?? '', g.shThem ?? '', g.sv ?? ''];
      r.eachCell({ includeEmpty: true }, (c, n) => { if (n > 8) return; c.font = font(); c.alignment = { horizontal: n === 2 ? 'left' : 'center' }; c.border = thin; });
      link(wj, 'B' + (4 + i), g.opp, gName.get(g)).font = font({ bold: true, underline: true, color: BLUE });
      const res = wj.getCell(4 + i, 4); res.font = font({ bold: true, color: WHITE }); res.fill = fill(g.us > g.them ? GREEN : g.us < g.them ? RED : 'FF617789');
    });
    // uma folha por jogo, igual à do Excel de um jogo
    HJ.forEach(g => {
      const sg = wb.addWorksheet(gName.get(g), { views: [{ showGridLines: false }] });
      fillGameSheet(sg, g, g.players.slice().sort((a, b) => a.num - b.num), p => pidName.get(p.pid), { text: '← Voltar aos jogos', target: wj.name });
    });
    // fichas da época
    ps.forEach(a => {
      const sh = wb.addWorksheet(names.get(a), { views: [{ showGridLines: false }] });
      sh.columns = [2, 13, 13, 2, 13, 13, 2, 13, 13, 2, 13, 13].map(w => ({ width: w }));
      link(sh, 'B1', '← Voltar à época', ws.name).font = font({ bold: true, underline: true, color: BLUE });
      band(sh, 'B2:L2', `#${a.num} ${a.name}` + (a.gk ? '  ·  GR' : ''));
      note(sh, 'B3', `Convocado em ${a.called} de ${H.length} jogos · jogou ${a.played} · titular ${a.st}`);
      const teamMin = sum(H.map(g => g.dur)), p40 = (v) => a.min >= 300 ? v * 2400 / a.min : '–';
      // em cima: o último jogo (só esse jogo); por baixo: o acumulado da época
      let s0 = 5;
      if (focus) {
        const fr = focus.players.find(x => x.pid === a.pid), fn = focus.opp + (focus.home != null ? ' (' + (focus.home ? 'C' : 'F') + ')' : '');
        subBand(sh, 5, `Último jogo · ${focus.us}–${focus.them} ${fn} · ${dstr(focus.date)}`);
        if (gName.get(focus)) link(sh, 'B6', 'Abrir este jogo', gName.get(focus)).font = font({ size: 9, underline: true, color: BLUE });
        if (fr) { note(sh, 'E6', `${fr.st ? 'Titular' : 'Suplente'} · ${mmss(fr.min)} jogados`); s0 = playerGameBlock(sh, fr, focus, 6) + 1; }
        else { note(sh, 'E6', 'Não esteve convocado neste jogo.'); s0 = 8; }
        subBand(sh, s0, `Época · acumulado de ${H.length} jogo${H.length === 1 ? '' : 's'}`); s0++;
      } else title(sh, 'B5', 'Na época');
      tile(sh, s0 + 1, 2, 'Minutos', mins(a.min)); tile(sh, s0 + 1, 5, 'Min/jogo', a.played ? mins(a.mpg) : '–');
      tile(sh, s0 + 1, 8, '% do tempo', FRAC(a.min, teamMin), SKY, NAVY, '', PCT); tile(sh, s0 + 1, 11, 'Momentos altos', a.hl);
      tile(sh, s0 + 4, 2, 'Golos', a.g); tile(sh, s0 + 4, 5, 'Golos/40 min', p40(a.g), SKY, NAVY, '', DEC); tile(sh, s0 + 4, 8, 'Remates', a.sht);
      tile(sh, s0 + 4, 11, 'Eficácia (golos/remate)', FRAC(a.g, a.sht), SKY, NAVY, '', PCT);
      tile(sh, s0 + 7, 2, 'GM em campo', a.gm, GKF, GREEN); tile(sh, s0 + 7, 5, 'GS em campo', a.gs, REDF, RED);
      tile(sh, s0 + 7, 8, 'Saldo +/-', a.pm, SKY, a.pm > 0 ? GREEN : a.pm < 0 ? RED : NAVY, '', PM); tile(sh, s0 + 7, 11, 'À baliza', a.son, SKY, NAVY, a.sht ? pct(a.son, a.sht) : '');
      tile(sh, s0 + 10, 2, 'Faltas cometidas', a.fc); tile(sh, s0 + 10, 5, 'Faltas sofridas', a.fs); tile(sh, s0 + 10, 8, 'Amarelos', a.y, YEL); tile(sh, s0 + 10, 11, 'Vermelhos', a.r, REDC);
      let row = s0 + 13;
      if (a.gk) { title(sh, 'B' + row, 'Guarda-redes'); row++; tile(sh, row, 2, 'Defesas', a.sv, GKF, GREEN); tile(sh, row, 5, 'Golos sofridos', a.gs, REDF, RED); tile(sh, row, 8, '% defesas', a.svp < 0 ? '–' : a.svp, GKF, GREEN, '', PCT); row += 3; }
      if (sum(a.zs)) { title(sh, 'B' + row, 'Onde remata'); row = goalMap(sh, row + 1, 2, shootCells(a.zs, a.zg), 3); note(sh, 'B' + (row - 1), readShoot(a.zs, a.zg, 'Remata')); row++; }
      if (a.gk && (sum(a.zsv) + sum(a.zga))) { title(sh, 'B' + row, 'Remates à baliza que recebeu'); row = goalMap(sh, row + 1, 2, gkCells(a.zsv, a.zga), 3); }
      // jogo a jogo
      title(sh, 'B' + row, 'Jogo a jogo'); row++;
      const th = sh.getRow(row); th.values = ['', 'Data', 'Adversário', '', 'Resultado', 'Minutos', '', 'Golos', 'Saldo +/-', '', a.gk ? 'Defesas' : 'Remates', a.gk ? 'Golos sofridos' : 'Faltas c.'].concat(a.gk ? ['Remates'] : []);
      if (a.gk) sh.getColumn(13).width = 13;
      th.eachCell(c => { if (c.value) { c.font = font({ bold: true, size: 9, color: NAVY }); c.fill = fill(SKY); c.alignment = center; } });
      sh.mergeCells(row, 3, row, 4);
      a.games.slice().sort((x, y) => x.g.date - y.g.date).forEach(({ g, r }, i) => {
        const rr = sh.getRow(row + 1 + i), pm = (r.gm || 0) - (r.gs || 0);
        rr.values = ['', dstr(g.date), g.opp + (g.home != null ? ' (' + (g.home ? 'C' : 'F') + ')' : ''), '', g.us + '–' + g.them, mins(r.min), '', r.g, pm, '', a.gk ? (r.sv || 0) : (r.son || 0) + (r.soff || 0), a.gk ? (r.gs || 0) : r.fc].concat(a.gk ? [(r.son || 0) + (r.soff || 0)] : []);
        fmtCols(rr, { 9: PM });
        rr.eachCell(c => { c.font = font(); c.alignment = { horizontal: 'center' }; });
        sh.mergeCells(row + 1 + i, 3, row + 1 + i, 4);
        if (gName.get(g)) link(sh, 'C' + (row + 1 + i), rr.getCell(3).value, gName.get(g)).font = font({ underline: true, color: BLUE });
        sh.getCell(row + 1 + i, 3).alignment = { horizontal: 'left', shrinkToFit: true };
        sh.getCell(row + 1 + i, 9).font = font({ bold: true, color: pm > 0 ? GREEN : pm < 0 ? RED : MUTED });
      });
    });
    return wb;
  }

  async function deliver(wb, filename) {
    const buf = await wb.xlsx.writeBuffer();
    const type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    // Dentro da página do Claude: a capacidade "downloads" pede confirmação ao utilizador e guarda o ficheiro.
    if (window.claude && typeof window.claude.use === 'function') {
      let dl = null; try { dl = await window.claude.use('downloads'); } catch (e) { dl = null; }
      if (dl) {
        try { await dl.save({ filename, data: new Blob([buf], { type }) }); return 'downloaded'; }
        catch (e) { if (e && e.code === 'declined') return 'cancelled'; if (e && e.code === 'rate_limited') throw Object.assign(new Error('busy'), { code: 'rate_limited' }); throw e; }
      }
    }
    const file = new File([buf], filename, { type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: filename }); return 'shared'; } catch (e) { if (e && e.name === 'AbortError') return 'cancelled'; }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    return 'downloaded';
  }

  let loading = null;
  function loadLib() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!loading) loading = new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'vendor/exceljs.min.js'; s.onload = () => res(window.ExcelJS); s.onerror = () => { loading = null; rej(new Error('lib')); }; document.head.appendChild(s); });
    return loading;
  }
  // Nome do ficheiro sem acentos nem caracteres especiais (alguns browsers recusam-nos).
  const safe = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9 ._-]+/g, '-').replace(/\s+/g, ' ').trim();
  window.ScoutXLSX = {
    async game(g) { const X = await loadLib(); const d = new Date(g.date); return deliver(gameBook(X, g), safe(`${g.team} ${g.us}-${g.them} ${g.opp} ${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}.xlsx`)); },
    async season(H, A, team, focus) { const X = await loadLib(); return deliver(seasonBook(X, H, A, team, focus), safe(`Época ${team || ''} ${new Date().toLocaleDateString('pt-PT').replace(/\//g, '-')}.xlsx`)); },
    _build: { gameBook, seasonBook, loadLib }
  };
})();
