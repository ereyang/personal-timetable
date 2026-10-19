// 전체시간표(첫 번째 시트) → 날짜별 수업 데이터
// 브라우저(window.XLSX)와 node(require('xlsx')) 모두에서 동작
(function (root) {
  const DAYS = ['월', '화', '수', '목', '금'];

  function u8(buf) { return buf instanceof Uint8Array ? buf : new Uint8Array(buf); }
  function dec(c) { return new TextDecoder('utf-8').decode(c.content instanceof Uint8Array ? c.content : new Uint8Array(c.content)); }

  // 빗금(대각선 테두리) 셀 = 삭제된 수업
  function diagonalCells(X, data, sheetIndex) {
    const out = new Set();
    try {
      const zip = X.CFB.read(u8(data), { type: 'array' });
      const get = p => { const f = X.CFB.find(zip, p); return f ? dec(f) : ''; };
      const wbx = get('/xl/workbook.xml'), rels = get('/xl/_rels/workbook.xml.rels');
      const sheets = [...wbx.matchAll(/<sheet [^>]*?r:id="([^"]+)"/g)].map(m => m[1]);
      const rid = sheets[sheetIndex];
      const rel = [...rels.matchAll(/<Relationship [^>]*>/g)].map(m => m[0]).find(r => r.includes('Id="' + rid + '"'));
      let target = rel.match(/Target="([^"]+)"/)[1].replace(/^\/?(xl\/)?/, '');
      const sheet = get('/xl/' + target), styles = get('/xl/styles.xml');
      const bl = (styles.match(/<borders[^>]*>([\s\S]*?)<\/borders>/) || [, ''])[1].match(/<border\b[^>]*\/>|<border\b[\s\S]*?<\/border>/g) || [];
      const diagB = new Set();
      bl.forEach((b, i) => {
        const head = b.match(/<border\b[^>]*>/)[0];
        const d = b.match(/<diagonal\b[^>]*?(\/>|>)/);
        if (/diagonal(Up|Down)="(1|true)"/.test(head) && d && /style="/.test(d[0])) diagB.add(i);
      });
      const xl = (styles.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/) || [, ''])[1].match(/<xf\b[^>]*?(\/>|>[\s\S]*?<\/xf>)/g) || [];
      const dx = new Set();
      xl.forEach((x, i) => { const m = x.match(/borderId="(\d+)"/); if (m && diagB.has(+m[1])) dx.add(String(i)); });
      for (const m of sheet.matchAll(/<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"/g)) if (dx.has(m[2])) out.add(m[1]);
    } catch (e) { console.warn('빗금 읽기 실패', e); }
    return out;
  }

  function parseDate(txt, baseYear) {
    const m = String(txt).match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일?/);
    if (!m) return null;
    const mo = +m[1], d = +m[2];
    const y = mo < 3 ? baseYear + 1 : baseYear;
    const note = String(txt).replace(m[0], '').replace(/^[\s(]+|[\s)]+$/g, '').trim();
    return { iso: y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0'), note };
  }

  const clean = v => String(v == null ? '' : v).replace(/\r/g, '').trim();
  const splitNames = v => clean(v).split(/[,\n\/·]+/).map(s => s.trim()).filter(Boolean);

  function placeOf(label) {
    const l = label.replace(/\s+/g, ' ').trim();
    const m = l.match(/물리학\s*2.*?\(([^)]+)\)/);
    if (m) return { place: m[1].trim(), sub: '물리학2' };
    if (/^\d+반$/.test(l)) return { place: l };
    return { place: l };
  }

  function parse(data, opts) {
    opts = opts || {};
    const X = opts.XLSX || root.XLSX;
    const now = opts.today ? new Date(opts.today) : new Date();
    const baseYear = opts.baseYear || (now.getMonth() < 2 ? now.getFullYear() - 1 : now.getFullYear());
    const wb = X.read(u8(data), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const diag = diagonalCells(X, data, 0);
    const rng = X.utils.decode_range(ws['!ref']);
    const cell = (r, c) => { const x = ws[X.utils.encode_cell({ r, c })]; return x ? clean(x.w != null && typeof x.v !== 'string' ? x.w : x.v) : ''; };
    const merges = ws['!merges'] || [];
    const mergeAt = (r, c) => merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);

    // 날짜 머리행: '월' 다음 행이 '1교시'
    const blocks = [];
    for (let r = rng.s.r; r <= rng.e.r - 2; r++) {
      const dayRow = r + 1, perRow = r + 2;
      const days = [];
      for (let c = 1; c <= Math.min(rng.e.c, 40); c++) {
        const v = cell(dayRow, c);
        if (DAYS.includes(v)) {
          const dt = parseDate(cell(r, c), baseYear);
          if (dt) days.push({ c, day: v, ...dt });
        }
      }
      if (days.length < 3) continue;
      days.forEach((d, i) => {
        const end = i + 1 < days.length ? days[i + 1].c - 1 : d.c + 7;
        d.periods = [];
        for (let c = d.c; c <= end; c++) {
          const m = cell(perRow, c).match(/(\d+)\s*교시/);
          if (m) d.periods.push({ c, p: +m[1] });
        }
        // 교시 머리칸까지 병합된 날(성탄절·신정 등): 칸 순서대로 교시 번호
        if (!d.periods.length) for (let c = d.c, p = 1; c <= end && p <= 6; c++, p++) d.periods.push({ c, p });
      });
      blocks.push({ r, perRow, days });
      r = perRow;
    }

    const lessons = [];       // {date, p, place, teacher, deleted, sub}
    const dayInfo = {};       // date → {day, note, events:[{text, periods}], notes:[], excluded:[]}
    const teachers = new Set();
    const eventSeen = new Set();
    const rangeNotes = [];    // 시간표 블록이 없는 기간 안내 (기말고사 등)

    blocks.forEach((b, bi) => {
      const lastRow = bi + 1 < blocks.length ? blocks[bi + 1].r - 1 : rng.e.r;
      b.days.forEach(d => {
        dayInfo[d.iso] = dayInfo[d.iso] || { day: d.day, note: d.note, events: [], notes: [], excluded: [], maxP: Math.max(0, ...d.periods.map(x => x.p)) };
      });
      // 행 라벨
      let classRowsEnd = b.perRow;
      for (let r = b.perRow + 1; r <= lastRow; r++) {
        const label = clean(cell(r, 0)).replace(/\s+/g, ' ');
        if (!label) {
          // 라벨 없는 행: 날짜별 메모 (예: 현장체험 임장, 보강예정)
          b.days.forEach(d => d.periods.forEach(({ c }) => {
            const m = mergeAt(r, c);
            const v = m ? cell(m.s.r, m.s.c) : cell(r, c);
            if (v && !dayInfo[d.iso].notes.includes(v)) dayInfo[d.iso].notes.push(v);
          }));
          continue;
        }
        if (/^\d+\s*월.*\d+\s*일/.test(label) && !/반$/.test(label)) {
          // "11월 23일~11월 27일 3학년 기말고사" 같은 안내 행
          const ds = [...label.matchAll(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/g)].map(m => parseDate(m[0], baseYear).iso);
          const text = label.replace(/^.*\d+\s*일\s*/, '').trim();
          if (ds.length) rangeNotes.push({ from: ds[0], to: ds[ds.length - 1], text: text || label });
          continue;
        }
        if (/제외/.test(label)) {
          b.days.forEach(d => {
            const v = d.periods.map(({ c }) => cell(r, c)).find(Boolean) || cell(r, d.c);
            splitNames(v).forEach(n => { dayInfo[d.iso].excluded.push(n); });
          });
          continue;
        }
        const { place, sub } = placeOf(label);
        classRowsEnd = r;
        b.days.forEach(d => d.periods.forEach(({ c, p }) => {
          const m = mergeAt(r, c);
          if (m && m.e.r > m.s.r) {
            const v = cell(m.s.r, m.s.c);
            if (!v) return;
            // 여러 칸 병합 = 행사 (학력평가, 수능, 재량휴업일 …)
            const key = d.iso + '|' + v + '|' + m.s.r + '|' + m.s.c;
            if (!eventSeen.has(key)) {
              eventSeen.add(key);
              const ps = d.periods.filter(x => x.c >= m.s.c && x.c <= m.e.c).map(x => x.p);
              dayInfo[d.iso].events.push({ text: v, periods: ps });
            }
            return;
          }
          // 한 행 안의 가로 병합(연속 수업)은 각 교시에 같은 이름
          const v = m ? cell(m.s.r, m.s.c) : cell(r, c);
          if (!v) return;
          const addr = X.utils.encode_cell(m ? m.s : { r, c });
          const deleted = diag.has(addr);
          splitNames(v).forEach(t => {
            if (!deleted) teachers.add(t);
            lessons.push({ date: d.iso, p, place, teacher: t, deleted, ...(sub ? { sub } : {}) });
          });
        }));
      }
    });

    // 삭제된 칸에만 나온 이름도 목록에 포함
    lessons.forEach(l => teachers.add(l.teacher));
    // 당일수업제외 이름 검증
    Object.values(dayInfo).forEach(di => { di.excluded = [...new Set(di.excluded)]; });

    const dates = Object.keys(dayInfo).sort();
    return {
      title: clean(cell(0, 0)) || wb.SheetNames[0],
      sheet: wb.SheetNames[0],
      teachers: [...teachers].sort((a, b) => a.localeCompare(b, 'ko')),
      dates, dayInfo, lessons, rangeNotes,
      from: dates[0], to: dates[dates.length - 1],
    };
  }

  // 요일별 전체시간표(1·2·3학년 칸, 셀 "과목\n교사") → 지정 학년만 추출
  function parseWeekly(data, opts) {
    opts = opts || {};
    const X = opts.XLSX || root.XLSX;
    const keep = opts.grades || [1, 2];
    const wb = X.read(u8(data), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rng = X.utils.decode_range(ws['!ref']);
    const raw = (r, c) => { const x = ws[X.utils.encode_cell({ r, c })]; return x && x.v != null ? String(x.v) : ''; };
    const merges = ws['!merges'] || [];
    const mv = (r, c) => { const m = merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c); return m ? raw(m.s.r, m.s.c) : raw(r, c); };

    // '교시' 머리칸 찾기
    let hr = -1, pc = -1;
    for (let r = 0; r <= Math.min(rng.e.r, 10) && hr < 0; r++)
      for (let c = 0; c <= Math.min(rng.e.c, 5); c++) if (clean(raw(r, c)) === '교시') { hr = r; pc = c; break; }
    if (hr < 0) throw new Error('‘교시’ 머리칸을 찾지 못했어요.');
    // 학년 행(hr) + 반 행(hr+1)
    const cols = [];
    for (let c = pc + 1; c <= rng.e.c; c++) {
      const g = clean(mv(hr, c)).match(/(\d)\s*학년/);
      if (!g || !keep.includes(+g[1])) continue;
      const cls = clean(raw(hr + 1, c));
      if (!cls) continue;
      cols.push({ c, grade: +g[1], cls });
    }
    if (!cols.length) throw new Error('1·2학년 칸을 찾지 못했어요.');
    // 같은 이름 칸(교과교실 여러 개) 번호 붙이기
    const seen = {};
    cols.forEach(o => {
      const k = o.grade + '|' + o.cls;
      seen[k] = (seen[k] || 0) + 1;
    });
    const cnt = {};
    cols.forEach(o => {
      const k = o.grade + '|' + o.cls;
      cnt[k] = (cnt[k] || 0) + 1;
      o.label = /^\d+$/.test(o.cls) ? o.grade + '-' + o.cls : o.grade + '학년 ' + o.cls + (seen[k] > 1 ? cnt[k] : '');
      o.key = o.label;
    });

    const lessons = [];   // {dow, p, col, subject, teacher}
    const maxP = {};
    const teachers = new Set();
    for (let r = hr + 2; r <= rng.e.r; r++) {
      const dow = clean(mv(r, pc - 1 >= 0 ? pc - 1 : 0)).charAt(0);
      const p = parseInt(clean(raw(r, pc)), 10);
      if (!DAYS.includes(dow) || !p) continue;
      maxP[dow] = Math.max(maxP[dow] || 0, p);
      cols.forEach(o => {
        const lines = raw(r, o.c).replace(/\r/g, '\n').split('\n').map(s => s.trim()).filter(Boolean);
        if (!lines.length) return;
        const teacher = lines.length > 1 ? lines[lines.length - 1] : '';
        const subject = lines.length > 1 ? lines.slice(0, -1).join(' ') : lines[0];
        splitNames(teacher).forEach(t => teachers.add(t));
        lessons.push({ dow, p, col: o.key, subject, teacher });
      });
    }
    return {
      title: clean(raw(0, 0)) || wb.SheetNames[0],
      cols: cols.map(({ grade, cls, label, key }) => ({ grade, cls, label, key })),
      lessons, maxP,
      teachers: [...teachers].sort((a, b) => a.localeCompare(b, 'ko')),
    };
  }

  const api = { parse, parseWeekly };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TTParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
