// 3학년 개인시간표 앱 저장소 (GitHub Pages 앱에서 fetch로 사용)
// GET  ?a=load            → 저장된 data(3학년)/g12(1·2학년)/cal(특별한 날)/chg(수업 교체) 전부 (누구나)
// POST {a:'setpw',pw}      → 담당자 암호가 없을 때만 처음 정하기
// POST {a:'check',pw}      → 담당자 암호 확인
// POST {a:'save',pw,set}   → set의 키들(data) 저장 (담당자만)
// POST {a:'chpw',pw,newPw} → 암호 바꾸기
var KEYS = ['data', 'g12', 'cal', 'chg'];
var CHUNK = 2500; // 속성 하나당 9KB 제한 → 한글 3바이트 고려해 2500자씩 나눠 저장

function props() { return PropertiesService.getScriptProperties(); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'tt3:' + s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function readKey(k) {
  var p = props(), n = Number(p.getProperty('k:' + k + ':n') || 0);
  if (!n) return null;
  var s = '';
  for (var i = 0; i < n; i++) s += p.getProperty('k:' + k + ':' + i) || '';
  try { return JSON.parse(s); } catch (e) { return null; }
}
function writeKey(k, v) {
  var p = props(), s = JSON.stringify(v), old = Number(p.getProperty('k:' + k + ':n') || 0);
  var obj = {}, n = Math.ceil(s.length / CHUNK);
  for (var i = 0; i < n; i++) obj['k:' + k + ':' + i] = s.slice(i * CHUNK, (i + 1) * CHUNK);
  obj['k:' + k + ':n'] = String(n);
  p.setProperties(obj);
  for (var j = n; j < old; j++) p.deleteProperty('k:' + k + ':' + j);
}

function doGet(e) {
  var p = props(), r = { ok: true, hasPw: !!p.getProperty('pw'), updated: p.getProperty('updated') || '' };
  KEYS.forEach(function (k) { r[k] = readKey(k); });
  return json(r);
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, msg: '잘못된 요청' }); }
  var p = props(), pw = p.getProperty('pw');
  if (b.a === 'setpw') {
    if (pw) return json({ ok: false, msg: '이미 담당자 암호가 있어요.' });
    if (!b.pw || String(b.pw).length < 4) return json({ ok: false, msg: '암호는 4자 이상으로 해 주세요.' });
    p.setProperty('pw', hash(b.pw));
    return json({ ok: true });
  }
  if (!pw || hash(b.pw || '') !== pw) return json({ ok: false, msg: '담당자 암호가 틀려요.' });
  if (b.a === 'check') return json({ ok: true });
  if (b.a === 'chpw') {
    if (!b.newPw || String(b.newPw).length < 4) return json({ ok: false, msg: '암호는 4자 이상으로 해 주세요.' });
    p.setProperty('pw', hash(b.newPw));
    return json({ ok: true });
  }
  if (b.a === 'save') {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      Object.keys(b.set || {}).forEach(function (k) { if (KEYS.indexOf(k) >= 0) writeKey(k, b.set[k]); });
      var now = new Date().toISOString();
      p.setProperty('updated', now);
      return json({ ok: true, updated: now });
    } finally { lock.releaseLock(); }
  }
  return json({ ok: false, msg: '알 수 없는 요청' });
}
