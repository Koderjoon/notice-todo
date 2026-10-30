/******************************************************
 * 디데이 25 — 카톡 공지 → 마감·일정 정리 (Google Apps Script)
 *
 *  - 드라이브 폴더에 올린 '대화 내용 내보내기' 파일로 목록을 매번 다시 만듦
 *    (최근 2주 공지 기준 → 삭제·수정된 공지도 반영)
 *  - 관리자(TOKEN)는 목록 페이지에서 직접 추가·수정·삭제 → 다시 만들어도 유지
 *  - 학급(VIEW_KEY)은 보기만
 *
 *  ※ 실제 기능 코드(이 파일, apps-script/app.js)는 GitHub에만 있습니다.
 *    Apps Script에는 이 파일을 불러와 실행하는 작은 로더(Code.gs)만 붙여넣어 두고,
 *    로더가 실행 때마다 최신 app.js를 가져와 씁니다(5분 캐시). 배포를 다시 할 필요 없음.
 *  ※ 비밀값은 코드에 두지 않고 [프로젝트 설정 → 스크립트 속성]에 둡니다:
 *    TOKEN(관리자 키), VIEW_KEY(학급 보기 키), GEMINI_API_KEY (또는 ANTHROPIC_API_KEY)
 ******************************************************/

const CONFIG = {
  ROOMS: ['25 실습공지', '전남대치전원 25 공지방'],   // 카톡 방 이름 그대로
  LLM: 'gemini',                                 // 'gemini' 또는 'claude'
  GEMINI_MODEL: 'gemini-flash-lite-latest',
  GEMINI_FALLBACKS: ['gemini-2.5-flash', 'gemini-flash-latest'],   // 붐빌 때 차례로 시도
  CLAUDE_MODEL: 'claude-haiku-5-5',
  EXPORT_FOLDER: '카톡공지_내보내기',
  REBUILD_DAYS: 21,          // 내보내기 파일에서 최근 며칠 공지로 목록을 만들지
  SUPPLIES: false,           // 실습 준비물 기능 (true로 바꾸면 다시 켜짐)
  // 요일별 실습 (1=월 … 5=금)
  LABS: {
    1: ['소아치과학'],
    2: ['예방치과학', '금관가공의치학', '보존수복학'],
    3: ['국소의치학', '근관치료학'],
    4: ['예방치과학', '치주학'],
    5: ['총의치학', '교정학'],
  },
  BATCH: 40,                 // 한 번에 AI에 넘길 메시지 수
};

const TZ = 'Asia/Seoul';
const SH_TASK = 'tasks';
const SH_HIDDEN = 'hidden';
const TASK_HEAD = ['id', 'title', 'due', 'room', 'detail', 'status', 'created', 'updated', 'source', 'src'];
const HIDDEN_HEAD = ['src', 'titleKey', 'dueDate', 'title', 'created'];
const SH_SUP = 'supplies';
const SUP_HEAD = ['id', 'subject', 'date', 'items', 'detail', 'room', 'status', 'updated', 'source', 'src'];
const HEADS = { tasks: TASK_HEAD, hidden: HIDDEN_HEAD, supplies: SUP_HEAD };

/* ---------- 최초 1회(그리고 이번 업데이트 후 1회) 실행 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SH_TASK, TASK_HEAD);
  ensureSheet_(ss, SH_HIDDEN, HIDDEN_HEAD);
  ensureSheet_(ss, SH_SUP, SUP_HEAD);
  getExportFolder_();
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processExports').timeBased().everyMinutes(CHECK_MINUTES).create();
  PropertiesService.getScriptProperties().setProperty('trigger_v', TRIGGER_V);
  ['TOKEN', 'VIEW_KEY'].forEach(function (k) {
    if (!secret_(k)) console.warn('스크립트 속성에 ' + k + ' 가 없습니다. 프로젝트 설정 → 스크립트 속성에서 추가하세요.');
  });
  console.log('설정 완료: ' + CHECK_MINUTES + '분마다 내보내기 폴더를 확인합니다.');
}

/* 자동 확인 주기가 바뀌었으면 예약을 알아서 새로 맞춤 (setup을 다시 안 해도 되게) */
const CHECK_MINUTES = 1;          // 1·5·10·15·30 중 하나 (구글이 허용하는 최소 간격이 1분)
const TRIGGER_V = 'every' + CHECK_MINUTES + 'min';
function ensureTriggers_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('trigger_v') === TRIGGER_V) return;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processExports' || t.getHandlerFunction() === 'updateFromGitHub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processExports').timeBased().everyMinutes(CHECK_MINUTES).create();
  props.setProperty('trigger_v', TRIGGER_V);
  console.log('자동 확인 주기를 ' + CHECK_MINUTES + '분으로 바꿨습니다.');
}

/* 설치 확인용: 가짜 공지로 AI 추출만 해보고 저장은 하지 않음 */
function testSample() {
  const msgs = [{ i: 0, room: CONFIG.ROOMS[0], sender: '조교', time: fmt_(new Date()),
    text: '[공지] 다음 주 화요일 오후 1시 보존 실습 있습니다. 러버댐 키트 지참, 실습 보고서는 이번 주 금요일 18시까지 이클래스 제출 바랍니다. 그리고 엔도 실습은 앞으로 매번 K-file 세트와 확대경 챙겨오세요.' }];
  console.log(JSON.stringify(callLLM_(buildPrompt_([], msgs, [])), null, 2));
}

/* ---------- 관리자 편집 (목록 페이지에서 호출) ---------- */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  const admin = secret_('TOKEN');
  if (!admin || body.token !== admin) return json_({ ok: false, error: 'token' });
  if (body.type === 'sync') {           // 관리자 페이지의 '지금 정리' 버튼
    try {
      const r = processExports();
      return json_({ ok: true, busy: !r, log: r || ['다른 정리가 진행 중입니다. 잠시 후 새로고침해 보세요.'] });
    } catch (err) {
      console.error(err);
      return json_({ ok: false, error: '정리 중 오류: ' + err });   // 오류도 JSON으로 돌려줘야 페이지가 읽을 수 있음
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = adminOp_(body);
    CacheService.getScriptCache().remove('page');
    return json_(r);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function adminOp_(b) {
  if (/^s(add|edit|delete)$/.test(b.type)) return supplyOp_(b);
  const sh = sheet_(SH_TASK);
  const now = fmt_(new Date());
  const title = String(b.title || '').trim().slice(0, 60);
  const due = normDue_(b.due);
  const detail = String(b.detail || '').trim().slice(0, 500);
  const room = String(b.room || '').trim().slice(0, 40);

  if (b.type === 'add') {
    if (!title || !due) return { ok: false, error: '제목과 날짜는 꼭 필요합니다' };
    sh.appendRow([newId_(), title, due, room, detail, 'open', now, now, 'manual', '']);
    return { ok: true };
  }

  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(b.id)) { row = i; break; }
  if (row === -1) return { ok: false, error: '항목을 찾을 수 없습니다. 새로고침해 주세요' };
  const t = rowToTask_(vals[row]);

  if (b.type === 'edit') {
    if (!title || !due) return { ok: false, error: '제목과 날짜는 꼭 필요합니다' };
    if (t.source === 'manual') {
      sh.getRange(row + 1, 2, 1, 4).setValues([[title, due, room, detail]]);
      sh.getRange(row + 1, 8).setValue(now);
    } else {
      // 자동 항목을 고치면: 원래 것은 숨기고, 고친 내용은 '직접 입력'으로 따로 보관
      hide_(t);
      sh.getRange(row + 1, 6).setValue('replaced');
      sh.appendRow([newId_(), title, due, room, detail, 'open', now, now, 'manual', t.src]);
    }
    return { ok: true };
  }

  if (b.type === 'delete') {
    if (t.source !== 'manual') hide_(t);
    sh.getRange(row + 1, 6).setValue('deleted');
    sh.getRange(row + 1, 8).setValue(now);
    return { ok: true };
  }
  return { ok: false, error: 'type' };
}

/* 준비물 편집: 날짜를 비우면 '매 실습' 준비물 */
function supplyOp_(b) {
  const sh = sheet_(SH_SUP);
  const now = fmt_(new Date());
  const subject = String(b.subject || '').trim();
  const items = String(b.items || '').trim().slice(0, 300);
  const detail = String(b.detail || '').trim().slice(0, 300);
  const date = b.date ? normDue_(b.date).slice(0, 10) : '';
  if (b.type !== 'sdelete') {
    if (allSubjects_().indexOf(subject) === -1) return { ok: false, error: '과목을 골라주세요' };
    if (!items) return { ok: false, error: '준비물을 입력하세요' };
    if (b.date && !date) return { ok: false, error: '날짜 형식이 맞지 않습니다' };
  }
  if (b.type === 'sadd') {
    sh.appendRow([newId_(), subject, date, items, detail, '', 'open', now, 'manual', '']);
    return { ok: true };
  }
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(b.id)) { row = i; break; }
  if (row === -1) return { ok: false, error: '항목을 찾을 수 없습니다. 새로고침해 주세요' };
  const s = rowToSupply_(vals[row]);
  if (b.type === 'sedit') {
    if (s.source === 'manual') {
      sh.getRange(row + 1, 2, 1, 4).setValues([[subject, date, items, detail]]);
      sh.getRange(row + 1, 8).setValue(now);
    } else {
      // 자동 항목을 고치면: 원래 것은 숨기고, 고친 내용은 '직접 입력'으로 따로 보관
      hideSupply_(s);
      sh.getRange(row + 1, 7).setValue('replaced');
      sh.appendRow([newId_(), subject, date, items, detail, s.room, 'open', now, 'manual', s.src]);
    }
    return { ok: true };
  }
  if (s.source !== 'manual') hideSupply_(s);
  sh.getRange(row + 1, 7).setValue('deleted');
  sh.getRange(row + 1, 8).setValue(now);
  return { ok: true };
}

function supplyKey_(s) { return 'S:' + s.subject + ':' + titleKey_(s.items); }
function hideSupply_(s) {
  sheet_(SH_HIDDEN).appendRow([s.src || '', supplyKey_(s), s.date || 'always', s.subject + ' ' + s.items, fmt_(new Date())]);
}
function isSupplyHidden_(s, hidden) {
  const day = s.date || 'always';
  const k = supplyKey_(s);
  return hidden.some(function (h) {
    if (h.dueDate !== day) return false;
    return (h.src && s.src && h.src === s.src && h.titleKey.indexOf('S:' + s.subject + ':') === 0) || h.titleKey === k;
  });
}
function allSubjects_() {
  const out = [];
  Object.keys(CONFIG.LABS).forEach(function (d) {
    CONFIG.LABS[d].forEach(function (x) { if (out.indexOf(x) === -1) out.push(x); });
  });
  return out;
}

function hide_(t) {
  sheet_(SH_HIDDEN).appendRow([t.src || '', titleKey_(t.title), String(t.due).slice(0, 10), t.title, fmt_(new Date())]);
}

/* ---------- 목록 데이터 ---------- */
function doGet(e) {
  const k = e.parameter.k || '';
  const admin = secret_('TOKEN'), view = secret_('VIEW_KEY');
  if (!k || (k !== view && k !== admin)) return json_({ ok: false, error: 'key' });
  try { ensureTriggers_(); } catch (err) { console.warn(err); }
  const cache = CacheService.getScriptCache();
  let d = cache.get('page');
  if (!d) {                             // 여러 명이 동시에 열어도 시트는 1분에 한 번만 읽음
    const o = pageData_(); o.ok = true;
    d = JSON.stringify(o);
    cache.put('page', d, 60);
  }
  if (admin && k === admin) {           // 관리자에게만 편집 버튼 표시
    const o = JSON.parse(d); o.admin = true;
    return json_(o);
  }
  return ContentService.createTextOutput(d).setMimeType(ContentService.MimeType.JSON);
}

function pageData_() {
  const now = Date.now();
  const hidden = readHidden_();
  const tasks = readTasks_().filter(function (t) {
    if (t.status !== 'open') return false;
    const end = dueEnd_(t.due);
    if (!end || end.getTime() < now) return false;          // 마감이 명확하고 아직 안 지난 것만
    return t.source === 'manual' || !isHidden_(t, hidden);
  }).sort(function (a, b) { return dueEnd_(a.due) - dueEnd_(b.due); })
    .map(function (t) {
      return { id: t.id, title: t.title, due: t.due, room: t.room, detail: t.detail, manual: t.source === 'manual' };
    });
  const today = fmt_(new Date()).slice(0, 10);
  const supplies = readSupplies_().filter(function (x) {
    if (x.status !== 'open' || !x.items) return false;
    if (x.date && x.date < today) return false;            // 지난 날짜 준비물은 숨김
    return x.source === 'manual' || !isSupplyHidden_(x, hidden);
  }).map(function (x) {
    return { id: x.id, subject: x.subject, date: x.date, items: x.items, detail: x.detail, manual: x.source === 'manual' };
  });
  return {
    tasks: tasks,
    supplies: CONFIG.SUPPLIES ? supplies : [],
    labs: CONFIG.SUPPLIES ? CONFIG.LABS : null,
    rooms: CONFIG.ROOMS,
    lastSync: PropertiesService.getScriptProperties().getProperty('lastSync') || '',
    now: fmt_(new Date())
  };
}

function isHidden_(t, hidden) {
  const day = String(t.due).slice(0, 10);
  const tk = titleKey_(t.title);
  return hidden.some(function (h) {
    if (h.dueDate !== day) return false;
    return (h.src && t.src && h.src === t.src) || (h.titleKey && h.titleKey === tk);
  });
}

/* ---------- 내보내기 파일 → 목록 다시 만들기 ---------- */
function processExports() {
  try { ensureTriggers_(); } catch (err) { console.warn(err); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return null;  // 이미 처리 중이면 다음 기회에
  const summary = [];
  const note = function (m) { console.log(m); summary.push(m); };
  try {
    const t0 = Date.now();
    const folder = getExportFolder_();
    const doneFolder = getSubFolder_(folder, '처리완료');
    const files = [];
    const it = folder.getFiles();
    while (it.hasNext()) files.push(it.next());
    files.sort(function (a, b) { return a.getDateCreated() - b.getDateCreated(); });   // 오래된 것부터

    if (!files.length) note('폴더에 새 내보내기 파일이 없습니다.');
    files.forEach(function (f) {
      if (Date.now() - t0 > 4 * 60 * 1000) { note('시간 부족 → 다음 확인 때 처리: ' + f.getName()); return; }
      const name = f.getName();
      const mime = f.getMimeType();
      if (/^application\/vnd\.google-apps/.test(mime) || /^(image|video|audio)\//.test(mime)) {
        console.log('건너뜀(카톡 내보내기 아님): ' + name + ' / ' + mime);
        return;
      }
      let texts = [];
      try { texts = extractTexts_(f.getBlob()); }
      catch (err) { console.warn('파일을 읽지 못함: ' + name + ' / ' + err); return; }

      let found = false, ok = true;
      texts.forEach(function (raw) {
        const parsed = parseKakaoExport_(raw, name);
        if (!parsed.room) {
          console.warn('방 이름을 못 찾음: ' + name + '\n첫 줄: ' + raw.split(/\r?\n/)[0] +
                       '\n→ CONFIG.ROOMS 이름과 비교해 보세요: ' + CONFIG.ROOMS.join(', '));
          return;
        }
        if (!parsed.messages.length) {
          console.warn('메시지를 하나도 못 읽음: ' + name + '\n' + raw.split(/\r?\n/).slice(1, 6).join('\n'));
          return;
        }
        found = true;
        const r = rebuildRoom_(parsed.room, parsed.messages);
        ok = !!r && ok;
        note(r ? r : parsed.room + ': AI 서버가 붐벼서 실패 → 기존 목록 유지, ' + CHECK_MINUTES + '분 뒤 자동 재시도');
      });
      if (found && ok) f.moveTo(doneFolder);
      else if (!found) note('읽을 수 없는 파일: ' + name);
    });
  } finally {
    lock.releaseLock();
  }
  return summary;
}

/* 한 방의 자동 항목을 최근 공지로 통째로 다시 만듦. AI가 끝까지 성공했을 때만 교체 */
function rebuildRoom_(room, messages) {
  const since = Date.now() - CONFIG.REBUILD_DAYS * 864e5;
  const msgs = messages.filter(function (m) {
    if (m.time.getTime() < since) return false;
    const s = m.text.trim();
    if (s.length < 2) return false;
    if (/^(삭제된 메시지입니다\.?|메시지가 삭제되었습니다\.?)$/.test(s)) return false;
    if (/^(사진|동영상|이모티콘|파일|사진 \d+장|음성메시지)$/.test(s)) return false;
    return true;
  }).map(function (m) {
    return { room: room, sender: m.sender, time: fmt_(m.time), text: m.text,
             key: [room, fmt_(m.time), m.sender, norm_(m.text).slice(0, 20)].join('|') };
  });

  let list = [];
  let slist = [];
  for (let i = 0; i < msgs.length; i += CONFIG.BATCH) {
    const batch = msgs.slice(i, i + CONFIG.BATCH).map(function (m, j) { m.i = j; return m; });
    let result;
    try {
      result = callLLM_(buildPrompt_(list.map(taskForPrompt_), batch, slist.map(supplyForPrompt_)));
    } catch (err) {
      console.error(room + ' 처리 중 오류: ' + err);
      return false;                     // 기존 목록은 그대로 유지
    }
    list = applyOpsMem_(list, result.ops || [], batch, room);
    if (CONFIG.SUPPLIES) slist = applySupplyOpsMem_(slist, result.ops || [], batch, room);
  }

  writeAutoTasks_(room, list.filter(function (t) { return t.status === 'open' && t.due; }));
  if (CONFIG.SUPPLIES) writeAutoSupplies_(room, slist.filter(function (x) { return x.status === 'open' && x.items; }));
  PropertiesService.getScriptProperties().setProperty('lastSync', fmt_(new Date()));
  CacheService.getScriptCache().remove('page');
  return (room + ': 최근 ' + CONFIG.REBUILD_DAYS + '일 공지 ' + msgs.length + '개 → 일정 ' +
              list.filter(function (t) { return t.status === 'open' && t.due; }).length + '개' +
              (CONFIG.SUPPLIES ? ', 준비물 ' + slist.filter(function (x) { return x.status === 'open'; }).length + '개' : '') + '로 다시 만듦');
}

function taskForPrompt_(t) { return { id: t.id, title: t.title, due: t.due || null, detail: t.detail }; }

function applyOpsMem_(list, ops, batch, room) {
  const now = fmt_(new Date());
  ops.forEach(function (op) {
    if (op.op === 'add' && op.title) {
      const due = normDue_(op.due);
      if (!due) return;
      const src = batch[Number(op.src)] ? batch[Number(op.src)].key : '';
      list.push({ id: newId_(), title: String(op.title).slice(0, 60), due: due, room: room,
                  detail: String(op.detail || ''), status: 'open', created: now, updated: now,
                  source: 'auto', src: src });
    } else if (op.op === 'update') {
      const t = list.filter(function (x) { return x.id === op.id; })[0];
      if (!t) return;
      if (op.title) t.title = String(op.title).slice(0, 60);
      if (op.due) { const d = normDue_(op.due); if (d) t.due = d; }
      if (op.detail) t.detail = String(op.detail);
    } else if (op.op === 'cancel') {
      const t = list.filter(function (x) { return x.id === op.id; })[0];
      if (t) t.status = 'cancelled';
    }
  });
  return list;
}

function supplyForPrompt_(x) { return { id: x.id, subject: x.subject, date: x.date || null, items: x.items, detail: x.detail }; }

function applySupplyOpsMem_(slist, ops, batch, room) {
  const now = fmt_(new Date());
  const subjects = allSubjects_();
  ops.forEach(function (op) {
    if (op.op === 'supply' && op.items && subjects.indexOf(op.subject) !== -1) {
      const date = op.date ? normDue_(op.date).slice(0, 10) : '';
      if (op.date && !date) return;
      const src = batch[Number(op.src)] ? batch[Number(op.src)].key : '';
      slist.push({ id: newId_(), subject: op.subject, date: date, items: String(op.items).slice(0, 300),
                   detail: String(op.detail || '').slice(0, 300), room: room, status: 'open', updated: now,
                   source: 'auto', src: src });
    } else if (op.op === 'supply_update') {
      const x = slist.filter(function (y) { return y.id === op.id; })[0];
      if (!x) return;
      if (op.items) x.items = String(op.items).slice(0, 300);
      if (op.detail) x.detail = String(op.detail).slice(0, 300);
      if (op.date !== undefined) { const d = op.date ? normDue_(op.date).slice(0, 10) : ''; if (!op.date || d) x.date = d; }
    } else if (op.op === 'supply_cancel') {
      const x = slist.filter(function (y) { return y.id === op.id; })[0];
      if (x) x.status = 'cancelled';
    }
  });
  return slist;
}

function writeAutoSupplies_(room, newRows) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH_SUP);
    const vals = sh.getDataRange().getValues();
    const keep = [];
    for (let i = 1; i < vals.length; i++) {
      if (!vals[i][0]) continue;
      const x = rowToSupply_(vals[i]);
      if (x.source !== 'manual' && x.room === room) continue;      // 이 방의 자동 항목은 교체
      keep.push(padTo_(vals[i], SUP_HEAD.length));
    }
    newRows.forEach(function (x) {
      keep.push([x.id, x.subject, x.date, x.items, x.detail, x.room, 'open', x.updated, 'auto', x.src]);
    });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, SUP_HEAD.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, SUP_HEAD.length).setValues(keep);
  } finally {
    lock.releaseLock();
  }
}

function writeAutoTasks_(room, newTasks) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH_TASK);
    const vals = sh.getDataRange().getValues();
    const keep = [];
    for (let i = 1; i < vals.length; i++) {
      if (!vals[i][0]) continue;
      const t = rowToTask_(vals[i]);
      if (t.source !== 'manual' && t.room === room) continue;      // 이 방의 자동 항목은 교체
      keep.push(pad_(vals[i]));
    }
    newTasks.forEach(function (t) {
      keep.push([t.id, t.title, t.due, t.room, t.detail, 'open', t.created, t.updated, 'auto', t.src]);
    });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TASK_HEAD.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, TASK_HEAD.length).setValues(keep);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- AI ---------- */
function buildPrompt_(known, msgs, knownSup) {
  const on = CONFIG.SUPPLIES;
  const supRules = on ? [
    '## 실습 준비물 규칙',
    '- 정규 실습 시간표: ' + labsText_(),
    '- subject는 반드시 이 중 하나: ' + allSubjects_().join(', '),
    '- 줄임말도 알맞은 과목으로 연결하라 (예: 보존·수복→보존수복학, 엔도·근관→근관치료학, 크라운·브릿지·금관→금관가공의치학, 파샬·RPD·국소→국소의치학, 덴처·총의치·CD→총의치학, 치주→치주학, 교정→교정학, 소아→소아치과학, 예방→예방치과학). 어느 과목인지 알 수 없으면 supply 하지 마라.',
    '- 실습에 가져와야 할 물품·기구·재료·복장·출력물이 공지되면 supply 하라.',
    '- 특정 날짜 실습에만 해당하면 date에 그 날짜("YYYY-MM-DD"). "다음 실습", "이번 주 실습" 같은 표현은 위 시간표와 메시지 time으로 실제 날짜를 계산하라.',
    '- "매 실습", "앞으로", "학기 내내"처럼 계속 챙겨야 하면 date는 null.',
    '- items는 쉼표로 구분한 물품 목록(예: "러버댐 키트, 핸드피스, 실습 가운"). detail에는 수량·주의사항을 짧게.',
    '- 이미 있는 준비물이 바뀌면 supply_update(id 사용), 필요 없어졌다고 하면 supply_cancel.',
    '- 준비물은 일정의 detail에도 적어도 되지만, 실습 준비물이면 supply도 반드시 함께 만들어라.',
    ''
  ] : [];
  return [
    '너는 치의학전문대학원 학급의 공지 비서다. 카카오톡 공지방 메시지에서 학생들이 챙겨야 할 ' + (on ? '"마감·일정"과 "실습 준비물"을' : '"마감·일정"을') + ' 뽑는다.',
    '현재 시각(KST): ' + fmt_(new Date()),
    '',
    '## 규칙',
    '- 대상: 제출·과제·신청·설문·서명·납부 마감, 실습/시험/특강/행사 일정, 출결 관련 지시.',
    '- 날짜가 명확히 제시된 것만 add 하라. "추후 공지", "곧", "빠른 시일 내" 처럼 날짜를 특정할 수 없으면 add 하지 마라.',
    '- 제출·신청류는 마감 시각을, 실습·시험·행사류는 시작 시각을 due로 한다. "오후 6시"→18:00, "자정까지"→23:59, "정오"→12:00.',
    '- 준비물·지참물·장소·대상(조, 번호)은 따로 add 하지 말고 해당 일정의 detail에 넣어라.',
    '- 무시: 인사, "넵/확인했습니다" 같은 답장, 행동이 필요 없는 단순 안내.',
    '- "내일", "다음 주 화요일", "이번 주 금요일까지" 같은 상대 날짜는 해당 메시지의 time 기준으로 절대 날짜로 바꿔라.',
    '- due 형식: 시각이 있으면 "YYYY-MM-DD HH:mm"(24시간제), 날짜만 있으면 "YYYY-MM-DD".',
    '- 기존 목록에 이미 있는 일을 다시 언급·변경·연기하면 add 하지 말고 update(id 사용). 취소되면 cancel.',
    '- 한 메시지에 일정이 여러 개면 각각 add (예: 실습 일정과 보고서 제출은 별개).',
    '- title: 25자 이내 명사형(예: "보존 실습 보고서 제출"). detail: 한두 줄, 없으면 "".',
    '- add 할 때 src에는 그 내용이 나온 메시지의 i 값을 넣어라.',
    '',
    ...supRules,
    '## 출력 (JSON만, 설명 금지)',
    '{"ops":[{"op":"add","src":0,"title":"","due":"","detail":""},{"op":"update","id":"","title":"","due":"","detail":""},{"op":"cancel","id":""}' +
      (on ? ',\n {"op":"supply","src":0,"subject":"","date":null,"items":"","detail":""},{"op":"supply_update","id":"","date":null,"items":"","detail":""},{"op":"supply_cancel","id":""}' : '') + ']}',
    '없으면 {"ops":[]}',
    '',
    '## 기존 일정 목록',
    JSON.stringify(known),
    '',
    on ? '## 기존 준비물 목록\n' + JSON.stringify(knownSup || []) : '',
    '',
    '## 메시지 (시간순)',
    JSON.stringify(msgs.map(function (m) { return { i: m.i, sender: m.sender, time: m.time, text: m.text }; }))
  ].join('\n');
}

function callLLM_(prompt) {
  const props = PropertiesService.getScriptProperties();
  let text = '';
  if (CONFIG.LLM === 'claude') {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': props.getProperty('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: CONFIG.CLAUDE_MODEL, max_tokens: 8000,
                                messages: [{ role: 'user', content: prompt }] })
    });
    if (res.getResponseCode() !== 200) throw new Error('Claude ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
    text = JSON.parse(res.getContentText()).content.map(function (c) { return c.text || ''; }).join('');
  } else {
    // 서버가 붐비면(503/429/500) 잠시 쉬었다 재시도하고, 그래도 안 되면 다른 모델로 넘어감
    const models = [CONFIG.GEMINI_MODEL].concat(CONFIG.GEMINI_FALLBACKS || [])
      .filter(function (m, i, a) { return m && a.indexOf(m) === i; });
    let data = null, lastErr = '';
    for (let mi = 0; mi < models.length && !data; mi++) {
      for (let attempt = 0; attempt < 2 && !data; attempt++) {
        const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[mi] + ':generateContent', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { 'x-goog-api-key': props.getProperty('GEMINI_API_KEY') },
          payload: JSON.stringify({
            contents: [{ role: 'user', parts: [{ text: prompt }] }],
            generationConfig: { responseMimeType: 'application/json' }
          })
        });
        const code = res.getResponseCode();
        if (code === 200) { data = JSON.parse(res.getContentText()); break; }
        lastErr = 'Gemini ' + models[mi] + ' ' + code + ': ' + res.getContentText().slice(0, 200);
        if (code === 404) break;                                   // 없는 모델 → 다음 모델
        if ([429, 500, 503].indexOf(code) === -1) throw new Error(lastErr);   // 키 오류 등은 바로 중단
        console.log(lastErr.slice(0, 80) + ' → 재시도');
        Utilities.sleep(attempt === 0 ? 4000 : 10000);
      }
    }
    if (!data) { const e = new Error(lastErr); e.busy = true; throw e; }
    const cand = (data.candidates || [])[0];
    const parts = (cand && cand.content && cand.content.parts) || [];
    text = parts.filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
  }
  return parseJson_(text);
}

function parseJson_(text) {
  const s = text.indexOf('{');
  const e = text.lastIndexOf('}');
  if (s === -1 || e === -1) throw new Error('JSON 없음: ' + text.slice(0, 200));
  return JSON.parse(text.slice(s, e + 1));
}

/* ---------- 파일 읽기 ---------- */
/* 파일 → 텍스트 목록 (txt, zip, 안드로이드 공유로 생긴 메일 형식(message/rfc822) 모두) */
function extractTexts_(blob) {
  const bytes = blob.getBytes();
  if (bytes.length > 1 && bytes[0] === 0x50 && bytes[1] === 0x4B) {   // zip
    return zipTexts_(blob);
  }
  const raw = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
  if (/카카오톡 대화|KakaoTalk/i.test(raw.split(/\r?\n/).slice(0, 3).join('\n'))) return [raw];
  if (/^[A-Za-z-]+:[^\n]*\r?\n/.test(raw) && /\r?\n\r?\n/.test(raw)) {   // 메일 형식
    const out = mimeTexts_(raw);
    if (out.length) return out;
  }
  return [raw];
}

function zipTexts_(blob) {
  return Utilities.unzip(blob)
    .filter(function (b) { return /\.txt$/i.test(b.getName()); })
    .map(function (b) { return b.getDataAsString('UTF-8').replace(/^\uFEFF/, ''); });
}

function mimeTexts_(raw) {
  const sep = /\r?\n\r?\n/.exec(raw);
  const headText = raw.slice(0, sep.index).replace(/\r?\n[ \t]+/g, ' ');
  const body = raw.slice(sep.index + sep[0].length);
  const hdr = {};
  headText.split(/\r?\n/).forEach(function (l) {
    const i = l.indexOf(':');
    if (i > 0) hdr[l.slice(0, i).trim().toLowerCase()] = l.slice(i + 1).trim();
  });
  const ctype = hdr['content-type'] || 'text/plain';
  const enc = (hdr['content-transfer-encoding'] || '').toLowerCase();

  const bm = /boundary="?([^";]+)"?/i.exec(ctype);
  if (/^multipart\//i.test(ctype) && bm) {
    let out = [];
    body.split('--' + bm[1]).slice(1).forEach(function (part) {
      if (/^--/.test(part)) return;                       // 끝 표시
      part = part.replace(/^\r?\n/, '');
      if (!/\r?\n\r?\n/.test(part)) return;
      out = out.concat(mimeTexts_(part));
    });
    return out;
  }

  if (enc === 'base64') {
    const b = Utilities.newBlob(Utilities.base64Decode(body.replace(/\s+/g, '')));
    const by = b.getBytes();
    if (by.length > 1 && by[0] === 0x50 && by[1] === 0x4B) return zipTexts_(b);
    if (/^(text\/|application\/octet-stream)/i.test(ctype) || /\.txt/i.test(ctype + (hdr['content-disposition'] || ''))) {
      return [b.getDataAsString('UTF-8').replace(/^\uFEFF/, '')];
    }
    return [];
  }
  if (enc === 'quoted-printable') return [qpDecode_(body)];
  if (/^text\//i.test(ctype) || !hdr['content-type']) return [body];
  return [];
}

function qpDecode_(s) {
  s = s.replace(/=\r?\n/g, '').replace(/%/g, '%25').replace(/=([0-9A-Fa-f]{2})/g, '%$1');
  try { return decodeURIComponent(s); } catch (e) { return s; }
}

/* 카톡 내보내기 파싱 (모바일·PC 형식 모두) */
function parseKakaoExport_(text, fileName) {
  text = String(text).replace(/^\uFEFF/, '');
  const lines = text.split(/\r?\n/);
  const head = lines.slice(0, 3).join('\n');
  let room = '';
  const nsp = function (s) { return String(s || '').replace(/[\s_]+/g, ''); };   // 띄어쓰기·밑줄 차이 무시
  CONFIG.ROOMS.forEach(function (r) {
    if (!room && (nsp(head).indexOf(nsp(r)) !== -1 || nsp(fileName).indexOf(nsp(r)) !== -1)) room = r;
  });

  const reMobile = /^(\d{4})[년.]\s*(\d{1,2})[월.]\s*(\d{1,2})[일.]?\s*(오전|오후)\s*(\d{1,2}):(\d{2}),\s*(.+?)\s:\s([\s\S]*)$/;
  const reMobileSys = /^(\d{4})[년.]\s*(\d{1,2})[월.]\s*(\d{1,2})[일.]?\s*(오전|오후)\s*\d{1,2}:\d{2}(,|$)/;
  const reDateHead = /^(\d{4})[년.]\s*(\d{1,2})[월.]\s*(\d{1,2})[일.]?\s*([월화수목금토일](요일)?)?\s*$/;
  const rePcDate = /^-{3,}\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/;
  const rePcMsg = /^\[(.+?)\]\s\[(오전|오후)\s*(\d{1,2}):(\d{2})\]\s([\s\S]*)$/;

  const out = [];
  let cur = null;
  let pcDate = null;
  function flush() { if (cur) { cur.text = cur.text.trim(); out.push(cur); } cur = null; }

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    let m;
    if ((m = reMobile.exec(line))) {
      flush();
      cur = { sender: m[7], time: kst_(m[1], m[2], m[3], hour24_(m[4], m[5]), m[6]), text: m[8] };
    } else if ((m = rePcDate.exec(line))) {
      flush(); pcDate = [m[1], m[2], m[3]];
    } else if (pcDate && (m = rePcMsg.exec(line))) {
      flush();
      cur = { sender: m[1], time: kst_(pcDate[0], pcDate[1], pcDate[2], hour24_(m[2], m[3]), m[4]), text: m[5] };
    } else if (reDateHead.test(line) || reMobileSys.test(line)) {
      flush();
    } else if (cur) {
      cur.text += '\n' + line;
    }
  }
  flush();
  return { room: room, messages: out };
}

/* ---------- 유틸 ---------- */
function cell_(v) { return v instanceof Date ? fmt_(v) : String(v == null ? '' : v); }

function rowToTask_(r) {
  return { id: String(r[0]), title: String(r[1]), due: normDue_(cell_(r[2])), room: String(r[3] || ''),
           detail: String(r[4] || ''), status: String(r[5] || 'open'), created: cell_(r[6]), updated: cell_(r[7]),
           source: String(r[8] || 'auto'), src: String(r[9] || '') };
}

function readTasks_() {
  const vals = sheet_(SH_TASK).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) if (vals[i][0]) out.push(rowToTask_(vals[i]));
  return out;
}

function labsText_() {
  const wd = ['일', '월', '화', '수', '목', '금', '토'];
  return Object.keys(CONFIG.LABS).map(function (d) { return wd[d] + '요일: ' + CONFIG.LABS[d].join(', '); }).join(' / ');
}

function rowToSupply_(r) {
  return { id: String(r[0]), subject: String(r[1] || ''), date: cell_(r[2]).slice(0, 10), items: String(r[3] || ''),
           detail: String(r[4] || ''), room: String(r[5] || ''), status: String(r[6] || 'open'), updated: cell_(r[7]),
           source: String(r[8] || 'auto'), src: String(r[9] || '') };
}

function readSupplies_() {
  const vals = sheet_(SH_SUP).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) if (vals[i][0]) out.push(rowToSupply_(vals[i]));
  return out;
}

function padTo_(r, n) {
  const out = [];
  for (let i = 0; i < n; i++) out.push(r[i] instanceof Date ? fmt_(r[i]) : (r[i] == null ? '' : r[i]));
  return out;
}

function readHidden_() {
  const vals = sheet_(SH_HIDDEN).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][1] && !vals[i][0]) continue;
    out.push({ src: String(vals[i][0] || ''), titleKey: String(vals[i][1] || ''), dueDate: cell_(vals[i][2]).slice(0, 10) });
  }
  return out;
}

function pad_(r) {
  const out = [];
  for (let i = 0; i < TASK_HEAD.length; i++) out.push(r[i] instanceof Date ? fmt_(r[i]) : (r[i] == null ? '' : r[i]));
  return out;
}

function titleKey_(s) { return String(s || '').toLowerCase().replace(/[^0-9a-z가-힣]/g, ''); }
function newId_() { return Utilities.getUuid().slice(0, 8); }

/* 여러 날짜 표기를 "YYYY-MM-DD" 또는 "YYYY-MM-DD HH:mm"으로 */
function normDue_(d) {
  if (!d) return '';
  if (d instanceof Date) return fmt_(d);
  const s = String(d).trim();
  const m = /(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})일?(?:[T\s]+(?:(오전|오후)\s*)?(\d{1,2})(?::(\d{2})|시(?:\s*(\d{1,2})분?)?))?/.exec(s);
  if (!m) return '';
  const p = function (n) { return ('0' + n).slice(-2); };
  const date = m[1] + '-' + p(m[2]) + '-' + p(m[3]);
  if (m[5] === undefined) return date;
  let h = Number(m[5]);
  if (m[4]) h = hour24_(m[4], h);
  const mi = Number(m[6] || m[7] || 0);
  if (h > 23 || mi > 59) return date;
  return date + ' ' + p(h) + ':' + p(mi);
}

/* 마감 시점: 시각 있으면 그 시각, 날짜만 있으면 그날 23:59 */
function dueEnd_(due) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?$/.exec(due || '');
  if (!m) return null;
  return m[4] ? kst_(m[1], m[2], m[3], m[4], m[5]) : kst_(m[1], m[2], m[3], 23, 59);
}

function norm_(s) { return String(s).replace(/\s+/g, '').replace(/[.…]+$/, ''); }
function hour24_(ampm, h) { h = Number(h); if (ampm === '오전') return h === 12 ? 0 : h; return h === 12 ? 12 : h + 12; }
function kst_(y, mo, d, h, mi) { return new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h) - 9, Number(mi))); }
function fmt_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm'); }
function secret_(name) { return String(PropertiesService.getScriptProperties().getProperty(name) || '').trim(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ensureSheet_(ss, name, HEADS[name]);
}

function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');   // 날짜 자동변환 방지
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function getExportFolder_() {
  const it = DriveApp.getFoldersByName(CONFIG.EXPORT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.EXPORT_FOLDER);
}
function getSubFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* 로더(Code.gs)가 바깥에서 부를 수 있게 내보내는 함수들 */
return { doGet: doGet, doPost: doPost, setup: setup, processExports: processExports, testSample: testSample };
