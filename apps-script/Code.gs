/******************************************************
 * 디데이 25 — 로더 (Apps Script에는 이 파일만 붙여넣어 두면 됩니다)
 *
 * 실제 기능 코드는 GitHub(Koderjoon/notice-todo)의 apps-script/app.js에 있고,
 * 이 로더가 실행될 때마다 최신 코드를 가져와 실행합니다(5분 캐시).
 * 그래서 GitHub의 코드가 바뀌면 다시 배포하지 않아도 5분 안에 반영됩니다.
 *
 * GitHub에 연결이 안 되거나 새 코드에 문제가 있으면,
 * 마지막으로 정상 동작했던 코드를 대신 씁니다.
 ******************************************************/

const APP_URL = 'https://raw.githubusercontent.com/Koderjoon/notice-todo/main/apps-script/app.js';
const APP_CACHE_SECONDS = 300;

function doGet(e) { return app_().doGet(e); }
function doPost(e) { return app_().doPost(e); }
function setup() { return app_().setup(); }
function processExports() { return app_().processExports(); }
function testSample() { return app_().testSample(); }

/* 지금 바로 최신 코드로 바꾸고 싶을 때 실행 */
function reloadApp() {
  CacheService.getScriptCache().remove('appsrc');
  app_();
  console.log('최신 코드를 불러왔습니다. 버전: ' + (PropertiesService.getScriptProperties().getProperty('app_hash') || '').slice(0, 8));
}

let APP_ = null;
function app_() {
  if (APP_) return APP_;
  const cache = CacheService.getScriptCache();
  let src = cache.get('appsrc');
  if (!src) {
    try {
      const r = UrlFetchApp.fetch(APP_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
      src = r.getContentText('UTF-8');
      APP_ = compile_(src);                 // 문법 오류면 여기서 실패 → 아래 예비 코드로
      saveGood_(src);
      try { cache.put('appsrc', src, APP_CACHE_SECONDS); } catch (e) {}
      return APP_;
    } catch (err) {
      console.warn('GitHub 코드를 쓸 수 없어 마지막 정상 코드를 사용합니다: ' + err);
      src = loadGood_();
      if (!src) throw new Error('불러올 코드가 없습니다. 인터넷 연결이나 GitHub 저장소를 확인하세요. (' + err + ')');
    }
  }
  APP_ = compile_(src);
  return APP_;
}

function compile_(src) {
  const app = new Function(src)();
  if (!app || typeof app.doGet !== 'function') throw new Error('app.js 형식이 올바르지 않습니다');
  return app;
}

/* 마지막 정상 코드를 스크립트 속성에 나눠 보관 (속성 하나당 용량 제한 때문) */
function saveGood_(src) {
  const props = PropertiesService.getScriptProperties();
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, src, Utilities.Charset.UTF_8));
  if (props.getProperty('app_hash') === hash) return;
  const size = 2500, parts = {};
  const n = Math.ceil(src.length / size);
  for (let i = 0; i < n; i++) parts['app_part_' + i] = src.slice(i * size, (i + 1) * size);
  parts.app_parts = String(n);
  parts.app_hash = hash;
  const old = Number(props.getProperty('app_parts') || 0);
  for (let i = n; i < old; i++) props.deleteProperty('app_part_' + i);
  props.setProperties(parts);
}

function loadGood_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('app_parts') || 0);
  let s = '';
  for (let i = 0; i < n; i++) s += props.getProperty('app_part_' + i) || '';
  return s;
}
