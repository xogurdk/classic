/**
 * 고전독서 박람회 기록 저장소 (구글 Apps Script 웹 앱)
 *
 * 기존에 쓰던 Apps Script에 넣는 법 (웹 앱 주소가 그대로라 index.html은 안 바꿔도 돼요)
 * 1. 지금 API_URL로 쓰고 있는 Apps Script 프로젝트를 열어요.
 * 2. Code.gs의 내용을 모두 지우고 이 파일 내용을 붙여 넣은 뒤 저장해요.
 *    (다른 .gs 파일이 있으면 doGet/doPost가 겹치지 않게 지워 주세요.)
 * 3. [배포] → [배포 관리] → 지금 배포 옆 연필(수정) → 버전: [새 버전] → [배포]
 *    ※ [새 배포]를 누르면 주소가 바뀌어요. 꼭 [배포 관리]에서 수정해 주세요.
 *
 * 기록은 시트의 '박람회_' 로 시작하는 탭에 저장돼요. 처음 쓸 때 자동으로 만들어지고,
 * 원래 있던 탭은 건드리지 않아요. 시트에 붙어 있지 않은 스크립트라면
 * '고전독서 박람회 기록' 시트를 내 드라이브에 새로 만들어 써요.
 * 선생님 PIN은 처음 1234예요. 선생님 화면에서 바꿔 주세요.
 */

var SHEETS = {
  books:     ["k", "canva", "web", "intro", "why", "at"],
  comments:  ["id", "book", "nick", "name", "text", "at", "ok", "dev"],
  reactions: ["who", "book", "kind", "at"],
  posts:     ["name", "b1", "b2", "reason", "feel", "at"],
  expects:   ["name", "book", "text", "at"]
};
var TAB = { books: "박람회_책자료", comments: "박람회_댓글", reactions: "박람회_반응", posts: "박람회_소감", expects: "박람회_기대되는부분" };
var NUM = { at: 1 }, BOOL = { ok: 1 };
var BOOK_FIELDS = { canva: 1, web: 1, intro: 1, why: 1 };
var BADGE_MIN = 40;

function doGet() {
  return ContentService.createTextOutput("고전독서 박람회 저장소가 동작 중이에요.");
}

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    out = handle(body);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 설정 ---------- */
function props() { return PropertiesService.getScriptProperties(); }
function pin() { return props().getProperty("fair_pin") || "1234"; }
function mode() { return props().getProperty("fair_mode") === "share" ? "share" : "class"; }
function timer() { try { return JSON.parse(props().getProperty("fair_timer") || "null"); } catch (e) { return null; } }

/* ---------- 시트 ---------- */
function book() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  var id = props().getProperty("fair_sheet");
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create("고전독서 박람회 기록");
  props().setProperty("fair_sheet", ss.getId());
  return ss;
}
var _ss = null;
function sheet(name) {
  _ss = _ss || book();
  var s = _ss.getSheetByName(TAB[name]);
  if (!s) { s = _ss.insertSheet(TAB[name]); s.appendRow(SHEETS[name]); s.setFrozenRows(1); }
  return s;
}
function toObj(name, row) {
  var o = {};
  SHEETS[name].forEach(function (k, i) {
    var v = row[i];
    if (NUM[k]) o[k] = Number(v) || 0;
    else if (BOOL[k]) o[k] = v === true || String(v).toUpperCase() === "TRUE";
    else o[k] = v == null ? "" : String(v);
  });
  return o;
}
function cell(v) {
  if (typeof v === "string" && /^[=+\-@]/.test(v)) return "'" + v; // 수식으로 읽히지 않게
  return v == null ? "" : v;
}
function toRow(name, o) { return SHEETS[name].map(function (k) { return cell(o[k]); }); }
function rows(name) {
  var v = sheet(name).getDataRange().getValues();
  v.shift();
  return v.filter(function (r) { return r.join("") !== ""; }).map(function (r) { return toObj(name, r); });
}
function upsert(name, match, obj) {
  var s = sheet(name), v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (match(toObj(name, v[i]))) { s.getRange(i + 1, 1, 1, SHEETS[name].length).setValues([toRow(name, obj)]); return; }
  }
  s.appendRow(toRow(name, obj));
}
function removeWhere(name, match) {
  var s = sheet(name), v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (match(toObj(name, v[i]))) s.deleteRow(i + 1);
}
function clearSheet(name) {
  var s = sheet(name);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
}

/* ---------- 캐시: 태블릿 여러 대가 몇 초마다 불러와도 시트를 매번 읽지 않게 ---------- */
function readAll() {
  var c = CacheService.getScriptCache(), hit = c.get("fair_db");
  if (hit) return JSON.parse(hit);
  var d = {};
  Object.keys(SHEETS).forEach(function (n) { d[n] = rows(n); });
  try { c.put("fair_db", JSON.stringify(d), 300); } catch (e) {}
  return d;
}
function dirty() { CacheService.getScriptCache().remove("fair_db"); }

function str(v, max) { return String(v == null ? "" : v).trim().slice(0, max); }

/* ---------- 요청 처리 ---------- */
function handle(b) {
  var act = b.action, isT = !!b.pin && String(b.pin) === pin();
  function needT() { if (!isT) throw new Error("PIN이 맞지 않아요."); }

  if (act === "all") return { ok: true, data: allData(isT, str(b.name, 20)) };
  if (act === "login") { needT(); return { ok: true }; }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var now = Date.now(), res = { ok: true };
    switch (act) {
      case "comment": {
        var nick = str(b.nick, 10), text = str(b.text, 100), book = str(b.book, 40);
        if (!nick || !text || !book) throw new Error("별칭과 내용을 적어 주세요.");
        var ok = mode() !== "share";
        sheet("comments").appendRow(toRow("comments", { id: Utilities.getUuid().slice(0, 8), book: book, nick: nick, name: str(b.name, 20), text: text, at: now, ok: ok, dev: str(b.dev, 60) }));
        res.approved = ok;
        break;
      }
      case "react": {
        var who = b.name ? "name:" + str(b.name, 20) : str(b.dev, 60), bk = str(b.book, 40), kind = b.kind === "fun" ? "fun" : "want";
        if (!who || !bk) throw new Error("이름이 필요해요.");
        var m = function (r) { return r.who === who && r.book === bk && r.kind === kind; };
        if (b.on) upsert("reactions", m, { who: who, book: bk, kind: kind, at: now });
        else removeWhere("reactions", m);
        break;
      }
      case "expect": {
        var en = str(b.name, 20), eb = str(b.book, 40), et = str(b.text, 200);
        if (!en || !eb) throw new Error("이름이 필요해요.");
        var em = function (r) { return r.name === en && r.book === eb; };
        if (et) upsert("expects", em, { name: en, book: eb, text: et, at: now });
        else removeWhere("expects", em);
        break;
      }
      case "post": {
        var pn = str(b.name, 20);
        if (!pn) throw new Error("이름이 필요해요.");
        upsert("posts", function (r) { return r.name === pn; }, { name: pn, b1: str(b.b1, 40), b2: str(b.b2, 40), reason: str(b.reason, 100), feel: str(b.feel, 200), at: now });
        break;
      }
      case "unexpect": {
        needT();
        var xn = str(b.name, 20), xb = str(b.book, 40);
        removeWhere("expects", function (r) { return r.name === xn && r.book === xb; });
        break;
      }
      case "unpost": {
        var un = str(b.name, 20);
        if (!un) throw new Error("이름이 필요해요.");
        removeWhere("posts", function (r) { return r.name === un; });
        break;
      }
      case "book": {
        needT();
        var k = str(b.k, 40), f = String(b.f);
        if (!k || !BOOK_FIELDS[f]) throw new Error("잘못된 항목이에요.");
        var cur = rows("books").filter(function (r) { return r.k === k; })[0] || { k: k };
        cur[f] = str(b.v, 500); cur.at = now;
        upsert("books", function (r) { return r.k === k; }, cur);
        break;
      }
      case "books": {
        needT();
        var all = rows("books");
        (b.items || []).forEach(function (it) {
          var k2 = str(it.k, 40); if (!k2) return;
          var c2 = all.filter(function (r) { return r.k === k2; })[0] || { k: k2 };
          c2.canva = str(it.canva, 500); c2.web = str(it.web, 500); c2.at = now;
          upsert("books", function (r) { return r.k === k2; }, c2);
        });
        break;
      }
      case "approve": {
        needT();
        var cs = sheet("comments"), cv = cs.getDataRange().getValues();
        for (var i = 1; i < cv.length; i++) if (String(cv[i][0]) === String(b.id)) cs.getRange(i + 1, 7).setValue(true);
        break;
      }
      case "delete": { needT(); removeWhere("comments", function (r) { return r.id === String(b.id); }); break; }
      case "wipe": { needT(); ["comments", "reactions", "posts", "expects"].forEach(clearSheet); break; }
      case "mode": { needT(); props().setProperty("fair_mode", b.mode === "share" ? "share" : "class"); break; }
      case "setpin": {
        needT();
        var np = str(b.newpin, 12);
        if (np.length < 4) throw new Error("4자리 이상으로 적어 주세요.");
        props().setProperty("fair_pin", np);
        break;
      }
      case "timer": {
        needT();
        var t = b.timer;
        if (!t || typeof t.i !== "number") throw new Error("타이머 정보가 없어요.");
        props().setProperty("fair_timer", JSON.stringify({ i: t.i, running: !!t.running, endsAt: Number(t.endsAt) || 0, remain: Number(t.remain) || 0, at: Number(t.at) || now }));
        return res; // 시트는 바뀌지 않음
      }
      default: throw new Error("알 수 없는 요청이에요.");
    }
    dirty();
    return res;
  } finally {
    lock.releaseLock();
  }
}

function allData(isT, name) {
  var d = readAll(), out = {};
  out.books = {};
  d.books.forEach(function (r) { out.books[r.k] = { canva: r.canva, web: r.web, intro: r.intro, why: r.why }; });
  out.comments = d.comments.filter(function (c) { return isT || c.ok; })
    .map(function (c) { return { id: c.id, book: c.book, nick: c.nick, text: c.text, at: c.at, ok: c.ok }; });
  out.posts = {};
  d.posts.forEach(function (p) { out.posts[p.name] = p; });
  out.reactions = {};
  d.reactions.forEach(function (r) {
    var x = out.reactions[r.book] || (out.reactions[r.book] = { want: 0, fun: 0 });
    x[r.kind] = (x[r.kind] || 0) + 1;
  });
  out.mode = mode();
  out.timer = timer();
  out.now = Date.now();
  if (name) {
    var me = { rx: {}, expect: {} }, key = "name:" + name;
    d.reactions.forEach(function (r) { if (r.who === key) { (me.rx[r.book] = me.rx[r.book] || {})[r.kind] = true; } });
    d.expects.forEach(function (x) { if (x.name === name) me.expect[x.book] = { t: x.text, at: x.at }; });
    out.me = me;
  }
  out.expects = d.expects.map(function (x) { return { name: x.name, book: x.book, text: x.text, at: x.at }; }); // 기대평은 댓글처럼 모두에게 보여요
  // 뱃지: 기대평을 공백 빼고 BADGE_MIN자 이상 쓴 책 (화면의 BADGE_MIN과 같은 값)
  out.badges = {};
  d.expects.forEach(function (x) {
    if (String(x.text).replace(/\s/g, "").length >= BADGE_MIN) (out.badges[x.name] = out.badges[x.name] || []).push(x.book);
  });
  return out;
}
