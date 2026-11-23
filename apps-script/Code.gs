/**
 * 고전독서 박람회 기록 저장소 (구글 Apps Script 웹 앱)
 *
 * 설치
 * 1. 새 구글 시트를 만들고 [확장 프로그램] → [Apps Script]를 열어요.
 * 2. 이 파일 내용을 Code.gs에 모두 붙여 넣고 저장해요.
 * 3. [배포] → [새 배포] → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 나온 웹 앱 주소(https://script.google.com/macros/s/.../exec)를
 *    index.html 의 API_URL 에 넣어요.
 * 처음 선생님 PIN은 1234예요. 선생님 화면에서 바꿔 주세요.
 *
 * 시트 탭(books, comments, reactions, posts, expects)은 처음 쓸 때 자동으로 만들어져요.
 */

var SHEETS = {
  books:     ["k", "canva", "web", "intro", "why", "at"],
  comments:  ["id", "book", "nick", "name", "text", "at", "ok", "dev"],
  reactions: ["who", "book", "kind", "at"],
  posts:     ["name", "b1", "b2", "reason", "feel", "at"],
  expects:   ["name", "book", "text", "at"]
};
var NUM = { at: 1 }, BOOL = { ok: 1 };
var BOOK_FIELDS = { canva: 1, web: 1, intro: 1, why: 1 };

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
function pin() { return props().getProperty("pin") || "1234"; }
function mode() { return props().getProperty("mode") === "share" ? "share" : "class"; }
function timer() { try { return JSON.parse(props().getProperty("timer") || "null"); } catch (e) { return null; } }

/* ---------- 시트 ---------- */
function sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(SHEETS[name]); s.setFrozenRows(1); }
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
  var c = CacheService.getScriptCache(), hit = c.get("db");
  if (hit) return JSON.parse(hit);
  var d = {};
  Object.keys(SHEETS).forEach(function (n) { d[n] = rows(n); });
  try { c.put("db", JSON.stringify(d), 300); } catch (e) {}
  return d;
}
function dirty() { CacheService.getScriptCache().remove("db"); }

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
      case "mode": { needT(); props().setProperty("mode", b.mode === "share" ? "share" : "class"); break; }
      case "setpin": {
        needT();
        var np = str(b.newpin, 12);
        if (np.length < 4) throw new Error("4자리 이상으로 적어 주세요.");
        props().setProperty("pin", np);
        break;
      }
      case "timer": {
        needT();
        var t = b.timer;
        if (!t || typeof t.i !== "number") throw new Error("타이머 정보가 없어요.");
        props().setProperty("timer", JSON.stringify({ i: t.i, running: !!t.running, endsAt: Number(t.endsAt) || 0, remain: Number(t.remain) || 0, at: Number(t.at) || now }));
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
  if (isT) out.expects = d.expects;
  return out;
}
