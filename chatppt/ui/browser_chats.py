"""Bridge that keeps chats in the user's own browser (IndexedDB) through an invisible inline CCv2 component.

Edit when: changing when chats are restored / saved / erased in the browser. Record layout, size limits and
validation live in chatppt/chat_store.py.

Flow per browser session:
1. First run: the component reads every stored record once and sends them up as state "loaded".
2. `_on_loaded` (a callback, before any widget exists) restores the chats; the untouched empty chat that
   was created at startup is replaced and the most recent restored chat is opened.
3. Later runs: for the active chat, each part ("chat", "source", "images") is fingerprinted; changed parts
   become versioned writes (latest write per record). The browser applies them in the background and does
   NOT report back, because any value sent up from a component costs a full app rerun. Instead each write is
   re-sent for RESEND_RUNS runs (the browser skips versions it already applied), so a rerun that interrupts
   a render cannot lose it, and the browser drains writes that arrive while it is busy.
Nothing is written before step 2, and nothing at all while "Keep my chats on this browser" is off.
"""

import time

import streamlit as st

from chatppt.chat_store import MAX_STORED_CHATS, PARTS, is_storable, part_payload, part_signature, record_key, restore_chats
from chatppt.ui.state import create_chat, new_chat

COMPONENT_KEY = "browser_chats"
# A queued write is included in this many runs (its own + the next ones) before it is dropped from the queue
RESEND_RUNS = 3

_JS = """
const DB_NAME = "chatppt";
const STORE = "records";

function openDb() {
  return new Promise((resolve, reject) => {
    const req = window.indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function readAll(db) {
  return new Promise((resolve, reject) => {
    const out = {};
    const req = db.transaction(STORE, "readonly").objectStore(STORE).openCursor();
    req.onsuccess = () => {
      const cursor = req.result;
      if (cursor) {
        out[cursor.key] = cursor.value;
        cursor.continue();
      } else {
        resolve(out);
      }
    };
    req.onerror = () => reject(req.error);
  });
}

function applyOps(db, ops) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    const store = tx.objectStore(STORE);
    for (const op of ops) {
      if (op.op === "put") {
        store.put(op.value, op.key);
      } else if (op.op === "delete_chat") {
        for (const part of ["chat", "source", "images"]) store.delete(op.id + "/" + part);
      } else if (op.op === "clear") {
        store.clear();
      }
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

// Writes are applied in the background without telling Python (that would cost a full app rerun per save).
// The newest ops seen are kept per session; after each batch the loop checks for ops that arrived meanwhile.
function drain(session, setStateValue) {
  const applied = window.__chatpptChatsApplied;
  const busy = window.__chatpptChatsBusy;
  if (busy[session]) return;
  const pending = (window.__chatpptChatsLatest[session] || [])
    .filter((op) => op.version > (applied[session] || 0))
    .sort((a, b) => a.version - b.version);
  if (!pending.length) return;
  busy[session] = true;
  const top = Math.max.apply(null, pending.map((op) => op.version));
  openDb()
    .then((db) => applyOps(db, pending))
    .then(() => {
      applied[session] = Math.max(applied[session] || 0, top);
      // Written ops (which may hold large source text / figures) are released
      window.__chatpptChatsLatest[session] = window.__chatpptChatsLatest[session].filter((op) => op.version > applied[session]);
    })
    .catch((e) => setStateValue("save_error", String(e)))
    .finally(() => {
      busy[session] = false;
      drain(session, setStateValue);
    });
}

export default function (component) {
  const { data, setStateValue } = component;
  if (!data) return;
  // Per browser-session bookkeeping on window (the component is not iframed, so window is the app page)
  const loaded = (window.__chatpptChatsLoaded = window.__chatpptChatsLoaded || {});
  window.__chatpptChatsApplied = window.__chatpptChatsApplied || {};
  window.__chatpptChatsBusy = window.__chatpptChatsBusy || {};
  const latest = (window.__chatpptChatsLatest = window.__chatpptChatsLatest || {});
  const session = data.session || "default";

  if (data.need_load && !loaded[session]) {
    loaded[session] = true;
    openDb()
      .then(readAll)
      .then((records) => setStateValue("loaded", { records: records, error: null }))
      .catch((e) => setStateValue("loaded", { records: {}, error: String(e) }));
  }

  // Merge, never replace: an op dropped from Python's resend window must still be written if it was not yet
  const known = latest[session] || [];
  const seen = new Set(known.map((op) => op.version));
  const done = window.__chatpptChatsApplied[session] || 0;
  latest[session] = known.concat((data.ops || []).filter((op) => op.version > done && !seen.has(op.version)));
  drain(session, setStateValue);
}
"""

_BROWSER_CHATS = st.components.v2.component("chatppt_browser_chats", html="<div></div>\n", js=_JS)


def _state() -> dict:
    return st.session_state.get(COMPONENT_KEY) or {}


def _on_loaded() -> None:
    """Restore the chats the browser kept. Runs as a callback, before any widget of this run exists."""
    ss = st.session_state
    loaded = _state().get("loaded") or {}
    ss["_chatdb_loaded"] = True
    ss["_chatdb_error"] = loaded.get("error")
    restored = restore_chats(loaded.get("records") or {})
    if not restored:
        return
    sigs = ss.setdefault("_chatdb_sigs", {})
    current = ss.chats.get(ss.active_chat)
    if current and not current["messages"] and not current["slides"] and current["pending"] is None:
        # The empty chat created at startup is replaced by the user's real chats
        ss.chats.pop(current["id"], None)
        if current["id"] in ss.chat_order:
            ss.chat_order.remove(current["id"])
    for record in restored:
        if record["id"] in ss.chats:
            continue
        chat = new_chat(record["id"])
        chat.update(record)
        ss.chats[chat["id"]] = chat
        ss.chat_order.append(chat["id"])
        sigs[chat["id"]] = {part: part_signature(chat, part) for part in PARTS}  # already stored as-is
    if not ss.chats:
        create_chat()
    if ss.active_chat not in ss.chats:
        ss.active_chat = ss.chat_order[0]


def erase_browser_chats() -> None:
    """Erase every chat kept in this browser (Forget, or turning 'Keep my chats' off)."""
    st.session_state["_chatdb_clear"] = True


def _queue(ops: list, op: dict, version: int) -> int:
    """Queue an op with the next version; an older queued write of the same record is superseded."""
    version += 1
    op["version"] = version
    op["runs_left"] = RESEND_RUNS
    if op["op"] == "clear":
        ops.clear()
    else:
        ops[:] = [o for o in ops if not (o.get("key") and o.get("key") == op.get("key"))]
    ops.append(op)
    return version


def sync_browser_chats() -> None:
    """Mount the invisible bridge (every run): restore once, then save the active chat when it changed."""
    ss = st.session_state
    ops = list(ss.get("_chatdb_ops", []))
    version = ss.get("_chatdb_version", 0)
    sigs = ss.setdefault("_chatdb_sigs", {})

    if ss.pop("_chatdb_clear", False):
        version = _queue(ops, {"op": "clear"}, version)
        sigs.clear()
    elif ss.get("_chatdb_loaded") and ss.settings.get("remember_chats", True):
        for cid in [c for c in sigs if c not in ss.chats]:  # chats deleted in the sidebar
            version = _queue(ops, {"op": "delete_chat", "id": cid}, version)
            del sigs[cid]
        chat = ss.chats.get(ss.active_chat)
        if chat and is_storable(chat):
            stored = sigs.setdefault(chat["id"], {})
            changed = [p for p in PARTS if stored.get(p) != part_signature(chat, p)]
            if changed:
                chat["updated_at"] = time.time()
            for part in changed:
                version = _queue(ops, {"op": "put", "key": record_key(chat["id"], part),
                                       "value": part_payload(chat, part)}, version)
                stored[part] = part_signature(chat, part)
        # Keep only the most recent chats in the browser
        if len(sigs) > MAX_STORED_CHATS:
            by_age = sorted(sigs, key=lambda c: ss.chats[c].get("updated_at", 0) if c in ss.chats else 0)
            for cid in by_age[: len(sigs) - MAX_STORED_CHATS]:
                version = _queue(ops, {"op": "delete_chat", "id": cid}, version)
                del sigs[cid]

    data = {"session": ss["_session_nonce"], "need_load": not ss.get("_chatdb_loaded"),
            "ops": [{k: v for k, v in o.items() if k != "runs_left"} for o in ops]}
    _BROWSER_CHATS(
        key=COMPONENT_KEY,
        data=data,
        on_loaded_change=_on_loaded,
        on_save_error_change=lambda: None,
        height="content",
    )
    # Each write rides along with RESEND_RUNS runs, then leaves the queue (the browser keeps it until written)
    for op in ops:
        op["runs_left"] -= 1
    ss["_chatdb_ops"] = [o for o in ops if o["runs_left"] > 0]
    ss["_chatdb_version"] = version
