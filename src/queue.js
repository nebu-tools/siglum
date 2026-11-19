// The queue of document operations. Works without Word: everything it needs
// from Word comes in through createQueue's arguments.
//
// Operations run one after another. An operation that plans (refresh, and
// everything that ends in one) may leave citations to write; the queue writes
// them in batches between operations, so the pane stays usable while a long
// renumbering goes on.
//
// The rule it keeps: the document is called current (the "settled" event, a
// state with nothing in progress, no `stale`) only after a plan that left
// nothing to write. Whenever unfinished writes are dropped, because an
// operation came or because a batch failed, the queue plans again after the
// operations that wait, unless one of them was itself a plan. If that plan
// fails too, the state says `stale` until a plan succeeds.

/** A plan's result, as refresh returns it: it has the count of citations it numbered. */
const isPlan = (r) => r?.cited != null;

/**
 * `word(fn)` runs fn(context) in a Word request context and returns its
 * result. `refresh(context)` plans from what the document holds and returns
 * {cited, work, signature, …}; `work` is {citations, done, total} or null.
 * `writeChunk(context, work, n)` writes the next n citations of `work` and
 * takes them off it. `signature(context)` is the order of the controls' tags.
 * `chunk` is the first batch's size and `now()` a clock in milliseconds.
 *
 * `emit(type, detail)` reports: "change" with the state, "done" with each
 * operation's result (or {error}), "settled" when the document is current.
 *
 * The state: {busy, progress: {done, total} or null, result: the last plan's
 * result, error: the error to show or null, stale: citation numbers are not
 * up to date and nothing will continue on its own}.
 */
export function createQueue({ word, refresh, writeChunk, signature, chunk = 30, now = () => Date.now(), emit = () => {} }) {
  const state = { busy: false, progress: null, result: null, error: null, stale: false };
  const tasks = []; // {fn, resolve, user}
  let work = null; // {r, w}: the citations the last plan left to write
  let unsettled = null; // the last plan, written in full, whose order is not yet checked
  let dirty = false; // writes were dropped: a plan must follow
  let recovering = false; // the queue has planned again on its own since the last operation
  let failure = null; // the last error that dropped writes
  let known = null; // the order of the tags when the document was last current
  let size = chunk;
  let pumping = false;

  const change = (patch) => {
    Object.assign(state, patch);
    emit("change", state);
  };

  /** Runs `fn(context)` after the operations before it. Resolves to its result, or null when it fails. */
  function run(fn) {
    return new Promise((resolve) => {
      tasks.push({ fn, resolve, user: true });
      pump();
    });
  }

  /** Plans again when the order of the controls changed since the document was last current, as after a deletion or a paste. */
  async function check() {
    if (state.busy || state.stale) return;
    const sig = await word((context) => signature(context));
    if (sig !== known && !state.busy) await run((context) => refresh(context));
  }

  async function pump() {
    if (pumping) return;
    pumping = true;
    change({ busy: true });
    try {
      for (;;) {
        if (tasks.length) {
          const task = tasks.shift();
          if (work) { work = null; dirty = true; }
          await perform(task);
        } else if (work) {
          await batch();
        } else if (dirty) {
          if (recovering) {
            // Planning again did not finish the work: say so, and wait for an operation.
            change({ stale: true, progress: null, error: state.error ?? failure });
            break;
          }
          recovering = true;
          tasks.push({ fn: (context) => refresh(context), resolve: () => {}, user: false });
        } else if (unsettled) {
          await settle();
        } else {
          break;
        }
      }
    } finally {
      pumping = false;
      change({ busy: false });
    }
  }

  async function perform({ fn, resolve, user }) {
    if (user) {
      recovering = false;
      change({ error: null });
    }
    try {
      const r = await word((context) => fn(context));
      if (isPlan(r)) {
        dirty = false;
        work = r.work ? { r, w: r.work } : null;
        unsettled = r.work ? null : r;
        change({ result: r, stale: false, progress: r.work ? { done: r.work.done, total: r.work.total } : null });
      }
      emit("done", r);
      resolve(r);
    } catch (e) {
      console.error(e);
      failure = e;
      if (user) change({ error: e });
      emit("done", { error: e });
      resolve(null);
    }
  }

  async function batch() {
    const { r, w } = work;
    const n = Math.min(size, w.citations.length);
    const t0 = now();
    try {
      await word((context) => writeChunk(context, w, n));
    } catch (e) {
      console.error(e);
      work = null;
      dirty = true;
      failure = e;
      return;
    }
    // Batches of about half a second of Word's time: Word does not answer
    // the keyboard while it works through one.
    size = Math.max(10, Math.min(200, Math.round((500 * n) / Math.max(1, now() - t0))));
    if (!w.citations.length) { work = null; unsettled = r; }
    change({ progress: work ? { done: w.done, total: w.total } : null });
  }

  /**
   * Notes the order of the controls and says that the document is current.
   * If the order changed since the plan read it, as when someone pasted a
   * citation during a long renumbering, it plans again.
   */
  async function settle() {
    const r = unsettled;
    unsettled = null;
    let sig;
    try {
      sig = await word((context) => signature(context));
    } catch (e) {
      console.error(e);
      dirty = true;
      failure = e;
      return;
    }
    if (r.signature != null && r.signature !== sig) {
      // Not current until that plan is written.
      dirty = true;
      tasks.push({ fn: (context) => refresh(context), resolve: () => {}, user: false });
      return;
    }
    known = sig;
    recovering = false;
    emit("settled");
  }

  return { run, check, state };
}
