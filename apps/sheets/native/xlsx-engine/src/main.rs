use std::collections::VecDeque;
use std::io::{self, BufRead, BufReader, BufWriter, Write};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, mpsc};
use std::thread;

use serde::{Deserialize, Serialize};
use serde_json::Value;
use xlsx_sidecar::archive::{
    EntryContent, archive_manifest, read_entries_to_dir, save_archive, scan_entries_for_text,
};
use xlsx_sidecar::recalc::{
    RecalcCache, RecalcEdit, RecalcRead, PREWARM_MAX_SOURCE_BYTES, PREWARM_MIN_SOURCE_BYTES,
    recalc_cells,
};
use xlsx_sidecar::{CellRange, SidecarError, WorkbookSessions};

const PROTOCOL_VERSION: u8 = 1;

fn default_locale() -> String {
    "zh".to_owned()
}

#[derive(Debug, Deserialize)]
#[serde(tag = "command", rename_all = "snake_case")]
enum Command {
    Open {
        path: PathBuf,
        #[serde(default = "default_locale")]
        locale: String,
        #[serde(default, rename = "shortDateFormat")]
        short_date_format: Option<String>,
    },
    ReadRange {
        #[serde(rename = "sessionId")]
        session_id: String,
        #[serde(rename = "sheetId")]
        sheet_id: String,
        range: CellRange,
    },
    Close {
        #[serde(rename = "sessionId")]
        session_id: String,
    },
    ReadMedia {
        #[serde(rename = "sessionId")]
        session_id: String,
        #[serde(rename = "visualId")]
        visual_id: String,
    },
    ReadFormulaCells {
        #[serde(rename = "sessionId")]
        session_id: String,
        #[serde(rename = "sheetId")]
        sheet_id: String,
    },
    Cancel {
        #[serde(rename = "targetRequestId")]
        target_request_id: String,
    },
    #[serde(rename_all = "camelCase")]
    ArchiveManifest { path: PathBuf },
    #[serde(rename_all = "camelCase")]
    ReadEntries {
        path: PathBuf,
        entries: Vec<String>,
        output_dir: PathBuf,
    },
    #[serde(rename_all = "camelCase")]
    ScanEntries {
        path: PathBuf,
        entries: Vec<String>,
        needle: String,
    },
    #[serde(rename_all = "camelCase")]
    ConvertWorkbook { path: PathBuf, target_path: PathBuf },
    #[serde(rename_all = "camelCase")]
    SaveArchive {
        source_path: PathBuf,
        target_path: PathBuf,
        replacements: Vec<EntryContent>,
        removals: Vec<String>,
        additions: Vec<EntryContent>,
    },
    #[serde(rename_all = "camelCase")]
    RecalcCells {
        path: PathBuf,
        edits: Vec<RecalcEdit>,
        reads: Vec<RecalcRead>,
    },
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Request {
    version: u8,
    request_id: String,
    #[serde(flatten)]
    command: Command,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Response {
    version: u8,
    request_id: String,
    ok: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    result: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<ErrorBody>,
}

#[derive(Serialize)]
struct ErrorBody {
    code: &'static str,
    message: String,
}

type SharedOutput = Arc<Mutex<BufWriter<io::Stdout>>>;

fn write_response(output: &SharedOutput, response: &Response) -> Result<(), ()> {
    let mut output = output
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if serde_json::to_writer(&mut *output, response).is_err()
        || output.write_all(b"\n").is_err()
        || output.flush().is_err()
    {
        return Err(());
    }
    Ok(())
}

/// Owns the formula-engine state so recalc_cells can run off the request
/// loop: a cold IronCalc import of a formula-heavy workbook takes minutes,
/// and running it inline starved every interactive read_range behind it
/// (each timing out client-side while the sheet rendered blank).
struct RecalcWorker {
    cache: Arc<Mutex<RecalcCache>>,
    /// Purges arrive from the request loop while a recalc may hold the cache
    /// lock for minutes; they queue here and apply at the next recalc. Safe
    /// to defer: recalc_cells re-imports on mtime+size mismatch anyway.
    pending_purges: Arc<Mutex<Vec<PathBuf>>>,
    busy: Arc<AtomicBool>,
    /// Serializes prewarms against each other (two tabs opened at once must
    /// not import concurrently). Real recalcs never wait on this flag: they
    /// either run before the prewarm starts or queue on the cache lock the
    /// prewarm holds, reusing the fresh model.
    prewarm_active: Arc<AtomicBool>,
}

/// Clears the busy flag even when recalc_cells panics mid-import.
struct BusyReset(Arc<AtomicBool>);

impl Drop for BusyReset {
    fn drop(&mut self) {
        self.0.store(false, Ordering::Release);
    }
}

impl RecalcWorker {
    fn new() -> Self {
        Self {
            cache: Arc::new(Mutex::new(RecalcCache::new())),
            pending_purges: Arc::new(Mutex::new(Vec::new())),
            busy: Arc::new(AtomicBool::new(false)),
            prewarm_active: Arc::new(AtomicBool::new(false)),
        }
    }

    fn queue_purge(&self, path: &PathBuf) {
        self.pending_purges
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .push(path.clone());
    }

    /// Close must reclaim the resident model right away when it can: a queued
    /// purge only applies at the next recalc, and a closed workbook may never
    /// be recalced again. Falls back to the queue while a recalc holds the
    /// cache lock.
    fn purge_now_or_queue(&self, path: &PathBuf) {
        match self.cache.try_lock() {
            Ok(mut cache) => cache.purge(path),
            Err(std::sync::TryLockError::Poisoned(poisoned)) => poisoned.into_inner().purge(path),
            Err(std::sync::TryLockError::WouldBlock) => self.queue_purge(path),
        }
    }

    fn drain_purges(&self, cache: &mut RecalcCache) {
        let purges: Vec<PathBuf> = self
            .pending_purges
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .drain(..)
            .collect();
        for purge in &purges {
            cache.purge(purge);
        }
    }

    /// The synchronous recalc body (also the unit-test entry point).
    fn run(
        &self,
        request_id: String,
        path: &PathBuf,
        edits: &[RecalcEdit],
        reads: &[RecalcRead],
    ) -> Response {
        let mut cache = self
            .cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        self.drain_purges(&mut cache);
        let response = match recalc_cells(&mut cache, path, edits, reads).and_then(to_json_value) {
            Ok(value) => Response::success(request_id, value),
            Err(error) => Response::from_error(request_id, error),
        };
        // A Close racing this recalc queued its purge after the drain above;
        // apply it before releasing the lock so the just-inserted model does
        // not outlive its session.
        self.drain_purges(&mut cache);
        response
    }

    /// Runs the recalc on a worker thread, writing the response when done.
    /// Returns an immediate failure response when one is already in flight —
    /// callers fail soft to cached values and retry later.
    fn dispatch(
        &self,
        request_id: String,
        path: PathBuf,
        edits: Vec<RecalcEdit>,
        reads: Vec<RecalcRead>,
        output: &SharedOutput,
    ) -> Option<Response> {
        if self
            .busy
            .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
            .is_err()
        {
            return Some(Response::failure(
                request_id,
                "recalc_busy",
                "Formula engine is busy with another recalculation.".into(),
            ));
        }
        let worker = Self {
            cache: Arc::clone(&self.cache),
            pending_purges: Arc::clone(&self.pending_purges),
            busy: Arc::clone(&self.busy),
            prewarm_active: Arc::clone(&self.prewarm_active),
        };
        let output = Arc::clone(output);
        let reply_id = request_id.clone();
        let spawned = thread::Builder::new()
            .name("xlsx-recalc".into())
            .spawn(move || {
                let reset = BusyReset(Arc::clone(&worker.busy));
                let response = worker.run(request_id, &path, &edits, &reads);
                // Clear busy before replying: a sequential caller issues its
                // next recalc the moment the response lands, and hitting the
                // still-set flag bounced it with recalc_busy.
                drop(reset);
                let _ = write_response(&output, &response);
            });
        match spawned {
            Ok(_) => None,
            Err(error) => {
                self.busy.store(false, Ordering::Release);
                Some(Response::failure(
                    reply_id,
                    "io_error",
                    format!("Unable to start the recalc worker: {error}"),
                ))
            }
        }
    }

    /// Builds the resident formula model in the background right after a
    /// successful open (PERF-1664). The first user edit then pays IronCalc's
    /// evaluate() alone instead of a full file import — on a 100k×20 book the
    /// import is ~4.7s of the ~5.3s cold recalc (13.8s on the audit's book),
    /// evaluate only ~0.5s. Size-gated: small books import in well under a
    /// second on demand, and mega-books (the model-blowup class) keep the
    /// on-demand behavior so their import never runs speculatively.
    ///
    /// A recalc arriving while the prewarm builds neither bounces nor
    /// double-imports: the prewarm holds the cache lock across the import,
    /// so the recalc queues on it and then reuses the fresh resident model —
    /// its total wait is bounded by the same import the old inline path
    /// paid in full. A recalc already in flight when the open completes
    /// makes the prewarm skip: that recalc imports its own path anyway, and
    /// two concurrent IronCalc imports would contend for memory and CPU.
    fn prewarm(&self, path: &PathBuf) {
        self.prewarm_within(path, PREWARM_MIN_SOURCE_BYTES, PREWARM_MAX_SOURCE_BYTES);
    }

    fn prewarm_within(&self, path: &PathBuf, min_bytes: u64, max_bytes: u64) {
        let Ok(metadata) = std::fs::metadata(path) else {
            return;
        };
        let size = metadata.len();
        if size < min_bytes || size > max_bytes {
            return;
        }
        // A recalc in flight populates the cache for its own path already;
        // racing it would mean two concurrent imports. Skip and let it win.
        if self.busy.load(Ordering::Acquire) {
            return;
        }
        // Only one prewarm at a time (two tabs opened back to back).
        if self
            .prewarm_active
            .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
            .is_err()
        {
            return;
        }
        let worker = Self {
            cache: Arc::clone(&self.cache),
            pending_purges: Arc::clone(&self.pending_purges),
            busy: Arc::clone(&self.busy),
            prewarm_active: Arc::clone(&self.prewarm_active),
        };
        let path = path.clone();
        let spawned = thread::Builder::new()
            .name("xlsx-prewarm".into())
            .spawn(move || {
                let _reset = BusyReset(Arc::clone(&worker.prewarm_active));
                // The lock is held across the whole build: a recalc that
                // arrives mid-build queues here and reuses the model instead
                // of importing the same bytes again. A panic-poisoned lock is
                // handled like run() does.
                let mut cache = worker
                    .cache
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
                let epoch_before = cache.purge_epoch();
                worker.drain_purges(&mut cache);
                // recalc_cells does the whole cold path — import, pin,
                // evaluate, residency insert with its eviction rules — and
                // contains its own panics; the value is discarded.
                let _ = recalc_cells(&mut cache, &path, &[], &[]);
                // A purge in between (close/save racing the build) means the
                // workbook this model belongs to is gone or superseded: drop
                // the fresh entry instead of pinning megabytes for a dead
                // session until some later eviction.
                if cache.purge_epoch() != epoch_before {
                    cache.purge(&path);
                }
                worker.drain_purges(&mut cache);
            });
        if spawned.is_err() {
            self.prewarm_active.store(false, Ordering::Release);
        }
    }
}

/// Request ids cancelled out of band by the reader thread, awaiting their
/// turn in the request queue. A cancel that lands after its target already
/// replied would pin the id here forever (ids are never reused), so the list
/// is capped — forgetting an id merely lets that request run as before.
#[derive(Clone)]
struct CancelledRequests(Arc<Mutex<VecDeque<String>>>);

const MAX_CANCELLED_REQUESTS: usize = 256;

impl CancelledRequests {
    fn new() -> Self {
        Self(Arc::new(Mutex::new(VecDeque::new())))
    }

    fn insert(&self, request_id: String) {
        let mut ids = self
            .0
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if ids.len() >= MAX_CANCELLED_REQUESTS {
            ids.pop_front();
        }
        ids.push_back(request_id);
    }

    fn take(&self, request_id: &str) -> bool {
        let mut ids = self
            .0
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        match ids.iter().position(|id| id == request_id) {
            Some(index) => {
                ids.remove(index);
                true
            }
            None => false,
        }
    }
}

/// Best-effort recovery of the client's request id from a line that failed to
/// parse. The failure reply echoes it (BUG-1666): clients dispatch replies by
/// requestId, so an unattributable error reply is dropped by the client and
/// its request waits out the full timeout — indistinguishable from the
/// sidecar never answering. Empty when the line is not even valid JSON with
/// a string requestId (nothing recoverable).
fn recovered_request_id(line: &str) -> String {
    match serde_json::from_str::<Value>(line) {
        Ok(value) => value
            .get("requestId")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .to_owned(),
        Err(_) => String::new(),
    }
}

fn main() {
    let stdin = io::stdin();
    let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));
    let mut sessions = WorkbookSessions::new();
    let recalc = RecalcWorker::new();
    let cancelled = CancelledRequests::new();

    // Requests execute on this thread in arrival order; a dedicated reader
    // drains stdin so a cancel can jump the queue while earlier requests are
    // still pending. The small bound keeps stdin backpressure close to the
    // old read-one-handle-one loop.
    let (queue, requests) = mpsc::sync_channel::<Request>(8);
    let reader_output = Arc::clone(&output);
    let reader_cancelled = cancelled.clone();
    let reader = thread::Builder::new()
        .name("xlsx-reader".into())
        .spawn(move || {
            for line in BufReader::new(stdin.lock()).lines() {
                let line = match line {
                    Ok(line) => line,
                    Err(error) => {
                        let _ = write_response(
                            &reader_output,
                            &Response::failure(
                                String::new(),
                                "stdin_error",
                                format!("Unable to read sidecar request: {error}"),
                            ),
                        );
                        break;
                    }
                };
                let request: Request = match serde_json::from_str(&line) {
                    Ok(request) => request,
                    Err(error) => {
                        let request_id = recovered_request_id(&line);
                        if write_response(
                            &reader_output,
                            &Response::failure(
                                request_id,
                                "invalid_json",
                                format!("Invalid sidecar request: {error}"),
                            ),
                        )
                        .is_err()
                        {
                            break;
                        }
                        continue;
                    }
                };
                if let Command::Cancel { target_request_id } = request.command {
                    let response = if request.version == PROTOCOL_VERSION {
                        reader_cancelled.insert(target_request_id);
                        Response::success(
                            request.request_id,
                            serde_json::json!({ "cancelled": true }),
                        )
                    } else {
                        Response::failure(
                            request.request_id,
                            "unsupported_version",
                            "Unsupported sidecar protocol version.".into(),
                        )
                    };
                    if write_response(&reader_output, &response).is_err() {
                        break;
                    }
                    continue;
                }
                if queue.send(request).is_err() {
                    break;
                }
            }
        });
    if reader.is_err() {
        return;
    }

    for request in requests {
        let response = handle_request(request, &mut sessions, &recalc, &cancelled, &output);
        if let Some(response) = response
            && write_response(&output, &response).is_err()
        {
            break;
        }
    }
}

#[cfg(test)]
fn handle_line(
    line: &str,
    sessions: &mut WorkbookSessions,
    recalc: &RecalcWorker,
    cancelled: &CancelledRequests,
    output: &SharedOutput,
) -> Option<Response> {
    let request: Request = match serde_json::from_str(line) {
        Ok(request) => request,
        Err(error) => {
            return Some(Response::failure(
                recovered_request_id(line),
                "invalid_json",
                format!("Invalid sidecar request: {error}"),
            ));
        }
    };
    handle_request(request, sessions, recalc, cancelled, output)
}

fn handle_request(
    request: Request,
    sessions: &mut WorkbookSessions,
    recalc: &RecalcWorker,
    cancelled: &CancelledRequests,
    output: &SharedOutput,
) -> Option<Response> {
    if cancelled.take(&request.request_id) {
        return Some(Response::failure(
            request.request_id,
            "cancelled",
            "Request was cancelled by the client.".into(),
        ));
    }
    if request.version != PROTOCOL_VERSION {
        return Some(Response::failure(
            request.request_id,
            "unsupported_version",
            "Unsupported sidecar protocol version.".into(),
        ));
    }

    let request_id = request.request_id;
    let result = match request.command {
        Command::Open {
            path,
            locale,
            short_date_format,
        } => {
            let result = sessions
                .open_with_locale(&path, &locale, short_date_format.as_deref())
                .and_then(to_json_value);
            if result.is_ok() {
                // Speculative background build of the formula model so the
                // user's first edit does not pay the cold import inline
                // (PERF-1664). Runs on its own thread and never blocks the
                // request loop; the open reply is not delayed by it.
                recalc.prewarm(&path);
            }
            result
        }
        Command::ReadRange {
            session_id,
            sheet_id,
            range,
        } => sessions
            .read_range(&session_id, &sheet_id, &range)
            .and_then(to_json_value),
        Command::Close { session_id } => {
            // A resident recalc model can dwarf the session itself (an
            // 8.8M-cell workbook's model holds ~1.2GB) — it must not outlive
            // the tab that needed it.
            if let Some(path) = sessions.session_path(&session_id) {
                recalc.purge_now_or_queue(&path);
            }
            sessions
                .close(&session_id)
                .and_then(|()| to_json_value(serde_json::json!({ "closed": true })))
        }
        Command::ReadMedia {
            session_id,
            visual_id,
        } => sessions
            .read_media(&session_id, &visual_id)
            .and_then(to_json_value),
        Command::ReadFormulaCells {
            session_id,
            sheet_id,
        } => sessions
            .read_formula_cells(&session_id, &sheet_id)
            .and_then(to_json_value),
        // Normally intercepted out of band by the reader thread; through this
        // path (tests) it still marks the target for the skip check above.
        Command::Cancel { target_request_id } => {
            cancelled.insert(target_request_id);
            to_json_value(serde_json::json!({ "cancelled": true }))
        }
        Command::ConvertWorkbook { path, target_path } => {
            xlsx_sidecar::convert::convert_to_xlsx(&path, &target_path).and_then(|result| {
                to_json_value(serde_json::json!({
                    "sheets": result.sheets,
                    "cells": result.cells,
                    // raw source size as this sidecar actually read it (BUG-1305)
                    "sourceBytes": result.source_bytes,
                }))
            })
        }
        Command::ArchiveManifest { path } => archive_manifest(&path)
            .and_then(|entries| to_json_value(serde_json::json!({ "entries": entries }))),
        Command::ReadEntries {
            path,
            entries,
            output_dir,
        } => read_entries_to_dir(&path, &entries, &output_dir)
            .and_then(|extracted| to_json_value(serde_json::json!({ "entries": extracted }))),
        Command::ScanEntries {
            path,
            entries,
            needle,
        } => scan_entries_for_text(&path, &entries, &needle)
            .and_then(|matches| to_json_value(serde_json::json!({ "matches": matches }))),
        Command::SaveArchive {
            source_path,
            target_path,
            replacements,
            removals,
            additions,
        } => {
            // The saved bytes supersede any resident formula model for the
            // target. The source is only read, so its warm model stays valid;
            // purging it would make every 30s crash-recovery copy of the open
            // workbook force a cold re-import on the next recalc. If the file
            // behind the source path changes later (e.g. the target is renamed
            // over it), the mtime+size guard in recalc_cells rebuilds anyway.
            recalc.queue_purge(&target_path);
            save_archive(
                &source_path,
                &target_path,
                &replacements,
                &removals,
                &additions,
            )
            .and_then(to_json_value)
        }
        Command::RecalcCells { path, edits, reads } => {
            return recalc.dispatch(request_id, path, edits, reads, output);
        }
    };
    Some(match result {
        Ok(value) => Response::success(request_id, value),
        Err(error) => Response::from_error(request_id, error),
    })
}

fn to_json_value<T: Serialize>(value: T) -> Result<Value, SidecarError> {
    serde_json::to_value(value).map_err(SidecarError::from)
}

impl Response {
    fn success(request_id: String, result: Value) -> Self {
        Self {
            version: PROTOCOL_VERSION,
            request_id,
            ok: true,
            result: Some(result),
            error: None,
        }
    }

    fn from_error(request_id: String, error: SidecarError) -> Self {
        let code = match error {
            SidecarError::InvalidRequest(_) => "invalid_request",
            SidecarError::Io(_) => "io_error",
            SidecarError::Workbook(_) => "workbook_error",
        };
        Self::failure(request_id, code, error.to_string())
    }

    fn failure(request_id: String, code: &'static str, message: String) -> Self {
        Self {
            version: PROTOCOL_VERSION,
            request_id,
            ok: false,
            result: None,
            error: Some(ErrorBody { code, message }),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc::base::Model;
    use ironcalc::export::save_to_xlsx;
    use std::path::Path;

    fn write_workbook(path: &Path) {
        let mut model = Model::new_empty("fixture", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "=A1*2".to_string()).unwrap();
        model.evaluate();
        save_to_xlsx(&model, path.to_str().unwrap()).unwrap();
    }

    fn expect_ok(response: Response) -> Value {
        assert!(
            response.ok,
            "{:?}",
            response.error.map(|error| error.message)
        );
        response.result.unwrap()
    }

    fn recalc_reads() -> Vec<RecalcRead> {
        serde_json::from_value(serde_json::json!([{
            "sheet": "Sheet1",
            "range": { "startRow": 0, "endRow": 1, "startColumn": 0, "endColumn": 0 },
        }]))
        .unwrap()
    }

    /// Crash-recovery copies save every 30s with the open workbook as the
    /// (unchanged) source; that must not evict its warm formula model even
    /// though save_archive queues a purge for the target path.
    #[test]
    fn save_archive_keeps_the_source_resident_model_warm() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        let target = dir.path().join("recovery.xlsx");
        write_workbook(&source);
        let mut sessions = WorkbookSessions::new();
        let recalc = RecalcWorker::new();
        let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));

        let first = expect_ok(recalc.run("r".into(), &source, &[], &recalc_reads()));
        assert_eq!(first["cached"], serde_json::json!(false));

        let save = serde_json::json!({
            "version": 1,
            "requestId": "s",
            "command": "save_archive",
            "sourcePath": source,
            "targetPath": target,
            "replacements": [],
            "removals": [],
            "additions": [],
        });
        expect_ok(
            handle_line(
                &save.to_string(),
                &mut sessions,
                &recalc,
                &CancelledRequests::new(),
                &output,
            )
            .unwrap(),
        );

        let second = expect_ok(recalc.run("r2".into(), &source, &[], &recalc_reads()));
        assert_eq!(second["cached"], serde_json::json!(true));
    }

    /// A cancel marks its target; the target is then skipped unexecuted, and
    /// the mark is consumed (a later request reusing the id would run).
    #[test]
    fn cancelled_requests_are_skipped_once() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let mut sessions = WorkbookSessions::new();
        let recalc = RecalcWorker::new();
        let cancelled = CancelledRequests::new();
        let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));
        let cancel = serde_json::json!({
            "version": 1,
            "requestId": "c",
            "command": "cancel",
            "targetRequestId": "victim",
        });
        expect_ok(
            handle_line(
                &cancel.to_string(),
                &mut sessions,
                &recalc,
                &cancelled,
                &output,
            )
            .unwrap(),
        );

        let open = serde_json::json!({
            "version": 1,
            "requestId": "victim",
            "command": "open",
            "path": source,
        });
        let skipped = handle_line(
            &open.to_string(),
            &mut sessions,
            &recalc,
            &cancelled,
            &output,
        )
        .unwrap();
        assert!(!skipped.ok);
        assert_eq!(skipped.error.unwrap().code, "cancelled");

        expect_ok(
            handle_line(
                &open.to_string(),
                &mut sessions,
                &recalc,
                &cancelled,
                &output,
            )
            .unwrap(),
        );
    }

    /// A queued purge for a path evicts its resident model at the next run.
    #[test]
    fn queued_purge_applies_before_the_next_recalc() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();

        expect_ok(recalc.run("r".into(), &source, &[], &recalc_reads()));
        recalc.queue_purge(&source);
        let second = expect_ok(recalc.run("r2".into(), &source, &[], &recalc_reads()));
        assert_eq!(second["cached"], serde_json::json!(false));
    }

    /// Close purges the resident model on the spot when no recalc holds the
    /// cache — a queued purge would wait for a recalc that may never come.
    #[test]
    fn close_purge_applies_immediately_when_cache_is_free() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();

        expect_ok(recalc.run("r".into(), &source, &[], &recalc_reads()));
        recalc.purge_now_or_queue(&source);
        assert!(recalc.pending_purges.lock().unwrap().is_empty());
        let second = expect_ok(recalc.run("r2".into(), &source, &[], &recalc_reads()));
        assert_eq!(second["cached"], serde_json::json!(false));
    }

    /// While a recalc holds the cache lock, Close falls back to the queue
    /// instead of blocking the request loop.
    #[test]
    fn close_purge_queues_while_cache_is_locked() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();

        let guard = recalc.cache.lock().unwrap();
        recalc.purge_now_or_queue(&source);
        drop(guard);
        assert_eq!(recalc.pending_purges.lock().unwrap().len(), 1);
    }

    /// While one recalculation is in flight, a second request short-circuits
    /// with recalc_busy instead of queueing behind the import.
    #[test]
    fn concurrent_recalc_reports_busy() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();
        let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));

        recalc.busy.store(true, Ordering::Release);
        let response = recalc
            .dispatch("r".into(), source, Vec::new(), recalc_reads(), &output)
            .unwrap();
        assert!(!response.ok);
        assert_eq!(response.error.unwrap().code, "recalc_busy");
    }

    /// BUG-1666: a request violating the wire shape (read_range with the
    /// range as a string) must get an immediate failure reply that echoes the
    /// client's requestId — clients dispatch replies by requestId, so an
    /// unattributable error is dropped and the request waits out its timeout,
    /// which looks exactly like the sidecar staying silent.
    #[test]
    fn wire_shape_violation_replies_with_the_clients_request_id() {
        let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));
        let response = handle_line(
            r#"{"version":1,"requestId":"bad1","command":"read_range","sessionId":"s","sheetId":"sh","range":"A1:B2"}"#,
            &mut WorkbookSessions::new(),
            &RecalcWorker::new(),
            &CancelledRequests::new(),
            &output,
        )
        .unwrap();
        assert!(!response.ok);
        assert_eq!(response.request_id, "bad1");
        assert_eq!(response.error.unwrap().code, "invalid_json");
    }

    /// A line that is not JSON at all still fails immediately; there is no id
    /// to recover, so the reply keeps the legacy empty requestId.
    #[test]
    fn non_json_line_fails_closed_with_an_empty_request_id() {
        let output: SharedOutput = Arc::new(Mutex::new(BufWriter::new(io::stdout())));
        let response = handle_line(
            "not json at all",
            &mut WorkbookSessions::new(),
            &RecalcWorker::new(),
            &CancelledRequests::new(),
            &output,
        )
        .unwrap();
        assert!(!response.ok);
        assert_eq!(response.request_id, "");
        assert_eq!(response.error.unwrap().code, "invalid_json");
    }

    /// Polls the cache until `predicate` holds or the deadline passes
    /// (prewarm builds its model on a background thread).
    fn wait_for(cache: &Arc<Mutex<RecalcCache>>, predicate: impl Fn(&RecalcCache) -> bool) -> bool {
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
        while std::time::Instant::now() < deadline {
            let cache = cache
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            if predicate(&cache) {
                return true;
            }
            drop(cache);
            thread::sleep(std::time::Duration::from_millis(20));
        }
        false
    }

    /// PERF-1664: a successful open of a size-eligible book starts a
    /// background prewarm whose resident model then serves the first real
    /// recalc as cached=true — the edit pays evaluate(), not the import.
    #[test]
    fn prewarm_serves_the_first_recalc_from_the_resident_model() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();

        recalc.prewarm_within(&source, 1, u64::MAX);
        assert!(wait_for(&recalc.cache, |cache| cache.has_resident(&source)));

        // The first real recalc hits the resident model.
        let first = expect_ok(recalc.run("r".into(), &source, &[], &recalc_reads()));
        assert_eq!(first["cached"], serde_json::json!(true));
        // And the prewarm flag was released with the prewarm thread.
        assert!(!recalc.prewarm_active.load(Ordering::Acquire));
    }

    /// Books outside the size window keep today's on-demand behavior: tiny
    /// books import in well under a second at the first edit, and mega-books
    /// (the model-blowup class) never import speculatively on open.
    #[test]
    fn prewarm_skips_books_outside_the_size_window() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();

        // The real threshold (1MB minimum): the fixture is far below it.
        recalc.prewarm(&source);
        // ...and an inverse window above the fixture's size.
        recalc.prewarm_within(&source, 0, 0);
        thread::sleep(std::time::Duration::from_millis(200));
        assert!(recalc.cache.lock().unwrap().is_empty());
        assert!(!recalc.prewarm_active.load(Ordering::Acquire));
    }

    /// A recalc in flight owns the worker: the prewarm must not race it with
    /// a second import of the same bytes.
    #[test]
    fn prewarm_skips_while_a_recalc_is_busy() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();
        recalc.busy.store(true, Ordering::Release);

        recalc.prewarm_within(&source, 1, u64::MAX);
        thread::sleep(std::time::Duration::from_millis(200));
        assert!(recalc.cache.lock().unwrap().is_empty());
        // The skip releases nothing: the flag belongs to the in-flight recalc.
        assert!(recalc.busy.load(Ordering::Acquire));
        assert!(!recalc.prewarm_active.load(Ordering::Acquire));
        recalc.busy.store(false, Ordering::Release);
    }

    /// Two tabs opened back to back must not import concurrently: the second
    /// prewarm skips while the first is still building.
    #[test]
    fn prewarms_do_not_overlap() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();
        recalc.prewarm_active.store(true, Ordering::Release);

        recalc.prewarm_within(&source, 1, u64::MAX);
        thread::sleep(std::time::Duration::from_millis(200));
        assert!(recalc.cache.lock().unwrap().is_empty());
        // The flag belongs to the first prewarm and is untouched.
        assert!(recalc.prewarm_active.load(Ordering::Acquire));
        recalc.prewarm_active.store(false, Ordering::Release);
    }

    /// A close that lands while the prewarm holds the cache lock queues its
    /// purge; the prewarm applies it before finishing, so the model of an
    /// already-closed workbook does not linger.
    #[test]
    fn prewarm_honors_a_close_that_raced_it() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("open.xlsx");
        write_workbook(&source);
        let recalc = RecalcWorker::new();
        recalc.queue_purge(&source);

        recalc.prewarm_within(&source, 1, u64::MAX);
        // Wait until the prewarm thread finished (busy released), then the
        // drained purge must have dropped the just-built model.
        let released = wait_for(&recalc.cache, |_| !recalc.prewarm_active.load(Ordering::Acquire));
        assert!(released);
        assert!(recalc.cache.lock().unwrap().is_empty());
    }
}
