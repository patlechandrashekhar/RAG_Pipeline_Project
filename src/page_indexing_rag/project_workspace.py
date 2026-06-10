from __future__ import annotations

import os
import subprocess
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from hashlib import md5
from pathlib import Path

import streamlit as st

MAX_OUTPUT_CHARS = 12000
MAX_HISTORY_ENTRIES = 40
DEFAULT_SCAN_LIMIT = 800
SKIP_DIRS = {".git", ".venv", "node_modules", "__pycache__", ".pytest_cache", ".mypy_cache"}


def _state_key(name: str) -> str:
    return f"project_workspace_{name}"


def _key_token(value: str) -> str:
    return md5(value.encode("utf-8", errors="ignore")).hexdigest()[:10]


def _resolve_directory(path_text: str) -> Path | None:
    if not path_text or not path_text.strip():
        return None
    try:
        return Path(path_text).expanduser().resolve()
    except Exception:
        return None


def _normalize_project_input(path_text: str) -> Path | None:
    candidate = _resolve_directory(path_text)
    if candidate is None or not candidate.exists():
        return None
    if candidate.is_file():
        return candidate.parent
    if candidate.is_dir():
        return candidate
    return None


def _load_project_registry(default_root: Path) -> list[str]:
    key = _state_key("projects_registry")
    raw_items = st.session_state.get(key)
    if not isinstance(raw_items, list) or not raw_items:
        raw_items = [str(default_root)]

    normalized: list[str] = []
    seen: set[str] = set()
    for item in raw_items:
        normalized_path = _normalize_project_input(str(item))
        if normalized_path is None:
            continue
        as_str = str(normalized_path)
        if as_str not in seen:
            normalized.append(as_str)
            seen.add(as_str)

    if not normalized:
        normalized = [str(default_root)]

    st.session_state[key] = normalized
    return normalized


def _discover_projects(root_dir: Path) -> list[Path]:
    projects = []
    try:
        for child in sorted(root_dir.iterdir(), key=lambda p: p.name.lower()):
            if child.is_dir():
                projects.append(child)
    except Exception:
        pass
    return projects


def _project_label(project_path: str, root_dir: Path) -> str:
    path_obj = Path(project_path)
    if path_obj == root_dir:
        root_label = path_obj.name if path_obj.name else str(path_obj)
        return f"{root_label} - {path_obj}"
    return f"{path_obj.name} - {path_obj}"


def _ensure_threads(project_path: str) -> dict[str, dict]:
    thread_store = st.session_state.setdefault(_state_key("threads"), {})
    project_threads = thread_store.setdefault(project_path, {})
    if not project_threads:
        project_threads["Thread 1"] = {"notes": "", "terminal_history": []}
    return project_threads


def _ensure_active_thread(project_path: str, project_threads: dict[str, dict]) -> tuple[dict[str, str], str]:
    active_map = st.session_state.setdefault(_state_key("active_thread_by_project"), {})
    active_thread = active_map.get(project_path)
    if active_thread not in project_threads:
        active_thread = next(iter(project_threads))
        active_map[project_path] = active_thread
    return active_map, active_thread


def _to_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def _trim_output(text: str, limit: int = MAX_OUTPUT_CHARS) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + "\n\n... [output truncated]"


def _run_single_command(command: str, cwd: Path, timeout_seconds: int) -> dict:
    started_at = time.strftime("%Y-%m-%d %H:%M:%S")
    t0 = time.perf_counter()
    invoke = ["powershell", "-NoProfile", "-Command", command] if os.name == "nt" else ["/bin/bash", "-lc", command]

    try:
        completed = subprocess.run(
            invoke,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
        return {
            "command": command,
            "cwd": str(cwd),
            "started_at": started_at,
            "duration_sec": round(time.perf_counter() - t0, 3),
            "timed_out": False,
            "returncode": completed.returncode,
            "stdout": _trim_output(_to_text(completed.stdout)),
            "stderr": _trim_output(_to_text(completed.stderr)),
        }
    except subprocess.TimeoutExpired as exc:
        return {
            "command": command,
            "cwd": str(cwd),
            "started_at": started_at,
            "duration_sec": round(time.perf_counter() - t0, 3),
            "timed_out": True,
            "returncode": -1,
            "stdout": _trim_output(_to_text(exc.stdout)),
            "stderr": _trim_output(_to_text(exc.stderr)),
        }
    except Exception as exc:
        return {
            "command": command,
            "cwd": str(cwd),
            "started_at": started_at,
            "duration_sec": round(time.perf_counter() - t0, 3),
            "timed_out": False,
            "returncode": -2,
            "stdout": "",
            "stderr": _trim_output(str(exc)),
        }


def _run_commands(commands: list[str], cwd: Path, max_workers: int, timeout_seconds: int) -> list[dict]:
    clean_commands = [cmd.strip() for cmd in commands if cmd and cmd.strip()]
    if not clean_commands:
        return []

    if max_workers <= 1 or len(clean_commands) == 1:
        return [_run_single_command(cmd, cwd, timeout_seconds) for cmd in clean_commands]

    workers = min(max_workers, len(clean_commands))
    results: list[dict | None] = [None] * len(clean_commands)
    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(_run_single_command, cmd, cwd, timeout_seconds): idx
            for idx, cmd in enumerate(clean_commands)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results[idx] = future.result()
            except Exception as exc:
                results[idx] = {
                    "command": clean_commands[idx],
                    "cwd": str(cwd),
                    "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "duration_sec": 0.0,
                    "timed_out": False,
                    "returncode": -3,
                    "stdout": "",
                    "stderr": _trim_output(str(exc)),
                }

    return [item for item in results if item is not None]


@st.cache_data(show_spinner=False)
def _scan_project_files(project_dir: str, max_depth: int, max_files: int, include_hidden: bool) -> tuple[list[str], bool]:
    root = Path(project_dir)
    if not root.exists() or not root.is_dir():
        return [], False

    files: list[str] = []
    truncated = False

    for current_root, dirs, filenames in os.walk(root):
        current = Path(current_root)
        try:
            rel_dir = current.relative_to(root)
        except Exception:
            continue

        depth = 0 if str(rel_dir) == "." else len(rel_dir.parts)
        dirs.sort(key=str.lower)
        filenames.sort(key=str.lower)

        if not include_hidden:
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in SKIP_DIRS]
            filenames = [f for f in filenames if not f.startswith(".")]

        if depth >= max_depth:
            dirs[:] = []

        for filename in filenames:
            rel_path = filename if str(rel_dir) == "." else f"{rel_dir.as_posix()}/{filename}"
            files.append(rel_path)
            if len(files) >= max_files:
                truncated = True
                return files, truncated

    return files, truncated


def _read_file_preview(file_path: Path, max_chars: int) -> tuple[str, bool, str | None]:
    try:
        with file_path.open("r", encoding="utf-8", errors="replace") as handle:
            content = handle.read(max_chars + 1)
        if len(content) > max_chars:
            return content[:max_chars], True, None
        return content, False, None
    except Exception as exc:
        return "", False, str(exc)


def _save_text_file(file_path: Path, content: str) -> str | None:
    try:
        with file_path.open("w", encoding="utf-8") as handle:
            handle.write(content)
        return None
    except Exception as exc:
        return str(exc)


def _to_relative_project_path(path_value: str, project_root: Path) -> str | None:
    raw = path_value.strip()
    if not raw:
        return None

    root_resolved = project_root.resolve()
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = root_resolved / raw

    try:
        resolved = candidate.expanduser().resolve()
    except Exception:
        return None

    if not resolved.exists() or not resolved.is_file():
        return None

    try:
        rel = resolved.relative_to(root_resolved)
    except Exception:
        return None
    return rel.as_posix()


def render_project_workspace(default_root: Path | str) -> None:
    st.markdown(
        """
        <style>
        .chipagent-workspace-muted {
            color: #64748b;
            font-size: 0.85rem;
            margin-top: -0.2rem;
            margin-bottom: 0.75rem;
        }
        div[data-testid="stTabs"] button[role="tab"] {
            border-radius: 10px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    default_root_path = Path(default_root).resolve()
    project_options = _load_project_registry(default_root_path)

    project_key = _state_key("active_project")
    if st.session_state.get(project_key) not in project_options:
        st.session_state[project_key] = project_options[0]

    add_path_key = _state_key("project_add_path")
    add_col1, add_col2, add_col3, add_col4 = st.columns([2, 2, 1, 1])

    with add_col1:
        active_project = st.selectbox(
            "Projects",
            options=project_options,
            key=project_key,
            format_func=lambda p: _project_label(p, default_root_path),
            help="Choose the project you want ChipAgent to work on.",
        )

    with add_col2:
        st.text_input(
            "Add project path",
            key=add_path_key,
            placeholder=r"C:\path\to\project or C:\path\to\project\file.py",
        )

    with add_col3:
        st.write("")
        if st.button("Add", key=_state_key("add_project"), use_container_width=True):
            candidate = _normalize_project_input(st.session_state.get(add_path_key, ""))
            if candidate is None:
                st.warning("Enter a valid folder path (or a file path inside a project).")
            else:
                candidate_str = str(candidate)
                if candidate_str in project_options:
                    st.info("Project already added.")
                else:
                    project_options.append(candidate_str)
                    st.session_state[_state_key("projects_registry")] = project_options
                    st.session_state[project_key] = candidate_str
                    st.session_state[add_path_key] = ""
                    _scan_project_files.clear()
                    st.rerun()

    with add_col4:
        st.write("")
        if st.button("Remove", key=_state_key("remove_project"), use_container_width=True):
            if len(project_options) == 1:
                st.warning("At least one project must stay in the list.")
            else:
                project_options = [item for item in project_options if item != active_project]
                st.session_state[_state_key("projects_registry")] = project_options
                st.session_state[project_key] = project_options[0]
                _scan_project_files.clear()
                st.rerun()

    import_col1, import_col2 = st.columns([1, 3])
    with import_col1:
        if st.button("Import From Root", key=_state_key("import_from_root"), use_container_width=True):
            discovered = [str(path) for path in _discover_projects(default_root_path)]
            merged = project_options[:]
            seen = set(merged)
            for path in discovered:
                if path not in seen:
                    merged.append(path)
                    seen.add(path)
            st.session_state[_state_key("projects_registry")] = merged
            _scan_project_files.clear()
            st.rerun()
    with import_col2:
        st.markdown(
            f"<div class='chipagent-workspace-muted'>Active workspace: <code>{active_project}</code></div>",
            unsafe_allow_html=True,
        )

    project_threads = _ensure_threads(active_project)
    active_map, active_thread = _ensure_active_thread(active_project, project_threads)
    project_token = _key_token(active_project)
    thread_key = _state_key(f"active_thread_{project_token}")
    thread_names = list(project_threads.keys())

    if st.session_state.get(thread_key) not in thread_names:
        st.session_state[thread_key] = active_thread

    layout_left, layout_right = st.columns([1, 2.2])

    with layout_left:
        st.markdown("#### Threads")
        selected_thread = st.selectbox(
            "Current thread",
            options=thread_names,
            key=thread_key,
            help="Each thread keeps separate notes and terminal history.",
        )
        active_map[active_project] = selected_thread

        new_thread_key = _state_key(f"new_thread_name_{project_token}")
        st.text_input("New thread", key=new_thread_key, placeholder=f"Thread {len(thread_names) + 1}")

        thread_btn_col1, thread_btn_col2 = st.columns(2)
        with thread_btn_col1:
            if st.button("Create", key=_state_key(f"create_thread_{project_token}"), use_container_width=True):
                requested_name = st.session_state.get(new_thread_key, "").strip()
                thread_name = requested_name or f"Thread {len(project_threads) + 1}"
                if thread_name in project_threads:
                    st.warning("Thread name already exists.")
                else:
                    project_threads[thread_name] = {"notes": "", "terminal_history": []}
                    active_map[active_project] = thread_name
                    st.session_state[thread_key] = thread_name
                    st.session_state[new_thread_key] = ""
                    st.rerun()

        with thread_btn_col2:
            if st.button("Delete", key=_state_key(f"delete_thread_{project_token}"), use_container_width=True):
                if len(project_threads) == 1:
                    st.warning("At least one thread is required per project.")
                else:
                    project_threads.pop(selected_thread, None)
                    fallback_thread = next(iter(project_threads))
                    active_map[active_project] = fallback_thread
                    st.session_state[thread_key] = fallback_thread
                    st.rerun()

        thread_data = project_threads[selected_thread]
        thread_token = _key_token(selected_thread)
        notes_key = _state_key(f"thread_notes_{project_token}_{thread_token}")
        if notes_key not in st.session_state:
            st.session_state[notes_key] = thread_data.get("notes", "")

        notes_text = st.text_area(
            "Notes",
            key=notes_key,
            height=160,
            placeholder="Track plan, checkpoints, or findings for this thread.",
        )
        thread_data["notes"] = notes_text

    with layout_right:
        project_path = Path(active_project)
        files_tab, terminal_tab = st.tabs(["Files", "Terminal"])

        with files_tab:
            with st.expander("Explorer options", expanded=False):
                explorer_col1, explorer_col2, explorer_col3 = st.columns(3)
                max_depth = explorer_col1.slider(
                    "Depth",
                    1,
                    10,
                    4,
                    key=_state_key(f"max_depth_{project_token}"),
                )
                max_files = explorer_col2.slider(
                    "Max files",
                    100,
                    5000,
                    DEFAULT_SCAN_LIMIT,
                    step=100,
                    key=_state_key(f"max_files_{project_token}"),
                )
                include_hidden = explorer_col3.toggle(
                    "Include hidden",
                    value=False,
                    key=_state_key(f"include_hidden_{project_token}"),
                )

            files, is_truncated = _scan_project_files(str(project_path), max_depth, max_files, include_hidden)
            filter_key = _state_key(f"file_filter_{project_token}")
            file_filter = st.text_input("Find files", key=filter_key, placeholder="src/, .py, tests, README")

            if file_filter.strip():
                file_filter_lc = file_filter.strip().lower()
                filtered_files = [item for item in files if file_filter_lc in item.lower()]
            else:
                filtered_files = files

            truncation_text = " (truncated)" if is_truncated else ""
            st.caption(f"Showing {len(filtered_files)} / {len(files)} files{truncation_text}")

            if filtered_files:
                preview_options = filtered_files[:1500]
                preview_file_key = _state_key(f"preview_file_{project_token}")
                if st.session_state.get(preview_file_key) not in preview_options:
                    st.session_state[preview_file_key] = preview_options[0]

                open_path_key = _state_key(f"open_path_{project_token}")
                open_col1, open_col2 = st.columns([4, 1])
                open_col1.text_input(
                    "Open file path",
                    key=open_path_key,
                    placeholder=r"src\module.py or C:\full\path\to\file.py",
                )
                if open_col2.button("Open", key=_state_key(f"open_path_btn_{project_token}"), use_container_width=True):
                    rel_path = _to_relative_project_path(st.session_state.get(open_path_key, ""), project_path)
                    if rel_path is None:
                        st.warning("File path must point to an existing file inside the selected project.")
                    elif rel_path not in files:
                        st.warning("File is outside the current explorer scope. Increase depth/max files and retry.")
                    else:
                        st.session_state[filter_key] = ""
                        st.session_state[preview_file_key] = rel_path
                        st.rerun()

                selected_file = st.selectbox("Open preview", options=preview_options, key=preview_file_key)
                preview_chars = st.slider(
                    "Preview size (chars)",
                    500,
                    20000,
                    5000,
                    step=500,
                    key=_state_key(f"preview_chars_{project_token}"),
                )

                text, text_truncated, read_error = _read_file_preview(project_path / selected_file, preview_chars)
                if read_error:
                    st.error(f"Unable to read file: {read_error}")
                else:
                    if text_truncated:
                        st.caption(f"Preview truncated to {preview_chars} characters.")
                    st.code(text, language="text")

                st.markdown("##### Edit file")
                file_path = project_path / selected_file
                editor_buffer_key = _state_key(f"editor_buffer_{project_token}")
                editor_target_key = _state_key(f"editor_target_{project_token}")

                if st.session_state.get(editor_target_key) != selected_file:
                    editable_text, _, editable_error = _read_file_preview(file_path, 300000)
                    if editable_error is None:
                        st.session_state[editor_buffer_key] = editable_text
                    else:
                        st.session_state[editor_buffer_key] = ""
                    st.session_state[editor_target_key] = selected_file

                edited_content = st.text_area(
                    "Editor",
                    key=editor_buffer_key,
                    height=320,
                    placeholder="Selected file content appears here for editing.",
                )

                save_col1, save_col2 = st.columns([1, 4])
                with save_col1:
                    if st.button("Save File", key=_state_key(f"save_file_{project_token}"), use_container_width=True):
                        save_error = _save_text_file(file_path, edited_content)
                        if save_error:
                            st.error(f"Save failed: {save_error}")
                        else:
                            _scan_project_files.clear()
                            st.success(f"Saved: {selected_file}")
                with save_col2:
                    st.caption("Use this editor to change a specific file directly in the selected project.")
            else:
                st.info("No files matched the current filter.")

        with terminal_tab:
            command_key = _state_key(f"command_input_{project_token}_{thread_token}")
            commands_text = st.text_area(
                "Command input (one command per line)",
                key=command_key,
                height=140,
                placeholder="git status\npytest -q",
            )

            controls_col1, controls_col2, controls_col3 = st.columns([1, 1, 1.2])
            workers = controls_col1.slider(
                "Workers",
                1,
                8,
                2,
                key=_state_key(f"worker_count_{project_token}_{thread_token}"),
            )
            timeout_seconds = controls_col2.slider(
                "Timeout (s)",
                5,
                900,
                120,
                key=_state_key(f"timeout_seconds_{project_token}_{thread_token}"),
            )

            if controls_col3.button(
                "Run",
                key=_state_key(f"run_commands_{project_token}_{thread_token}"),
                use_container_width=True,
            ):
                commands = [line.strip() for line in commands_text.splitlines() if line.strip()]
                if not commands:
                    st.warning("Enter at least one command.")
                else:
                    mode = "parallel" if workers > 1 and len(commands) > 1 else "sequential"
                    with st.spinner(f"Running {len(commands)} command(s) in {mode} mode..."):
                        run_results = _run_commands(commands, project_path, workers, timeout_seconds)

                    history = thread_data.setdefault("terminal_history", [])
                    history.extend(run_results)
                    if len(history) > MAX_HISTORY_ENTRIES:
                        del history[:-MAX_HISTORY_ENTRIES]

                    success_count = sum(
                        1 for row in run_results if (not row.get("timed_out")) and row.get("returncode") == 0
                    )
                    st.success(
                        f"Completed {len(run_results)} command(s). "
                        f"Success: {success_count} | Failed/Timed out: {len(run_results) - success_count}"
                    )

            history = thread_data.get("terminal_history", [])
            history_col, clear_col = st.columns([3, 1])
            history_col.caption(f"History ({selected_thread}): {len(history)} entries")
            with clear_col:
                if history and st.button(
                    "Clear",
                    key=_state_key(f"clear_history_{project_token}_{thread_token}"),
                    use_container_width=True,
                ):
                    thread_data["terminal_history"] = []
                    st.rerun()

            if history:
                for index, row in enumerate(reversed(history[-15:]), start=1):
                    if row.get("timed_out"):
                        status = "TIMEOUT"
                    elif row.get("returncode") == 0:
                        status = "OK"
                    else:
                        status = f"ERR {row.get('returncode')}"

                    command_text = row.get("command", "")
                    with st.expander(f"{status} | {command_text}", expanded=(index == 1)):
                        st.caption(
                            f"Started: {row.get('started_at', '-')} | "
                            f"Duration: {row.get('duration_sec', 0.0)}s | "
                            f"CWD: `{row.get('cwd', '-')}`"
                        )

                        stdout_text = row.get("stdout", "")
                        stderr_text = row.get("stderr", "")

                        if stdout_text:
                            st.markdown("stdout")
                            st.code(stdout_text, language="text")
                        else:
                            st.caption("stdout: (empty)")

                        if stderr_text:
                            st.markdown("stderr")
                            st.code(stderr_text, language="text")
            else:
                st.info("No terminal commands executed in this thread yet.")
