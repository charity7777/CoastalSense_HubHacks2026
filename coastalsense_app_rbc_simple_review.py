"""CoastalSense — RBC interface + simpler optional MeanCount corrections.

Save beside detect.py and best.pt in the INNER yolov7-main folder, then use
your existing virtual environment to run:
    python -m streamlit run coastalsense_app_rbc_simple_review.py --server.address 127.0.0.1 --server.port 8505

Uses the original N-MARINE model, CPU inference, CLAHE rules, confidence
defaults and short-video sampling. Adds a provisional, per-frame MeanCount-style
average and an optional 15-minute / 15-sample plan from Coastal Action's DRAFT
Fish Video Analysis Procedure. No new packages or retraining are needed.
Adds an optional, lightweight count-correction panel and simplified two-sheet
Excel reporting. The preliminary mean always remains visible. Full-window species
notes are available separately without blocking counts. Model detections still
use one still image per sampled timestamp.
No ByteTrack or scientifically validated fish-counting accuracy is claimed.
All previous CoastalSense scripts remain unchanged.
"""

from collections import Counter
from datetime import datetime, timezone
from html import escape
from pathlib import Path
import inspect
import base64
import hashlib
import io
import math
import subprocess
import sys
import tempfile
import time

import cv2
import numpy as np
from PIL import Image
import streamlit as st
import xlsxwriter


PROJECT = Path(__file__).resolve().parent
MODEL = PROJECT / "best.pt"
SAMPLE_RATE = 2
PROTOCOL_SAMPLES = 15
PROTOCOL_INTERVAL_SECONDS = 60
PROTOCOL_WINDOW_SECONDS = 900
NON_FISH_CLASSES = {"Whelk"}
LOW_CONFIDENCE = 0.25
BLUE = "#0051A5"
NAVY = "#002750"
YELLOW = "#FEDF01"
STATE_KEY = "cs_simple_review_result"
REVIEW_KEY = "cs_simple_review_state"

# Retain the class order used by the supplied N-MARINE checkpoint.
CLASS_NAMES = [
    "Atlantic cod", "Roughhead grenadier", "Atlantic halibut",
    "Redfish mentella", "Thorny skate", "Striped wolffish",
    "Spinytail skate", "Whelk", "Northern wolffish",
]

STYLES = """
<style>
:root { --cs-blue:#0051A5; --cs-navy:#002750; --cs-yellow:#FEDF01; }
.stApp, [data-testid="stAppViewContainer"] { background:#F5F8FC; color:#18334E; }
[data-testid="stHeader"] { background:rgba(245,248,252,.96); }
.block-container { max-width:1400px; padding-top:2rem; padding-bottom:2.5rem; }
html, body, [class*="css"] { font-family:"Segoe UI",Arial,sans-serif; }
h1,h2,h3 { color:#002750; letter-spacing:-.035em; }
h3 { font-size:1.25rem !important; }
[data-testid="stMarkdownContainer"] p { line-height:1.6; }
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p { color:#536B83; }
[data-testid="stWidgetLabel"] p { color:#18334E; font-weight:600; }
section[data-testid="stSidebar"] { background:#FFF; border-right:1px solid #DEE7F1; }
section[data-testid="stSidebar"] > div { padding-top:1.6rem; }
section[data-testid="stSidebar"] h3 { font-size:1rem !important; }
section[data-testid="stSidebar"] hr { margin:1.25rem 0; border-color:#E2EAF3; }
[data-testid="stFileUploader"] { background:#FFF; border:1px dashed #9DB8D6;
    border-radius:16px; padding:8px; }
[data-testid="stFileUploaderDropzone"] { background:#F3F8FF; border-radius:12px; }
[data-testid="stFileUploaderDropzone"] [data-testid="stMarkdownContainer"] p { color:#18334E; }
[data-testid="stFileUploaderDropzone"] small { color:#536B83; }
.stButton > button, .stDownloadButton > button {
    border:1px solid #BDCFDF; border-radius:10px; background:#FFF; color:#0051A5;
    font-weight:650; min-height:2.8rem; transition:background .15s,border-color .15s; }
.stButton > button:hover, .stDownloadButton > button:hover {
    color:#002750; border-color:#0051A5; background:#EAF3FF; }
.stButton > button[kind="primary"] { background:#FEDF01; border-color:#FEDF01;
    color:#002750; box-shadow:0 4px 12px rgba(0,39,80,.08); }
.stButton > button[kind="primary"]:hover { background:#F3D500; border-color:#F3D500; }
.stButton > button:disabled { background:#E8EEF5; border-color:#E8EEF5; color:#6E8299; }
.stButton > button:focus-visible, .stDownloadButton > button:focus-visible {
    outline:3px solid #73B0E3; outline-offset:3px; }
[data-baseweb="input"], [data-baseweb="select"] > div {
    background:#FFF; color:#18334E; border-color:#CBD8E7; border-radius:8px; }
input { color:#18334E !important; }
[data-baseweb="radio"] > div:first-child { background-color:#CCD9E8; }
[data-baseweb="radio"]:has(input:checked) > div:first-child { background-color:#0051A5; }
[data-baseweb="slider"] [role="slider"] { background-color:#0051A5; }
[data-baseweb="tab-list"] { gap:1.25rem; background:transparent; border-bottom:1px solid #DBE5F0; }
[data-baseweb="tab"] { color:#5D7288; font-weight:600; padding:12px 2px; }
[data-baseweb="tab"][aria-selected="true"] { color:#0051A5; }
[data-baseweb="tab-highlight"] { background:#0051A5; height:3px; }
[data-testid="stExpander"] { background:#FFF; border-radius:12px; border-color:#DBE5F0; }
[data-testid="stExpander"] summary { color:#18334E; }
[data-testid="stImage"] img { border-radius:13px; }
[data-testid="stAlert"] { border-radius:12px; }
[data-testid="stProgressBar"] div[role="progressbar"] > div { background-color:#0051A5; }
.cs-brand { display:flex; align-items:center; gap:10px; margin-bottom:1.25rem; }
.cs-mark { background:#0051A5; border-radius:13px; width:43px; height:43px;
    display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.cs-brand-name { font-size:1.2rem; font-weight:750; color:#002750; letter-spacing:-.04em; }
.cs-brand-sub { font-size:.62rem; color:#677E94; letter-spacing:.14em; font-weight:700; }
.cs-nav { padding:11px 14px; border-radius:10px; color:#0051A5; background:#EAF3FF;
    border-left:3px solid #0051A5; font-size:.86rem; font-weight:650; margin-bottom:1.5rem; }
.cs-eyebrow { color:#5D7691; font-size:.67rem; font-weight:750; letter-spacing:.15em;
    text-transform:uppercase; margin-bottom:9px; }
.cs-side-note { background:#F3F7FC; border:1px solid #E3EBF5; border-radius:11px;
    padding:13px; font-size:.78rem; line-height:1.6; color:#536B83; }
.cs-side-note strong { color:#002750; }
.cs-topline { display:flex; align-items:center; justify-content:space-between;
    gap:12px; flex-wrap:wrap; margin-bottom:18px; }
.cs-breadcrumb { font-size:.77rem; color:#697F95; }
.cs-breadcrumb strong { color:#002750; font-weight:650; }
.cs-pill { display:inline-flex; align-items:center; gap:7px; font-size:.71rem; font-weight:650;
    padding:6px 11px; border-radius:30px; background:#E7F3EC; color:#256047; white-space:nowrap; }
.cs-pill:before { content:""; width:6px; height:6px; border-radius:50%; background:currentColor; }
.cs-pill.pending { background:#FFF5D7; color:#7B6115; }
.cs-hero { position:relative; overflow:hidden; border-radius:20px;
    background:linear-gradient(112deg,#002750 0%,#003D7A 60%,#0051A5 100%);
    padding:34px 37px; margin-bottom:22px; border-bottom:4px solid #FEDF01;
    min-height:215px; }
.cs-hero-copy { position:relative; z-index:1; width:70%; }
.cs-hero .cs-eyebrow { color:#FEDF01; margin-bottom:12px; }
.cs-hero h1 { color:#FFF; font-size:2.45rem; font-weight:700; line-height:1.14;
    letter-spacing:-.045em; margin:0 0 12px; padding:0; }
.cs-hero p { color:#D6E7FC; font-size:.91rem; line-height:1.65; margin:0; max-width:570px; }
.cs-hero-art { position:absolute; right:-10px; top:-25px; width:340px; height:290px; opacity:.85; }
.cs-chips { display:flex; flex-wrap:wrap; gap:7px; margin-top:19px; }
.cs-chip { font-size:.65rem; font-weight:600; color:#D8E9FD;
    border:1px solid rgba(216,233,253,.25); padding:4px 9px; border-radius:30px; }
.cs-section { margin:22px 0 13px; display:flex; justify-content:space-between; align-items:center; gap:12px; }
.cs-section h2 { margin:0; padding:0; font-size:1.18rem; font-weight:700; color:#002750; }
.cs-section span { font-size:.71rem; color:#60788F; }
.cs-card { background:#FFF; border:1px solid #DFE7F1; border-radius:15px; padding:21px; }
.cs-card h3 { margin:0 0 9px; padding:0; font-size:1.05rem !important; }
.cs-card p { color:#60758B; font-size:.83rem; margin:0 0 10px; }
.cs-step { display:flex; gap:12px; margin-top:17px; }
.cs-step-no { display:flex; align-items:center; justify-content:center; width:27px; height:27px;
    background:#EAF3FF; color:#0051A5; border-radius:8px; font-size:.76rem; font-weight:750; flex-shrink:0; }
.cs-step strong { display:block; font-size:.83rem; color:#002750; margin:0 0 2px; }
.cs-step p { font-size:.77rem; line-height:1.5; margin:0; }
.cs-empty { border:1px dashed #C7D7E8; border-radius:16px; padding:30px 24px;
    text-align:center; background:#FFF; margin-top:9px; }
.cs-empty svg { margin:0 auto 8px; }
.cs-empty h3 { font-size:1.05rem !important; margin:0 0 5px; padding:0; }
.cs-empty p { font-size:.81rem; color:#60758B; margin:0; }
.cs-kpi { background:#FFF; border:1px solid #DFE7F1; border-radius:13px;
    padding:17px 19px; margin:4px 0 18px; border-top:3px solid #0051A5; min-height:120px; }
.cs-kpi.gold { border-top-color:#E4C600; }
.cs-mean-card { background:#002750; border-radius:16px; padding:20px 24px; color:white; margin:12px 0 18px; }
.cs-mean-card strong { color:#FEDF01; font-size:2.1rem; display:block; margin:4px 0; }
.cs-mean-card .lead { color:#BFD8F6; font-size:.76rem; text-transform:uppercase; letter-spacing:.07em; }
.cs-mean-card .detail { color:#D9E9FC; font-size:.8rem; }
.cs-kpi-label { font-size:.74rem; color:#536B83; font-weight:600; }
.cs-kpi-value { font-size:2rem; font-weight:750; color:#002750; line-height:1.35; letter-spacing:-.04em; }
.cs-kpi-detail { font-size:.66rem; color:#6A7E93; }
.cs-meta { display:flex; gap:8px; flex-wrap:wrap; margin:9px 0 15px; }
.cs-meta span { background:#E9F0F9; border-radius:6px; padding:5px 8px; font-size:.7rem; color:#3C5A78; }
.cs-summary-row { display:flex; justify-content:space-between; gap:10px; margin-top:13px;
    font-size:.79rem; color:#18334E; }
.cs-summary-row strong { color:#0051A5; font-variant-numeric:tabular-nums; }
.cs-bar { height:7px; background:#E8EFF8; border-radius:6px; margin-top:6px; overflow:hidden; }
.cs-bar-fill { height:100%; background:#0051A5; border-radius:6px; }
.cs-notice { color:#526B84; font-size:.77rem; background:#EBF2FA; border-left:3px solid #0051A5;
    padding:11px 14px; border-radius:0 8px 8px 0; margin:12px 0; line-height:1.6; }
.cs-notice.gold { background:#FFF9DF; border-left-color:#D9BD00; color:#695718; }
.cs-table-wrap { overflow-x:auto; border:1px solid #DFE7F1; border-radius:12px; background:white; }
.cs-table { border-collapse:collapse; width:100%; text-align:left; font-size:.78rem; }
.cs-table th { background:#ECF3FC; color:#264563; font-size:.7rem; letter-spacing:.02em;
    padding:12px; font-weight:650; white-space:nowrap; }
.cs-table td { padding:11px 12px; border-top:1px solid #EDF1F7; color:#3C5670; }
.cs-table tr:nth-child(even) td { background:#FBFCFE; }
.cs-flag { color:#866012; background:#FFF3CE; padding:3px 7px; border-radius:5px; font-size:.68rem; }
.cs-footer { border-top:1px solid #DFE7F1; padding-top:15px; margin-top:32px; color:#698097;
    font-size:.7rem; display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; }
.cs-footer strong { color:#3C5A78; }
@media (max-width:850px) {
    .block-container { padding:1.2rem 1rem; }
    .cs-hero { padding:26px 23px; }
    .cs-hero h1 { font-size:2rem; }
    .cs-hero-copy { width:90%; }
    .cs-hero-art { opacity:.22; right:-100px; }
    .cs-section span { display:none; }
}
</style>
"""

FISH_ICON = """<svg width="28" height="28" viewBox="0 0 32 32" fill="none" aria-hidden="true">
<path d="M4 16c5-8 13-8 19 0-6 8-14 8-19 0Z" stroke="#FEDF01" stroke-width="1.8"/>
<path d="m23 16 6-5v10l-6-5Z" fill="#FEDF01"/><circle cx="10" cy="15" r="1.5" fill="#FEDF01"/>
</svg>"""

SONAR_ART = """<svg class="cs-hero-art" viewBox="0 0 340 290" fill="none" aria-hidden="true">
<circle cx="205" cy="146" r="114" stroke="#639BD5" stroke-opacity=".3"/>
<circle cx="205" cy="146" r="83" stroke="#639BD5" stroke-opacity=".35"/>
<circle cx="205" cy="146" r="51" stroke="#639BD5" stroke-opacity=".4"/>
<path d="M205 20v251M78 146h254" stroke="#639BD5" stroke-opacity=".22" stroke-dasharray="4 6"/>
<path d="M119 141c24-33 66-33 97 0-31 33-73 33-97 0Z" fill="#73B0E3" fill-opacity=".14" stroke="#AFD4FA" stroke-width="1.6"/>
<path d="m216 141 24-19v38l-24-19Z" stroke="#AFD4FA" stroke-width="1.6"/>
<circle cx="142" cy="137" r="3" fill="#FEDF01"/>
<path d="M110 111v-12h16m105 0h16v12m0 60v12h-16m-105 0h-16v-12" stroke="#FEDF01" stroke-width="2"/>
<circle cx="278" cy="89" r="4" fill="#FEDF01"/><circle cx="108" cy="203" r="3" fill="#73B0E3"/>
<path d="M7 238c48-38 75 35 128 0s87 30 143-1 51-10 62-6" stroke="#639BD5" stroke-opacity=".3"/>
</svg>"""


def clahe_enhance(bgr):
    """Apply the original CLAHE rule to luminance only."""
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB)
    lightness, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return cv2.cvtColor(
        cv2.merge((clahe.apply(lightness), a, b)), cv2.COLOR_LAB2BGR
    )


def get_visibility_metrics(bgr):
    """Brightness/contrast heuristics; these are not turbidity measurements."""
    h, w = bgr.shape[:2]
    scale = min(1.0, 320.0 / max(h, w))
    sample = (
        cv2.resize(bgr, (max(1, int(w * scale)), max(1, int(h * scale))))
        if scale < 1 else bgr
    )
    lightness = cv2.cvtColor(sample, cv2.COLOR_BGR2LAB)[:, :, 0]
    p10, p90 = np.percentile(lightness, (10, 90))
    return {
        "brightness": round(float(np.mean(lightness)), 1),
        "contrast": round(float(p90 - p10), 1),
    }


def prepare_frame(bgr, mode, contrast_cutoff):
    metrics = get_visibility_metrics(bgr)
    suggested = (
        metrics["contrast"] < contrast_cutoff
        or (metrics["brightness"] < 55 and metrics["contrast"] < 95)
    )
    apply = mode == "Always" or (mode == "Automatic" and suggested)
    output = clahe_enhance(bgr) if apply else bgr.copy()
    return output, apply, metrics, suggested


def run_yolov7(source, output_root, image_size, confidence):
    """Retain the existing detector entry point and inference arguments."""
    command = [
        sys.executable, str(PROJECT / "detect.py"),
        "--weights", str(MODEL), "--source", str(source),
        "--img-size", str(image_size), "--device", "cpu", "--no-trace",
        "--conf-thres", f"{confidence:.2f}", "--iou-thres", "0.65",
        "--save-txt", "--save-conf", "--project", str(output_root),
        "--name", "demo", "--exist-ok",
    ]
    return subprocess.run(
        command, cwd=str(PROJECT), capture_output=True, text=True,
        errors="replace", timeout=240,
    )


def read_detection_rows(result_dir, frame_metadata):
    rows = []
    for stem, metadata in frame_metadata.items():
        label_file = result_dir / "labels" / (stem + ".txt")
        if not label_file.exists():
            continue
        for line in label_file.read_text(encoding="utf-8").splitlines():
            parts = line.split()
            if len(parts) < 6:
                continue
            try:
                raw_class = float(parts[0])
                xc, yc, bw, bh, confidence = (float(x) for x in parts[1:6])
                if not all(math.isfinite(x) for x in (raw_class, xc, yc, bw, bh, confidence)):
                    continue
                if not raw_class.is_integer() or not 0 <= confidence <= 1:
                    continue
                class_id = int(raw_class)
            except (ValueError, OverflowError):
                continue
            name = CLASS_NAMES[class_id] if 0 <= class_id < len(CLASS_NAMES) else f"Class {class_id}"
            rows.append({
                "Frame": stem,
                "Time (s)": metadata["time"],
                "Predicted species": name,
                "Confidence": round(confidence, 3),
                "Review flag": "Low confidence" if confidence < LOW_CONFIDENCE else "Preliminary",
                "CLAHE applied": "Yes" if metadata["enhanced"] else "No",
                "Box x": round(xc, 4), "Box y": round(yc, 4),
                "Box w": round(bw, 4), "Box h": round(bh, 4),
            })
    return rows


def calculate_sample_counts(rows, frame_metadata):
    """Count detections per sampled still frame, INCLUDING frames with zero detections.

    MeanCount is an observational relative-abundance metric. Here it is only a
    preliminary model-generated approximation for staff to verify.
    """
    by_frame = {}
    for row in rows:
        by_frame.setdefault(row["Frame"], []).append(row)
    samples = []
    for index, (frame, metadata) in enumerate(frame_metadata.items(), start=1):
        detections = by_frame.get(frame, [])
        species_counts = Counter(d["Predicted species"] for d in detections)
        fish = sum(n for species, n in species_counts.items() if species not in NON_FISH_CLASSES)
        non_fish = sum(n for species, n in species_counts.items() if species in NON_FISH_CLASSES)
        samples.append({
            "Sample": index,
            "Frame": frame,
            "Time (s)": metadata["time"],
            "Fish detections": fish,
            "Other detected organisms": non_fish,
            "Low-confidence flags": sum(d["Review flag"] == "Low confidence" for d in detections),
            "CLAHE applied": "Yes" if metadata["enhanced"] else "No",
            "species_counts": dict(species_counts),
        })
    n = len(samples)
    species_means = {
        name: (sum(sample["species_counts"].get(name, 0) for sample in samples) / n if n else 0.0)
        for name in CLASS_NAMES
    }
    return {
        "samples": samples,
        "mean_fish": sum(s["Fish detections"] for s in samples) / n if n else 0.0,
        "mean_other": sum(s["Other detected organisms"] for s in samples) / n if n else 0.0,
        "species_means": species_means,
    }




def parse_species_lines(value):
    """Parse optional manual per-species counts, one 'Species: integer' per line."""
    parsed = {}
    for number, raw in enumerate(value.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        if ":" not in line:
            raise ValueError(f"Line {number} must look like 'Atlantic cod: 3'.")
        name, count_raw = line.rsplit(":", 1)
        name = name.strip()
        if not name or len(name) > 120:
            raise ValueError(f"Line {number} needs a short species name.")
        try:
            count = int(count_raw.strip())
        except ValueError as exc:
            raise ValueError(f"Line {number} needs a whole-number fish count.") from exc
        if count < 0:
            raise ValueError(f"Line {number} cannot have a negative count.")
        if name.casefold() in (n.casefold() for n in parsed):
            raise ValueError(f"The species '{name}' appears more than once.")
        if name.casefold() in ("unknown", "unknown fish", "unidentified", "unidentified fish"):
            raise ValueError("Enter unidentified fish in the separate field instead.")
        if name.casefold() == "whelk":
            raise ValueError("Whelk is not a fish; include it in invertebrate counts instead.")
        parsed[name] = count
    return parsed


def review_statistics(result, review_state):
    """An adjusted mean uses ONLY explicitly edited sample counts.

    It is NOT a verified mean or automatic approval of untouched observations.
    """
    samples = result.get("sample_counts", {}).get("samples", [])
    saved = (review_state or {}).get("samples", {})
    n = len(samples)
    adjusted_counts = [int(saved[str(s["Sample"])]["fish"])
                       if str(s["Sample"]) in saved else int(s["Fish detections"])
                       for s in samples]
    # Metadata-only updates are not presented as changed fish counts.
    edited = sum(str(s["Sample"]) in saved and
                 int(saved[str(s["Sample"])]["fish"]) != int(s["Fish detections"])
                 for s in samples)
    return {
        "edited_count": edited, "sample_count": n,
        "adjusted_mean": sum(adjusted_counts) / n if n else 0.0,
        "full_window_reviewed": bool((review_state or {}).get("full_window", {}).get("reviewed")),
    }


def fresh_review_state(result):
    return {"run_id": result.get("run_id"), "samples": {}, "full_window": {}}


def safe_video_player(data, start_s=0, end_s=None):
    """Use native MP4 playback at original frame rate; limit end when Streamlit supports it."""
    kwargs = {"start_time": max(0, int(start_s))}
    has_end_time = "end_time" in inspect.signature(st.video).parameters
    if has_end_time and end_s is not None:
        kwargs["end_time"] = max(kwargs["start_time"] + 1, int(end_s))
    st.video(data, **kwargs)
    if not has_end_time and end_s is not None:
        st.caption(f"Your Streamlit version cannot stop playback automatically. Pause the video at "
                   f"{int(end_s) // 60:02}:{int(end_s) % 60:02} and use the video timeline to replay.")


def timestamp_label(seconds):
    sec = int(float(seconds))
    return f"{sec // 60:02d}:{sec % 60:02d}"


def render_researcher_review(result, source_video, state):
    """Optional correction workflow. Nothing is mandatory before seeing/exporting AI results."""
    if not result.get("is_video"):
        return
    samples = result["sample_counts"]["samples"]
    saved = state.get("samples", {})
    stats = review_statistics(result, state)

    # Requiring verification of all 15 AI guesses would negate this prototype's purpose.
    # Corrections update the mean, but never silently become a certified result.
    with st.expander("Correct a count", expanded=False):
        n = st.selectbox(
            "Which sample needs a correction?",
            options=list(range(len(samples))),
            format_func=lambda i: (
                f"Sample {i+1:02d} · {timestamp_label(samples[i]['Time (s)'])}"
                + (" · Edited" if str(i+1) in saved and int(saved[str(i+1)]["fish"]) != int(samples[i]["Fish detections"]) else "")
            ),
            key=f"cs_simple_sample_picker_{result['run_id']}",
        )
        sample = samples[n]
        key = str(sample["Sample"])
        previous = saved.get(key)
        timestamp = float(sample["Time (s)"])
        st.caption(f"AI counted {sample['Fish detections']} fish. You can inspect the original video, "
                   "then change the count if necessary.")
        left, right = st.columns([1.1, 1], gap="large")
        with left:
            mode = st.radio("Video view", ["One-second observation", "Surrounding ±5 seconds"],
                            horizontal=True, key=f"cs_simple_video_view_{result['run_id']}")
            if mode == "One-second observation":
                safe_video_player(source_video, timestamp, timestamp+1)
            else:
                safe_video_player(source_video, max(0, timestamp-5), timestamp+6)
            annotated = result.get("sample_images", {}).get(sample["Frame"])
            if annotated is not None:
                st.image(annotated, caption="YOLOv7 detection at the sampled timestamp",
                         use_container_width=True)
        with right:
            with st.form(f"cs_simple_correction_form_{result['run_id']}_{key}"):
                fish = st.number_input("Correct fish count", min_value=0, max_value=100000,
                                       value=int((previous or {}).get("fish", sample["Fish detections"])), step=1)
                notes = st.text_input("Reason or note (optional)",
                                      value=(previous or {}).get("notes", ""),
                                      placeholder="e.g. two cod were missed")
                with st.expander("Other observations (optional)"):
                    other = st.number_input("Other organisms / invertebrates", min_value=0, max_value=100000,
                                            value=int((previous or {}).get("invertebrates", sample["Other detected organisms"])),
                                            step=1)
                    unknown = st.number_input("Fish of unknown species", min_value=0, max_value=100000,
                                              value=int((previous or {}).get("unknown_fish", 0)), step=1)
                    vis_options = ["Not assessed", "<0.5 m", "0.5–1.0 m", "1.0–1.5 m", "≥1.5 m"]
                    prev_visibility = (previous or {}).get("visibility", "Not assessed")
                    visibility = st.selectbox("Horizontal visibility", vis_options,
                                               index=vis_options.index(prev_visibility) if prev_visibility in vis_options else 0)
                save = st.form_submit_button("Save correction", type="primary", use_container_width=True)
            if save:
                if int(unknown) > int(fish):
                    st.error("Unknown fish cannot exceed the corrected total fish count.")
                else:
                    state["samples"][key] = {
                        "fish": int(fish), "invertebrates": int(other),
                        "unknown_fish": int(unknown), "visibility": visibility,
                        "notes": notes.strip(),
                        "edited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
                    }
                    st.session_state[REVIEW_KEY] = state
                    st.rerun()
            if previous is not None:
                if st.button("Undo this correction", key=f"cs_simple_undo_{result['run_id']}_{key}"):
                    state["samples"].pop(key, None)
                    st.session_state[REVIEW_KEY] = state
                    st.rerun()
        st.caption("Saving a correction updates the adjusted mean. It does not certify the "
                   "other samples or replace Coastal Action's required review and quality control.")

    with st.expander("Full-video species notes (optional)", expanded=False):
        st.caption("The draft protocol also requires a person to review the whole analysis window "
                   "for species that may appear between sampled timestamps. This is not automated here.")
        start = int(result.get("start_second", 0))
        duration = (PROTOCOL_WINDOW_SECONDS if result.get("sampling_plan") == "protocol"
                    else int(result.get("segment_seconds", 0)))
        if st.checkbox("Show the full survey video", key=f"cs_simple_full_video_{result['run_id']}"):
            safe_video_player(source_video, start, start + duration)
        full = state.get("full_window", {})
        with st.form(f"cs_simple_full_form_{result['run_id']}"):
            species_text = st.text_area("Species noticed across the recording (one per line)",
                                        value="\n".join(full.get("species", [])),
                                        placeholder="Atlantic cod\nThorny skate", height=90)
            checked = st.checkbox("I watched the full analysis window", value=bool(full.get("reviewed", False)))
            save_full = st.form_submit_button("Save species notes")
        if save_full:
            species = list(dict.fromkeys(s.strip() for s in species_text.splitlines() if s.strip()))
            state["full_window"] = {"reviewed": bool(checked), "species": species}
            st.session_state[REVIEW_KEY] = state
            st.rerun()
        if full.get("species"):
            st.caption("Saved species notes: " + ", ".join(full["species"]))


def excel_bytes(rows, survey=None, review_state=None):
    """Simple Excel: one overview, one straightforward sample/observation sheet.

    Values never get the label 'verified': a user editing one count doesn't prove
    that all other sampled observations passed Coastal Action's checks.
    """
    survey = survey or {}
    review_state = review_state or {}
    is_video = bool(survey.get("is_video"))
    samples = survey.get("sample_counts", {}).get("samples", []) if is_video else []
    stats = review_statistics(survey, review_state) if is_video else {}
    edits = review_state.get("samples", {})
    full_window = review_state.get("full_window", {})
    buffer = io.BytesIO()
    with xlsxwriter.Workbook(buffer, {
        "in_memory": True, "strings_to_formulas": False, "strings_to_urls": False
    }) as book:
        title = book.add_format({"font_name": "Aptos", "bold": True, "font_size": 21,
                                 "font_color": "#FFFFFF", "bg_color": NAVY, "valign": "vcenter"})
        bar = book.add_format({"bg_color": YELLOW})
        heading = book.add_format({"bold": True, "font_color": "#FFFFFF",
                                   "bg_color": BLUE, "valign": "vcenter"})
        label = book.add_format({"bold": True, "font_color": NAVY, "bg_color": "#EDF3FA"})
        number = book.add_format({"num_format": "0.00", "font_color": NAVY})
        note = book.add_format({"font_color": "#594A20", "bg_color": "#FFF7D9",
                                "text_wrap": True, "valign": "vcenter"})
        edited_fmt = book.add_format({"font_color": "#0051A5", "bold": True,
                                      "bg_color": "#EAF3FF"})
        normal = book.add_format({"font_color": "#284563"})
        summary = book.add_worksheet("Overview")
        summary.hide_gridlines(2)
        summary.set_column("A:A", 28)
        summary.set_column("B:B", 53)
        summary.set_column("C:D", 18)
        summary.merge_range("A1:D2", "CoastalSense | Survey overview", title)
        summary.set_row(0, 24)
        summary.merge_range("A3:D3", "", bar)
        values = [
            ("Recording", str(survey.get("filename", ""))),
            ("Date (UTC)", str(survey.get("created", ""))),
            ("Analysis", "15-minute draft protocol" if survey.get("sampling_plan") == "protocol"
             else "Quick video demonstration" if is_video else "Photograph"),
            ("Analyzed samples", len(samples) if is_video else survey.get("frame_count", 0)),
            ("AI detections", len(rows)),
        ]
        if is_video:
            values += [
                ("AI MeanCount / sampled frame", round(survey["sample_counts"]["mean_fish"], 2)),
                ("Adjusted mean (with edits)", round(stats["adjusted_mean"], 2)),
                ("Fish counts corrected", stats["edited_count"]),
                ("Full-window species review", "Recorded" if stats["full_window_reviewed"] else "Not marked complete"),
                ("Species noted by researcher", ", ".join(full_window.get("species", [])) or "None recorded"),
            ]
        for row_index, (k, value) in enumerate(values, start=4):
            summary.write_string(row_index, 0, k, label)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                summary.write_number(row_index, 1, value, number if isinstance(value, float) else normal)
            else:
                summary.write_string(row_index, 1, str(value), normal)
            summary.set_row(row_index, 24)
        note_index = len(values) + 6
        summary.merge_range(note_index, 0, note_index + 2, 3,
            "PRELIMINARY: AI counts may miss or misidentify fish. 'Adjusted mean' uses only corrections "
            "entered by a person; unchanged samples still use AI counts. Neither figure is a verified "
            "scientific MeanCount without the required full observation review and independent quality control.", note)
        summary.set_row(note_index, 25)
        summary.set_row(note_index+1, 25)
        summary.set_row(note_index+2, 25)

        if is_video:
            sheet = book.add_worksheet("Sample counts")
            sheet.hide_gridlines(2)
            sheet.freeze_panes(1, 2)
            widths = [10, 13, 13, 17, 15, 33, 16, 16, 12, 30]
            for j, w in enumerate(widths):
                sheet.set_column(j, j, w)
            columns = ["Sample", "Time (s)", "AI detected fish", "Current fish", "Count corrected?",
                       "AI species suggestions", "Other organisms", "Visibility", "Unknown fish", "Notes"]
            sheet.write_row(0, 0, columns, heading)
            sheet.set_row(0, 30)
            for idx, sample in enumerate(samples, start=1):
                saved = edits.get(str(sample["Sample"]))
                fish = int(saved["fish"]) if saved else int(sample["Fish detections"])
                raw_species = sample.get("species_counts", {})
                species = ", ".join(f"{key}: {val}" for key, val in raw_species.items() if val)
                changed = bool(saved and fish != int(sample["Fish detections"]))
                values_row = [sample["Sample"], sample["Time (s)"], sample["Fish detections"],
                              fish, "Yes" if changed else "No", species or "None detected",
                              int(saved.get("invertebrates", sample["Other detected organisms"])) if saved
                              else int(sample["Other detected organisms"]),
                              saved.get("visibility", "") if saved else "",
                              int(saved.get("unknown_fish", 0)) if saved else "",
                              saved.get("notes", "") if saved else ""]
                for j, value in enumerate(values_row):
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        sheet.write_number(idx, j, value, number if j == 1 else None)
                    else:
                        sheet.write_string(idx, j, str(value), edited_fmt if j == 4 and changed else None)
            sheet.autofilter(0, 0, max(1, len(samples)), len(columns)-1)
            sheet.set_landscape()
            sheet.fit_to_pages(1, 0)
        else:
            sheet = book.add_worksheet("Detections")
            sheet.hide_gridlines(2)
            sheet.freeze_panes(1, 0)
            for j, width in enumerate([19, 25, 15, 18, 17]):
                sheet.set_column(j, j, width)
            columns = ["Frame", "Predicted category", "Confidence", "Review flag", "CLAHE applied"]
            sheet.write_row(0, 0, columns, heading)
            for idx, row in enumerate(rows, start=1):
                for j, key in enumerate(["Frame", "Predicted species", "Confidence", "Review flag", "CLAHE applied"]):
                    value = row.get(key, "")
                    if key == "Confidence":
                        sheet.write_number(idx, j, float(value), number)
                    else:
                        sheet.write_string(idx, j, str(value))
            sheet.autofilter(0, 0, max(1,len(rows)), len(columns)-1)
    return buffer.getvalue()


class DetectionError(RuntimeError):
    def __init__(self, message, log=""):
        super().__init__(message)
        self.log = log


def analyze_upload(uploaded, filename, is_video, mode, confidence, contrast_cutoff,
                   start_second=0, segment_seconds=3, sampling_plan="quick", progress=None):
    """Run once per explicit click; return only portable results, never temp paths."""
    started = time.perf_counter()

    def report(value, message):
        if progress:
            progress(value, message)

    report(5, "Preparing your survey footage…")
    with tempfile.TemporaryDirectory(prefix="coastalsense_") as working_folder:
        work = Path(working_folder)
        output_root = work / "output"
        frame_metadata = {}
        enhanced_count = 0
        suggested_count = 0
        preview = None
        if not is_video:
            original = cv2.imdecode(np.frombuffer(uploaded, dtype=np.uint8), cv2.IMREAD_COLOR)
            if original is None:
                raise DetectionError("This image could not be read. Choose another JPG or PNG.")
            treated, applied, metrics, suggested = prepare_frame(original, mode, contrast_cutoff)
            source = work / "input.jpg"
            if not cv2.imwrite(str(source), treated):
                raise DetectionError("Could not prepare the image for detection.")
            frame_metadata[source.stem] = {"time": "—", "enhanced": applied}
            frame_names = [source.name]
            enhanced_count = int(applied)
            suggested_count = int(suggested)
            preview = (original, treated, metrics, applied, suggested)
            image_size = 640
        else:
            upload_path = work / "upload.mp4"
            upload_path.write_bytes(uploaded)
            cap = cv2.VideoCapture(str(upload_path))
            source = work / "frames"
            source.mkdir()
            frame_names = []
            try:
                if not cap.isOpened():
                    raise DetectionError("This MP4 could not be read. Try a different MP4 export.")
                if sampling_plan == "protocol":
                    # Main portion of Coastal Action's DRAFT protocol: 15 sampled
                    # points, one per minute, over a chosen 15-minute window.
                    fps = cap.get(cv2.CAP_PROP_FPS)
                    frame_total = cap.get(cv2.CAP_PROP_FRAME_COUNT)
                    if fps > 0 and frame_total > 0:
                        duration_s = frame_total / fps
                        if duration_s + 0.25 < start_second + PROTOCOL_WINDOW_SECONDS:
                            raise DetectionError(
                                f"The draft protocol needs a full 15-minute window after "
                                f"second {start_second}. This video is about {duration_s:.1f}s. "
                                "Switch to Quick demo for short footage."
                            )
                    sample_times = [float(start_second + index * PROTOCOL_INTERVAL_SECONDS)
                                    for index in range(PROTOCOL_SAMPLES)]
                else:
                    sample_times = [float(start_second) + index / SAMPLE_RATE
                                    for index in range(int(segment_seconds) * SAMPLE_RATE)]
                sample_count = len(sample_times)
                for index, second in enumerate(sample_times):
                    cap.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
                    success, bgr = cap.read()
                    if not success:
                        break
                    h, w = bgr.shape[:2]
                    scale = min(1.0, 640.0 / max(h, w))
                    if scale < 1:
                        bgr = cv2.resize(bgr, (max(1, round(w * scale)), max(1, round(h * scale))))
                    treated, applied, metrics, suggested = prepare_frame(bgr, mode, contrast_cutoff)
                    if preview is None:
                        preview = (bgr.copy(), treated.copy(), metrics, applied, suggested)
                    stem = f"frame_{index:03}"
                    name = stem + ".jpg"
                    if not cv2.imwrite(str(source / name), treated):
                        raise DetectionError("Could not prepare a video frame for detection.")
                    frame_names.append(name)
                    frame_metadata[stem] = {"time": round(second, 2), "enhanced": applied}
                    enhanced_count += int(applied)
                    suggested_count += int(suggested)
                    report(5 + int(25 * (index + 1) / sample_count),
                           f"Preparing sample {index + 1} of up to {sample_count}…")
            finally:
                cap.release()
            if not frame_names:
                raise DetectionError("No readable frames in that segment. Try starting at second 0.")
            if sampling_plan == "protocol" and len(frame_names) != PROTOCOL_SAMPLES:
                raise DetectionError("The draft protocol requires all 15 sampled timestamps. "
                                     "Some video frames could not be decoded. Try a different recording.")
            image_size = 416

        report(35, "Running YOLOv7 on your laptop’s CPU. This can take a few minutes…")
        process = run_yolov7(source, output_root, image_size, confidence)
        log = process.stdout + "\n" + process.stderr
        if process.returncode != 0:
            raise DetectionError("YOLOv7 could not complete detection. Open the technical details below.", log)
        result_dir = output_root / "demo"
        outputs = [result_dir / name for name in frame_names]
        if not all(path.is_file() for path in outputs):
            raise DetectionError("Detection finished, but the expected annotated frames were not saved.", log)
        report(85, "Preparing the preview and your observation workbook…")
        rows = read_detection_rows(result_dir, frame_metadata)
        result = {
            "filename": filename, "is_video": is_video,
            "mode": mode, "confidence": confidence, "contrast_cutoff": contrast_cutoff,
            "start_second": int(start_second), "segment_seconds": int(segment_seconds),
            "sampling_plan": sampling_plan,
            "run_id": str(time.time_ns()),
            "frame_count": len(frame_names), "enhanced_count": enhanced_count,
            "suggested_count": suggested_count, "rows": rows, "preview": preview,
            "log": log, "created": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        }
        if is_video:
            result["sample_counts"] = calculate_sample_counts(rows, frame_metadata)
            result["sample_images"] = {}
            frames = []
            for path in outputs:
                result["sample_images"][path.stem] = path.read_bytes()
                with Image.open(path) as image:
                    frames.append(image.convert("RGB").copy())
            gif = io.BytesIO()
            frames[0].save(gif, format="GIF", save_all=True, append_images=frames[1:],
                           duration=500 if sampling_plan == "quick" else 800,
                           loop=0, optimize=False)
            result["gif"] = gif.getvalue()
        else:
            result["image"] = outputs[0].read_bytes()
        result["elapsed"] = round(time.perf_counter() - started, 1)
        result["excel"] = excel_bytes(rows, result)
        report(100, "Analysis complete. Your results are ready.")
        return result


def section(title, subtitle=""):
    st.markdown(f'<div class="cs-section"><h2>{escape(title)}</h2><span>{escape(subtitle)}</span></div>',
                unsafe_allow_html=True)


def notice(text, gold=False):
    st.markdown(f'<div class="cs-notice{" gold" if gold else ""}">{escape(text)}</div>',
                unsafe_allow_html=True)


def render_sidebar(model_ready):
    with st.sidebar:
        st.markdown(f'<div class="cs-brand"><div class="cs-mark">{FISH_ICON}</div><div>'
                    '<div class="cs-brand-name">CoastalSense</div>'
                    '<div class="cs-brand-sub">MARINE MONITORING</div></div></div>'
                    '<div class="cs-nav">◉ &nbsp; Survey workspace</div>', unsafe_allow_html=True)
        st.markdown('<div class="cs-eyebrow">Survey setup</div>', unsafe_allow_html=True)
        st.subheader("Image enhancement")
        mode = st.radio("CLAHE mode", ["Automatic", "Always", "Off"], key="cs_rbc_mode",
                        help="Automatic uses a preliminary brightness/contrast rule. Always and Off allow comparison.")
        descriptions = {
            "Automatic": "Enhance frames when the brightness and contrast rule suggests it.",
            "Always": "Apply contrast enhancement to every analyzed frame.",
            "Off": "Run detection using the original frames without enhancement.",
        }
        st.caption(descriptions[mode])
        with st.expander("Advanced controls"):
            confidence = st.slider("Detection confidence threshold", 0.10, 0.50, 0.10, 0.05,
                key="cs_rbc_confidence", help="Lower values reveal weaker detections and may increase false positives.")
            cutoff = st.slider("Automatic contrast cutoff", 40, 100, 65, 5,
                key="cs_rbc_cutoff", help="A higher value causes Automatic mode to enhance more frames. This threshold is provisional.")
            st.caption("Scores below 25% receive a low-confidence flag. All predictions need review.")
        st.divider()
        st.markdown('<div class="cs-eyebrow">Processing</div>', unsafe_allow_html=True)
        state = "Model files found" if model_ready else "Model files needed"
        st.markdown(f'<div class="cs-side-note"><strong>{state}</strong><br>'
                    'Local YOLOv7 · CPU<br>Quick demo · 2 frames / second<br>'
                    'Draft protocol · 1 sample / minute<br>'
                    'Photo input · 640 px inference<br>Video input · 416 px inference</div>', unsafe_allow_html=True)
        st.divider()
        st.markdown('<div class="cs-side-note"><strong>Detection prototype</strong><br>'
                    'Repeated sightings across frames are not unique fish counts. '
                    'Identity tracking is a future capability.</div>', unsafe_allow_html=True)
        st.caption("Prepared for RBC HubHacks 2026")
    return mode, confidence, cutoff


def render_hero(model_ready):
    status = "Local model files found" if model_ready else "Add your model files to analyze"
    status_class = "" if model_ready else " pending"
    st.markdown('<div class="cs-topline"><div class="cs-breadcrumb">CoastalSense &nbsp; / &nbsp; '
                f'<strong>Survey workspace</strong></div><span class="cs-pill{status_class}">{status}</span></div>'
                '<div class="cs-hero"><div class="cs-hero-copy">'
                '<div class="cs-eyebrow">COASTALSENSE &nbsp; / &nbsp; RBC HUBHACKS 2026</div>'
                '<h1>A clearer view<br>beneath the surface.</h1>'
                '<p>Enhance underwater footage, identify marine species, and turn '
                'survey observations into evidence you can review.</p>'
                '<div class="cs-chips"><span class="cs-chip">Contrast enhancement</span>'
                '<span class="cs-chip">Species detection</span><span class="cs-chip">Excel observations</span>'
                f'</div></div>{SONAR_ART}</div>', unsafe_allow_html=True)


def render_steps():
    steps = [
        ("01", "Upload your survey", "Choose an underwater photograph or a short segment of video."),
        ("02", "Enhance and detect", "Review the enhancement mode, then run the local fish detector."),
        ("03", "Review and export", "Inspect preliminary predictions and download your observations."),
    ]
    content = '<div class="cs-card"><div class="cs-eyebrow">From footage to observations</div><h3>Your survey, in three steps</h3>'
    for number, title, text in steps:
        content += (f'<div class="cs-step"><div class="cs-step-no">{number}</div><div>'
                    f'<strong>{title}</strong><p>{text}</p></div></div>')
    st.markdown(content + '</div>', unsafe_allow_html=True)


def show_gif(data, caption):
    # An explicit data URL ensures the browser animates a GIF on older Streamlit
    # versions too; st.image may convert animated inputs into a static frame.
    encoded = base64.b64encode(data).decode("ascii")
    st.markdown(f'<img src="data:image/gif;base64,{encoded}" alt="Animated sampled fish detections" '
                'style="display:block;width:100%;border-radius:13px;">', unsafe_allow_html=True)
    st.caption(caption)


def render_species_summary(rows):
    counts = Counter(row["Predicted species"] for row in rows)
    total = len(rows)
    content = '<div class="cs-card"><div class="cs-eyebrow">Observed categories</div><h3>Species snapshot</h3>'
    if not counts:
        content += '<p>No detections at the selected threshold. This does not confirm that no fish are present.</p>'
    for species, count in counts.most_common():
        width = 100 * count / total
        content += (f'<div class="cs-summary-row"><span>{escape(species)}</span><strong>{count}</strong></div>'
                    f'<div class="cs-bar"><div class="cs-bar-fill" style="width:{width:.1f}%"></div></div>')
    content += '<p style="font-size:.7rem;margin-top:17px;margin-bottom:0;">Counts represent detections across frames.</p></div>'
    st.markdown(content, unsafe_allow_html=True)


def render_comparison(result):
    original, treated, metrics, applied, suggested = result["preview"]
    left, right = st.columns(2)
    left.image(original, channels="BGR", caption="Original sample", use_container_width=True)
    right.image(treated, channels="BGR", caption="CLAHE applied" if applied else "Original retained · no CLAHE",
                use_container_width=True)
    st.markdown('<div class="cs-meta">'
                f'<span>Brightness · {metrics["brightness"]} / 255</span>'
                f'<span>Contrast span · {metrics["contrast"]} / 255</span>'
                f'<span>Enhancement suggested · {"Yes" if suggested else "No"}</span></div>', unsafe_allow_html=True)
    st.caption(f'{result["enhanced_count"]} of {result["frame_count"]} analyzed frame(s) were enhanced. '
               f'{result["suggested_count"]} met the provisional automatic rule.')
    notice("Brightness and contrast are provisional image heuristics. They do not measure turbidity or recover fish that are hidden from view.")


def render_observation_table(rows):
    headings = ["Frame", "Time (s)", "Predicted species", "Confidence", "Review flag", "CLAHE applied"]
    content = '<div class="cs-table-wrap"><table class="cs-table"><thead><tr>'
    content += ''.join(f'<th>{escape(heading)}</th>' for heading in headings) + '</tr></thead><tbody>'
    for row in rows:
        content += '<tr>'
        for heading in headings:
            value = f'{row[heading]:.1%}' if heading == "Confidence" else str(row[heading])
            cell = escape(value)
            if heading == "Review flag" and value == "Low confidence":
                cell = f'<span class="cs-flag">{cell}</span>'
            content += f'<td>{cell}</td>'
        content += '</tr>'
    if not rows:
        content += '<tr><td colspan="6">No observations match this view.</td></tr>'
    st.markdown(content + '</tbody></table></div>', unsafe_allow_html=True)


def render_sampled_counts(result, review_state=None, source_video=None):
    mc = result["sample_counts"]
    samples = mc["samples"]
    state = review_state or fresh_review_state(result)
    stats = review_statistics(result, state)
    is_protocol = result.get("sampling_plan") == "protocol"
    section("MeanCount", "AI estimate first · correction only if needed")
    c1, c2, c3 = st.columns([1, 1, 1])
    c1.metric("Preliminary AI mean", f'{mc["mean_fish"]:.2f}', help="Mean fish detections across all sampled frames, including zeros.")
    c2.metric("Mean with corrections", f'{stats["adjusted_mean"]:.2f}',
              help="Unedited samples still use AI counts. This is NOT a verified MeanCount.")
    c3.metric("Counts corrected", f'{stats["edited_count"]} of {stats["sample_count"]}')
    if not stats["edited_count"]:
        st.caption("No corrections saved. The adjusted mean currently matches the AI estimate.")
    notice("Both means are preliminary until Coastal Action completes its required observation and quality checks. "
           "A corrected average is not automatically a verified result.", gold=True)
    headings = ["Sample", "Time (s)", "AI detected fish", "Current fish", "Status"]
    html = '<div class="cs-table-wrap"><table class="cs-table"><thead><tr>'
    html += ''.join(f'<th>{escape(h)}</th>' for h in headings) + '</tr></thead><tbody>'
    for sample in samples:
        edited = state.get("samples", {}).get(str(sample["Sample"]))
        current = edited["fish"] if edited else sample["Fish detections"]
        status = "Corrected" if edited and int(edited["fish"]) != int(sample["Fish detections"]) else ("Note added" if edited else "AI estimate")
        html += ('<tr>' + f'<td>{sample["Sample"]}</td>'
                 f'<td>{sample["Time (s)"]:.1f}</td><td>{sample["Fish detections"]}</td>'
                 f'<td>{current}</td><td>{status}</td></tr>')
    st.markdown(html + '</tbody></table></div>', unsafe_allow_html=True)
    st.caption("Fish totals exclude the model's Whelk category. Every sampled frame is included, even if no fish were detected.")
    if result.get("sampling_plan") != "protocol":
        st.info("Quick-demo mean only, not the 15-minute / 15-sample draft protocol result.")
    if source_video is not None:
        render_researcher_review(result, source_video, state)


def render_results(result, source_video=None, review_state=None):
    rows = result["rows"]
    counts = Counter(row["Predicted species"] for row in rows)
    flagged = sum(row["Review flag"] == "Low confidence" for row in rows)
    workbook = excel_bytes(rows, result, review_state=review_state) if result["is_video"] else result["excel"]
    section("Survey overview", f'Completed in {result["elapsed"]:.1f}s · {result["frame_count"]} frame(s) analyzed')
    cards = [
        ("Frame detections", len(rows), "May include the same fish again", False),
        ("Predicted species", len(counts), "Preliminary model categories", False),
        ("Frames analyzed", result["frame_count"], f'{result["enhanced_count"]} enhanced with CLAHE', False),
        ("Review flags", flagged, "Model score below 25%", True),
    ]
    for col, (label, value, detail, gold) in zip(st.columns(4), cards):
        col.markdown(f'<div class="cs-kpi{" gold" if gold else ""}"><div class="cs-kpi-label">{label}</div>'
                     f'<div class="cs-kpi-value">{value}</div><div class="cs-kpi-detail">{detail}</div></div>',
                     unsafe_allow_html=True)
    if result["is_video"]:
        preview_tab, count_tab, comparison_tab, observations_tab = st.tabs([
            "Detection preview", "MeanCount & corrections", "Enhancement comparison", "Observations & export",
        ])
    else:
        preview_tab, comparison_tab, observations_tab = st.tabs([
            "Detection preview", "Enhancement comparison", "Observations & export",
        ])
    with preview_tab:
        media, summary = st.columns([1.8, 1], gap="large")
        with media:
            if result["is_video"]:
                show_gif(result["gif"],
                          f'Sampled still-frame slideshow · {result["frame_count"]} frames · '
                          + ("60-second spacing (NOT continuous motion)" if result.get("sampling_plan") == "protocol"
                             else "2 sampled frames per second"))
                st.download_button("Download sampled-frame preview", result["gif"],
                    file_name="coastalsense_detection_preview.gif", mime="image/gif", key="cs_rbc_gif_download")
            else:
                st.image(result["image"], caption="YOLOv7 · annotated species predictions", use_container_width=True)
                st.download_button("Download annotated photograph", result["image"],
                    file_name="coastalsense_detection.jpg", mime="image/jpeg", key="cs_rbc_image_download")
        with summary:
            render_species_summary(rows)
            st.markdown(" ")
            st.download_button("Download survey workbook", workbook,
                file_name="coastalsense_observations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="cs_rbc_excel_quick")
            st.caption("Simple workbook: overview and sampled counts. AI vs edited values stay separate.")
        notice("These are frame detections, not unique fish or fish-passage counts. Predictions need human review.", gold=True)
    if result["is_video"]:
        with count_tab:
            render_sampled_counts(result, review_state=review_state, source_video=source_video)
    with comparison_tab:
        render_comparison(result)
    with observations_tab:
        section("Detection observations", "Filter this view · export always includes all observations")
        filter_species, filter_review = st.columns(2)
        species_filter = filter_species.selectbox("Predicted species", ["All species"] + sorted(counts),
                                                 key="cs_rbc_species_filter")
        only_flagged = filter_review.checkbox("Show low-confidence flags only", key="cs_rbc_review_filter")
        filtered = [row for row in rows if (species_filter == "All species" or row["Predicted species"] == species_filter)
                    and (not only_flagged or row["Review flag"] == "Low confidence")]
        render_observation_table(filtered)
        st.caption(f"Showing {len(filtered)} of {len(rows)} detections. Excel summarizes counts rather than listing bounding-box coordinates.")
        st.download_button("Download complete Excel workbook", workbook,
            file_name="coastalsense_observations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="cs_rbc_excel_full")
    with st.expander("Technical details"):
        st.caption(f'Processed {result["created"]} · CPU · enhancement {result["mode"]} · '
                   f'threshold {result["confidence"]:.0%}. Model scores are not validated probabilities of correctness.')
        st.text(result["log"][-4500:])


def render_model_guide():
    with st.expander("About the model, tracking, and interpretation"):
        st.markdown("**What the prototype does**")
        st.write("Applies optional contrast enhancement, runs the N-MARINE YOLOv7 detector locally, "
                 "and exports preliminary observations. For a quick demo it samples 2 frames per second "
                  "from up to 6 seconds. The separate draft-protocol mode samples 15 still frames, "
                  "one every 60 seconds over a 15-minute analysis window.")
        st.markdown("**Supported model labels**")
        st.write(" · ".join(CLASS_NAMES))
        st.markdown("**Why this is not tracking yet**")
        st.write("The detector labels each frame independently. Tracking would need to associate a fish "
                 "across consecutive frames and keep its ID through movement and temporary occlusion. "
                 "That needs a separate tracking stage and validation on the survey footage. A sampled "
                 "GIF alone cannot establish reliable identities or unique fish counts.")
        st.markdown("**Review before drawing conclusions**")
        st.write("Species classifications may be wrong on unfamiliar footage. Low-visibility thresholds "
                 "are provisional. This prototype does not measure turbidity, verify unusual species, "
                 "estimate fish populations, or perform acoustic imaging. The sample mean uses model "
                  "detections in still frames; researchers may correct wrong samples without reviewing every one "
                  "in the interface. Independent quality control "
                  "and validation are still required.")


def main():
    st.set_page_config(page_title="CoastalSense | Marine Monitoring", page_icon="🐟",
                       layout="wide", initial_sidebar_state="expanded")
    st.markdown(STYLES, unsafe_allow_html=True)
    model_ready = MODEL.is_file() and (PROJECT / "detect.py").is_file()
    mode, confidence, contrast_cutoff = render_sidebar(model_ready)
    render_hero(model_ready)
    section("Survey footage", "JPG, PNG or MP4 · processed locally")
    footage, setup = st.columns([1.6, 1], gap="large")
    uploaded_bytes = None
    original = None
    is_video = False
    start_second = 0
    segment_seconds = 3
    sampling_plan = "quick"
    upload_valid = True
    with footage:
        file = st.file_uploader("Upload an underwater photograph or video", type=["jpg", "jpeg", "png", "mp4"],
            key="cs_rbc_upload", help="For video, analyze a short segment. Your media is processed by the model on this machine.")
        if file is not None:
            uploaded_bytes = file.getvalue()
            is_video = file.name.lower().endswith(".mp4")
            if not uploaded_bytes:
                st.error("This file is empty. Choose a photograph or MP4 with content.")
                upload_valid = False
            elif is_video:
                st.video(uploaded_bytes)
                st.caption("Original footage · select the segment to analyze on the right.")
            else:
                # Reuse the decoded upload across UI interactions in this session.
                digest = hashlib.sha256(uploaded_bytes).hexdigest()
                cached = st.session_state.get("cs_rbc_source_preview")
                if cached is None or cached[0] != digest:
                    decoded = cv2.imdecode(np.frombuffer(uploaded_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
                    st.session_state["cs_rbc_source_preview"] = (digest, decoded)
                original = st.session_state["cs_rbc_source_preview"][1]
                if original is None:
                    st.error("Could not read this image. Choose another JPG or PNG.")
                    upload_valid = False
                else:
                    st.image(original, channels="BGR", caption="Original survey photograph", use_container_width=True)
        else:
            st.markdown('<div class="cs-empty"><svg width="46" height="46" viewBox="0 0 46 46" fill="none" aria-hidden="true">'
                        '<rect x="5" y="8" width="36" height="30" rx="7" stroke="#9CB8D6" stroke-width="1.5"/>'
                        '<path d="m10 31 9-10 7 7 5-5 7 8" stroke="#0051A5" stroke-width="1.8"/>'
                        '<circle cx="31" cy="17" r="3" fill="#FEDF01"/></svg>'
                        '<h3>Start with a view underwater.</h3><p>Your source preview will appear here after upload.</p></div>',
                        unsafe_allow_html=True)
    with setup:
        if file is None:
            render_steps()
        else:
            st.markdown('<div class="cs-card"><div class="cs-eyebrow">Analysis setup</div>'
                        f'<h3>{"Video survey" if is_video else "Photograph survey"}</h3>'
                        f'<p>{escape(file.name)}</p><div class="cs-meta">'
                        f'<span>{len(uploaded_bytes) / (1024 * 1024):.2f} MB</span>'
                        f'<span>CLAHE · {mode}</span><span>Threshold · {confidence:.0%}</span></div></div>',
                        unsafe_allow_html=True)
            if is_video:
                plan_label = st.radio("Video analysis", ["Quick demo", "Draft protocol pilot"],
                    key="cs_mc_video_plan", horizontal=True)
                sampling_plan = "protocol" if plan_label == "Draft protocol pilot" else "quick"
                if sampling_plan == "protocol":
                    start_second = st.number_input("Analysis window start (seconds)", 0, 3600, 60, 1,
                        key="cs_mc_protocol_start")
                    segment_seconds = PROTOCOL_WINDOW_SECONDS
                    st.caption("15-minute window · 15 still frames, 60 seconds apart. "
                               "Start at 60s if the recording begins with 1 minute of settling. "
                               "This is an experimental match to the DRAFT sampling schedule.")
                    st.warning("Requires a video of at least 15 minutes after the chosen start. "
                               "Each timestamp uses one still image; manual context review is still necessary.")
                else:
                    start_second = st.number_input("Start at second", 0, 3600, 0, 1, key="cs_rbc_video_start")
                    segment_seconds = st.slider("Seconds to analyze", 1, 6, 3, key="cs_rbc_video_seconds")
                    st.caption(f"Up to {segment_seconds * SAMPLE_RATE} sampled still frames at "
                               f"{SAMPLE_RATE} frames/second, with a preview mean. "
                               "Not an official survey MeanCount.")
            elif original is not None:
                metrics = get_visibility_metrics(original)
                st.markdown('<div class="cs-meta">'
                    f'<span>Brightness · {metrics["brightness"]} / 255</span>'
                    f'<span>Contrast · {metrics["contrast"]} / 255</span></div>', unsafe_allow_html=True)
            st.markdown(" ")
        run_clicked = st.button("Analyze survey footage", type="primary", use_container_width=True,
            disabled=file is None or not model_ready or not upload_valid, key="cs_rbc_analyze")
        if not model_ready:
            st.caption("To enable analysis, place this app beside detect.py and best.pt in the inner yolov7-main folder.")
        elif file is None:
            st.caption("Upload your footage to enable analysis.")
        else:
            st.caption("Runs locally on your laptop’s CPU. Allow a few minutes for detection.")

    run_key = None
    if file is not None and uploaded_bytes and upload_valid:
        digest = hashlib.sha256(uploaded_bytes).hexdigest()
        run_key = (digest, file.name, mode, confidence, contrast_cutoff, is_video,
                   int(start_second) if is_video else 0, int(segment_seconds) if is_video else 0,
                   sampling_plan if is_video else "photo")
    if run_clicked and run_key is not None:
        progress_bar = st.progress(0, text="Starting your survey analysis…")
        try:
            result = analyze_upload(
                uploaded_bytes, file.name, is_video, mode, confidence, contrast_cutoff,
                start_second, segment_seconds, sampling_plan=sampling_plan,
                progress=lambda value, message: progress_bar.progress(value, text=message),
            )
            result["key"] = run_key
            # Commit results only after processing and Excel generation succeed.
            st.session_state[STATE_KEY] = result
            st.session_state[REVIEW_KEY] = fresh_review_state(result)
            st.session_state.pop("cs_rbc_species_filter", None)
            st.session_state.pop("cs_rbc_review_filter", None)
        except subprocess.TimeoutExpired:
            st.error("Detection exceeded the four-minute limit. Try the short demo mode, "
                     "or run the protocol mode on a faster local computer.")
        except DetectionError as exc:
            st.error(str(exc))
            if exc.log:
                with st.expander("Detection error details", expanded=True):
                    st.code(exc.log[-4500:])
        except Exception as exc:
            st.error(f"Processing could not finish: {exc}")
        finally:
            progress_bar.empty()

    result = st.session_state.get(STATE_KEY)
    if result is not None and run_key is not None and result.get("key") == run_key:
        review_state = st.session_state.get(REVIEW_KEY)
        if result.get("is_video") and (review_state is None or review_state.get("run_id") != result.get("run_id")):
            review_state = fresh_review_state(result)
            st.session_state[REVIEW_KEY] = review_state
        render_results(result, source_video=uploaded_bytes if is_video else None,
                       review_state=review_state)
    elif file is not None and upload_valid:
        section("Survey overview", "Results appear after analysis")
        notice("Your footage or settings have changed. Analyze to produce results for this selection."
               if result else "Ready when you are. Run analysis to see species predictions, enhancement comparisons, and downloadable observations.")
    render_model_guide()
    st.markdown('<div class="cs-footer"><span><strong>CoastalSense</strong> &nbsp; Marine monitoring prototype</span>'
                '<span>N-MARINE YOLOv7 · CLAHE enhancement · Human review required</span></div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
