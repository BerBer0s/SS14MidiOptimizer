#!/usr/bin/env python3
"""
MIDI Оптимизатор для Space Station 14 — GUI версия (PySide6)
Требования: pip install PySide6 mido pillow
"""

import sys
import io
import os
import struct
import base64
import threading
import queue
import traceback
from pathlib import Path
from collections import defaultdict

# DPI awareness на Windows
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QSlider, QCheckBox, QTextEdit,
    QProgressBar, QScrollArea, QFrame, QFileDialog, QMessageBox,
    QSizePolicy, QToolTip, QStyleFactory, QDialog, QSplitter,
)
from PySide6.QtCore import (
    Qt, QSettings, Signal, QObject,
)
from PySide6.QtGui import (
    QIcon, QPixmap, QFont, QColor, QPalette, QDragEnterEvent, QDropEvent,
)


# ─── Константы по умолчанию ──────────────────────────────────────────────────
DEFAULT_DRUM_CAP        = 85
DEFAULT_MELODIC_CAP     = 127

GM_NAMES = [
    "Acoustic Grand Piano","Bright Acoustic Piano","Electric Grand Piano","Honky-tonk Piano",
    "Electric Piano 1","Electric Piano 2","Harpsichord","Clavinet",
    "Celesta","Glockenspiel","Music Box","Vibraphone","Marimba","Xylophone","Tubular Bells","Dulcimer",
    "Drawbar Organ","Percussive Organ","Rock Organ","Church Organ","Reed Organ","Accordion","Harmonica","Tango Accordion",
    "Nylon Guitar","Steel Guitar","Jazz Guitar","Clean Guitar","Muted Guitar","Overdriven Guitar","Distortion Guitar","Guitar Harmonics",
    "Acoustic Bass","Electric Bass (finger)","Electric Bass (pick)","Fretless Bass","Slap Bass 1","Slap Bass 2","Synth Bass 1","Synth Bass 2",
    "Violin","Viola","Cello","Contrabass","Tremolo Strings","Pizzicato Strings","Orchestral Harp","Timpani",
    "String Ensemble 1","String Ensemble 2","Synth Strings 1","Synth Strings 2","Choir Aahs","Voice Oohs","Synth Choir","Orchestra Hit",
    "Trumpet","Trombone","Tuba","Muted Trumpet","French Horn","Brass Section","Synth Brass 1","Synth Brass 2",
    "Soprano Sax","Alto Sax","Tenor Sax","Baritone Sax","Oboe","English Horn","Bassoon","Clarinet",
    "Piccolo","Flute","Recorder","Pan Flute","Blown Bottle","Shakuhachi","Whistle","Ocarina",
    "Square Lead","Sawtooth Lead","Calliope Lead","Chiff Lead","Charang Lead","Voice Lead","Fifths Lead","Bass+Lead",
    "New Age Pad","Warm Pad","Polysynth Pad","Choir Pad","Bowed Pad","Metallic Pad","Halo Pad","Sweep Pad",
    "Rain","Soundtrack","Crystal","Atmosphere","Brightness","Goblins","Echoes","Sci-fi",
    "Sitar","Banjo","Shamisen","Koto","Kalimba","Bag pipe","Fiddle","Shanai",
    "Tinkle Bell","Agogo","Steel Drums","Woodblock","Taiko Drum","Melodic Tom","Synth Drum","Reverse Cymbal",
    "Guitar Fret Noise","Breath Noise","Seashore","Bird Tweet","Telephone Ring","Helicopter","Applause","Gunshot",
]

def gm_name(p):
    return GM_NAMES[p] if 0 <= p < len(GM_NAMES) else f"Program {p}"


# ═══════════════════════════════════════════════════════════════════════════════
#  ЛОГИКА ОПТИМИЗАТОРА (без изменений)
# ═══════════════════════════════════════════════════════════════════════════════

class UserSkipped(Exception):
    pass

def msg_priority(msg):
    if msg.type == 'note_off': return 0
    if msg.type == 'note_on' and msg.velocity == 0: return 0
    if msg.type in ('program_change', 'control_change', 'pitchwheel'): return 1
    if msg.type == 'note_on': return 2
    return 1

def collect_by_channel(mid):
    channel_msgs     = defaultdict(list)
    channel_programs = {}
    channel_names    = defaultdict(set)
    for track in mid.tracks:
        abs_tick = 0
        for msg in track:
            abs_tick += msg.time
            if msg.is_meta: continue
            ch = getattr(msg, 'channel', None)
            if ch is None: continue
            channel_msgs[ch].append((abs_tick, msg))
            if msg.type == 'program_change' and ch not in channel_programs:
                channel_programs[ch] = msg.program
            if track.name:
                channel_names[ch].add(track.name)
    return channel_msgs, channel_programs, channel_names

def repair_notes(abs_msgs, ch):
    import mido
    sorted_msgs = sorted(abs_msgs, key=lambda x: (x[0], msg_priority(x[1])))
    open_notes  = {}
    repaired    = []
    fixed_dupes = 0
    fixed_hang  = 0
    for tick, msg in sorted_msgs:
        if msg.type == 'note_on' and msg.velocity > 0:
            if msg.note in open_notes:
                repaired.append((tick, mido.Message('note_off', channel=ch, note=msg.note, velocity=0, time=0)))
                fixed_dupes += 1
            open_notes[msg.note] = tick
            repaired.append((tick, msg))
        elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
            open_notes.pop(msg.note, None)
            repaired.append((tick, msg))
        else:
            repaired.append((tick, msg))
    if open_notes:
        last_tick = repaired[-1][0] if repaired else 0
        for pitch in open_notes:
            repaired.append((last_tick, mido.Message('note_off', channel=ch, note=pitch, velocity=0, time=0)))
            fixed_hang += 1
    return repaired, fixed_hang, fixed_dupes


def merge_and_relativize(abs_msgs):
    sorted_msgs = sorted(abs_msgs, key=lambda x: (x[0], msg_priority(x[1])))
    result = []
    prev_tick = 0
    for abs_tick, msg in sorted_msgs:
        result.append(msg.copy(time=max(0, abs_tick - prev_tick)))
        prev_tick = abs_tick
    return result

def find_first_note_tick(channel_msgs):
    first = None
    for msgs in channel_msgs.values():
        for tick, msg in msgs:
            if msg.type == 'note_on' and msg.velocity > 0:
                if first is None or tick < first:
                    first = tick
    return first or 0

def trim_silence(channel_msgs, tempo_msgs_abs):
    offset = find_first_note_tick(channel_msgs)
    if offset == 0:
        return channel_msgs, tempo_msgs_abs, offset
    trimmed = {ch: [(max(0, t - offset), m) for t, m in msgs] for ch, msgs in channel_msgs.items()}
    tempo_trimmed = [(max(0, t - offset), m) for t, m in tempo_msgs_abs]
    return trimmed, tempo_trimmed, offset

def normalize_velocity(messages, cap):
    max_vel = max((m.velocity for m in messages if m.type == 'note_on' and m.velocity > 0), default=0)
    if max_vel == 0 or max_vel >= cap:
        return messages, max_vel, max_vel
    scale = cap / max_vel
    result = []
    for msg in messages:
        if msg.type == 'note_on' and msg.velocity > 0:
            result.append(msg.copy(velocity=min(cap, int(msg.velocity * scale))))
        else:
            result.append(msg)
    return result, max_vel, cap

def collect_tempo_abs(mid):
    import mido
    result = []
    seen   = set()
    tracks_to_scan = [mid.tracks[0]] if mid.type == 1 and mid.tracks else mid.tracks
    for track in tracks_to_scan:
        abs_tick = 0
        for msg in track:
            abs_tick += msg.time
            if msg.is_meta and msg.type in ('set_tempo', 'time_signature', 'key_signature'):
                key = (msg.type, abs_tick)
                if key not in seen:
                    result.append((abs_tick, msg))
                    seen.add(key)
    result.sort(key=lambda x: x[0])
    if not any(m.type == 'set_tempo' for _, m in result):
        result.insert(0, (0, mido.MetaMessage('set_tempo', tempo=500000, time=0)))
    return result

def abs_to_delta(abs_msgs):
    result = []
    prev = 0
    for tick, msg in sorted(abs_msgs, key=lambda x: x[0]):
        result.append(msg.copy(time=max(0, tick - prev)))
        prev = tick
    return result

def _load_midi_safe(path, yes_all=False, log_fn=print, ask_fn=None):
    import mido
    try:
        return mido.MidiFile(path, clip=True), False
    except Exception:
        pass

    with open(path, 'rb') as f:
        raw = f.read()

    if raw[:4] != b'MThd' or len(raw) < 14:
        raise OSError("Файл не является корректным MIDI")

    fmt = struct.unpack('>H', raw[10:12])[0]
    tpb = struct.unpack('>H', raw[12:14])[0]
    fmt_broken = fmt not in (0, 1, 2)
    if fmt_broken:
        fmt = 1

    problems = []
    if fmt_broken:
        problems.append(f"битый заголовок формата (значение {fmt})")
    bad_bytes = set()
    for b in raw:
        if b in (0xf4, 0xf5, 0xf9, 0xfd, 0xfe):
            bad_bytes.add(hex(b))
    if bad_bytes:
        problems.append(f"запрещённые байты: {', '.join(sorted(bad_bytes))}")

    log_fn("  ╔══════════════════════════════════════════════╗")
    log_fn("  ║         ⚠  ФАЙЛ ПОВРЕЖДЁН  ⚠                ║")
    log_fn("  ╠══════════════════════════════════════════════╣")
    for p in problems:
        line = f"  ║  • {p}"
        log_fn(line[:52].ljust(52) + "║")
    log_fn("  ╠══════════════════════════════════════════════╣")
    log_fn("  ║  Результат может быть искажён.               ║")
    log_fn("  ╚══════════════════════════════════════════════╝")

    ya_requested = False
    if yes_all:
        log_fn("  [ya] Авто-подтверждение — обрабатываю.")
    elif ask_fn:
        answer = ask_fn("\n".join(problems))
        if answer == 'ya':
            ya_requested = True
            log_fn("  [ya] Авто-подтверждение включено.")
        elif answer != 'y':
            raise UserSkipped()
    else:
        raise UserSkipped()

    log_fn(f"  Читаю треки (fmt={fmt}, tpb={tpb})...")
    good = mido.MidiFile(type=fmt, ticks_per_beat=tpb)
    pos = 14
    track_num = 0
    while pos < len(raw) - 8:
        if raw[pos:pos+4] == b'MTrk':
            length = struct.unpack('>I', raw[pos+4:pos+8])[0]
            if pos + 8 + length > len(raw):
                break
            track_bytes = raw[pos:pos+8+length]
            header = b'MThd\x00\x00\x00\x06' + struct.pack('>HHH', 0, 1, tpb)
            try:
                tmp = mido.MidiFile(file=io.BytesIO(header + track_bytes), clip=True)
                if tmp.tracks:
                    good.tracks.append(tmp.tracks[0])
                    log_fn(f"  Трек {track_num}: ОК ({len(tmp.tracks[0])} событий)")
            except Exception as e:
                log_fn(f"  Трек {track_num}: пропущен — {e}")
            pos += 8 + length
            track_num += 1
        else:
            pos += 1

    if not good.tracks:
        raise OSError("Не удалось прочитать ни один трек")

    return good, ya_requested


def optimize(input_path: Path, output_path: Path,
             yes_all=False,
             drum_cap=DEFAULT_DRUM_CAP,
             melodic_cap=DEFAULT_MELODIC_CAP,
             log_fn=print,
             ask_fn=None):
    import mido

    log_fn(f"\n▶ {input_path.name}")

    ya_requested = False
    try:
        mid, ya_req = _load_midi_safe(input_path, yes_all=yes_all, log_fn=log_fn, ask_fn=ask_fn)
        ya_requested = ya_req
    except UserSkipped:
        log_fn("  ↩ Пропущен.")
        return False, False

    log_fn(f"  Дорожек: {len(mid.tracks)}  TPB: {mid.ticks_per_beat}")

    channel_msgs, channel_programs, channel_names = collect_by_channel(mid)
    tempo_abs = collect_tempo_abs(mid)

    channel_msgs, tempo_abs, silence_ticks = trim_silence(channel_msgs, tempo_abs)
    if silence_ticks > 0:
        log_fn(f"  Тишина обрезана: {silence_ticks} тиков ({silence_ticks/mid.ticks_per_beat:.2f} долей)")

    log_fn(f"  Каналов: {len(channel_msgs)}")
    for ch in sorted(channel_msgs.keys()):
        prog     = channel_programs.get(ch)
        n        = sum(1 for _, m in channel_msgs[ch] if m.type == 'note_on' and m.velocity > 0)
        drum_tag = " [УД]" if ch == 9 else ""
        log_fn(f"    Кан.{ch+1:2d}{drum_tag}: {n} нот | {gm_name(prog) if prog is not None else '—'}")

    out = mido.MidiFile(type=1, ticks_per_beat=mid.ticks_per_beat)
    meta_track = mido.MidiTrack()
    meta_track.name = "Tempo"
    for msg in abs_to_delta(tempo_abs):
        meta_track.append(msg)
    out.tracks.append(meta_track)

    total_hang = total_dupes = 0
    for ch in sorted(channel_msgs.keys()):
        abs_msgs = channel_msgs[ch]
        if not abs_msgs:
            continue

        is_drum = (ch == 9)
        cap     = drum_cap if is_drum else melodic_cap

        abs_msgs, hang, dupes = repair_notes(abs_msgs, ch)
        total_hang  += hang
        total_dupes += dupes

        messages = merge_and_relativize(abs_msgs)
        messages, vel_before, vel_after = normalize_velocity(messages, cap)

        n        = sum(1 for m in messages if m.type == 'note_on' and m.velocity > 0)
        prog     = channel_programs.get(ch)
        drum_tag = "[УД] " if is_drum else ""
        extra    = []
        if hang or dupes:
            extra.append(f"🔧зав:{hang} дуб:{dupes}")
        extra_str = "  " + " ".join(extra) if extra else ""
        log_fn(f"    Кан.{ch+1:2d} {drum_tag}vel {vel_before}→{vel_after}, {n} нот{extra_str}")

        names = ', '.join(sorted(channel_names[ch]))
        track = mido.MidiTrack()
        track.name = names[:28] if names else f"Channel {ch+1}"
        out.tracks.append(track)

        track.append(mido.Message('control_change', channel=ch, control=7,  value=127, time=0))
        track.append(mido.Message('control_change', channel=ch, control=11, value=127, time=0))

        if prog is not None and not is_drum:
            track.append(mido.Message('program_change', channel=ch, program=prog, time=0))

        for msg in messages:
            if msg.type == 'program_change':
                continue
            track.append(msg)

    if total_hang or total_dupes:
        log_fn(f"  🔧 Починено: {total_hang} зависших, {total_dupes} дублей")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(output_path)

    total = sum(
        sum(1 for m in t if not m.is_meta and m.type == 'note_on' and m.velocity > 0)
        for t in out.tracks
    )
    log_fn(f"  ✓ Готово → {output_path.name}  [{len(out.tracks)-1} дор. / {total} нот]")
    return True, ya_requested


# ═══════════════════════════════════════════════════════════════════════════════
#  ИКОНКА (embedded base64 — fallback если .ico не найден)
# ═══════════════════════════════════════════════════════════════════════════════

ICON_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAdElEQVR4nKWTSw7AIAgFGdP7X5lu"
    "SoOiqIElvB9ERYrFaqCqGsAQ8KExI2ZC7ZQMiPE87ldakSepxaAALUwvi8x9lcCnKCd4Th3NFRAf"
    "uEuQkb2Ir/IK7XPeW08KoLuB7Tj2UpEBvH3G3l0kHvFoFY8r/8YXiE81JOys32YAAAAASUVORK5C"
    "YII="
)

def get_app_icon():
    """Получить QIcon — сначала ищем .ico рядом с exe, затем base64-fallback."""
    candidates = [
        Path(os.path.abspath(sys.argv[0])).parent / "icon.ico",
        Path(sys.executable).parent / "icon.ico",
    ]
    if hasattr(sys, '_MEIPASS'):
        candidates.insert(0, Path(sys._MEIPASS) / "icon.ico")
    try:
        candidates.append(Path(__file__).parent / "icon.ico")
    except NameError:
        pass

    for ico_path in candidates:
        if ico_path.exists():
            icon = QIcon(str(ico_path))
            if not icon.isNull():
                return icon

    # base64 fallback
    png_data = base64.b64decode(ICON_B64)
    pixmap = QPixmap()
    pixmap.loadFromData(png_data)
    return QIcon(pixmap)


# ═══════════════════════════════════════════════════════════════════════════════
#  СТИЛИ
# ═══════════════════════════════════════════════════════════════════════════════

STYLESHEET = """
/* ── Глобальные ─────────────────────────────────────────────── */
QMainWindow, QWidget#central {
    background-color: #141414;
}

QToolTip {
    background-color: #1e1e1e;
    color: #e0e0e0;
    border: 1px solid #3a3a3a;
    padding: 6px 10px;
    font-size: 11px;
    border-radius: 4px;
}

/* ── Панели ─────────────────────────────────────────────────── */
QFrame#header {
    background-color: #0e0e0e;
    border-bottom: 1px solid #2a2a2a;
}

QFrame#panel {
    background-color: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 6px;
}

QFrame#section {
    background-color: #1e1e1e;
    border: 1px solid #2a2a2a;
    border-radius: 6px;
}

QFrame#footer {
    background-color: #0e0e0e;
    border-top: 1px solid #2a2a2a;
}

/* ── Текст ──────────────────────────────────────────────────── */
QLabel {
    color: #e0e0e0;
    font-size: 11px;
}

QLabel#dim {
    color: #777777;
}

QLabel#title {
    color: #f0f0f0;
    font-weight: bold;
    font-size: 13px;
}

QLabel#header_title {
    color: #f0f0f0;
    font-weight: bold;
    font-size: 15px;
}

QLabel#header_author {
    color: #555555;
    font-size: 11px;
}

QLabel#slider_value {
    color: #f0f0f0;
    font-weight: bold;
    font-size: 11px;
    min-width: 42px;
}

QLabel#file_done {
    color: #555555;
    font-size: 11px;
}

QLabel#tip_icon {
    color: #555555;
    font-size: 14px;
}

QLabel#drop_hint {
    color: #3a3a3a;
    font-size: 12px;
}

/* ── Кнопки ─────────────────────────────────────────────────── */
QPushButton#primary {
    background-color: #e8e8e8;
    color: #111111;
    border: none;
    border-radius: 5px;
    font-weight: bold;
    font-size: 13px;
    padding: 8px 20px;
}
QPushButton#primary:hover {
    background-color: #ffffff;
}
QPushButton#primary:pressed {
    background-color: #cccccc;
}
QPushButton#primary:disabled {
    background-color: #333333;
    color: #666666;
}

QPushButton#secondary {
    background-color: #2a2a2a;
    color: #bbbbbb;
    border: 1px solid #3a3a3a;
    border-radius: 5px;
    font-size: 11px;
    padding: 5px 14px;
}
QPushButton#secondary:hover {
    background-color: #363636;
    color: #e0e0e0;
}
QPushButton#secondary:pressed {
    background-color: #222222;
}
QPushButton#secondary:disabled {
    background-color: #1e1e1e;
    color: #444444;
    border-color: #2a2a2a;
}

QPushButton#small {
    background-color: #222222;
    color: #888888;
    border: 1px solid #2e2e2e;
    border-radius: 4px;
    font-size: 10px;
    padding: 4px 10px;
}
QPushButton#small:hover {
    background-color: #2e2e2e;
    color: #cccccc;
}

/* ── Слайдеры ───────────────────────────────────────────────── */
QSlider::groove:horizontal {
    background: #2a2a2a;
    height: 4px;
    border-radius: 2px;
}
QSlider::handle:horizontal {
    background: #d0d0d0;
    border: none;
    width: 14px;
    height: 14px;
    margin: -5px 0;
    border-radius: 7px;
}
QSlider::handle:horizontal:hover {
    background: #ffffff;
}
QSlider::sub-page:horizontal {
    background: #777777;
    border-radius: 2px;
}

/* ── Чекбоксы ───────────────────────────────────────────────── */
QCheckBox {
    color: #cccccc;
    spacing: 6px;
    font-size: 11px;
}
QCheckBox::indicator {
    width: 15px;
    height: 15px;
    border: 1px solid #555555;
    border-radius: 3px;
    background: #1a1a1a;
}
QCheckBox::indicator:checked {
    background: #d0d0d0;
    border-color: #d0d0d0;
}
QCheckBox::indicator:hover {
    border-color: #888888;
}

QCheckBox#toggle {
    font-size: 11px;
}

/* ── Лог ────────────────────────────────────────────────────── */
QTextEdit#log {
    background-color: #141414;
    color: #b0b0b0;
    border: 1px solid #2a2a2a;
    border-radius: 4px;
    font-family: 'Cascadia Code', 'Consolas', 'Courier New', monospace;
    font-size: 10px;
    padding: 6px;
    selection-background-color: #3a3a3a;
}

/* ── Прогресс-бар ───────────────────────────────────────────── */
QProgressBar {
    background-color: #1e1e1e;
    border: 1px solid #2a2a2a;
    border-radius: 4px;
    height: 10px;
    text-align: center;
    font-size: 0px;
}
QProgressBar::chunk {
    background-color: #aaaaaa;
    border-radius: 3px;
}

/* ── Скролл ─────────────────────────────────────────────────── */
QScrollArea {
    border: none;
    background-color: transparent;
}
QScrollBar:vertical {
    background: #141414;
    width: 8px;
    margin: 0;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #2e2e2e;
    min-height: 30px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #444444;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: none;
}

/* ── Сплиттер ───────────────────────────────────────────────── */
QSplitter::handle {
    background: transparent;
    width: 6px;
}
"""


# ═══════════════════════════════════════════════════════════════════════════════
#  СИГНАЛЫ ДЛЯ МЕЖПОТОЧНОГО ВЗАИМОДЕЙСТВИЯ
# ═══════════════════════════════════════════════════════════════════════════════

class WorkerSignals(QObject):
    log       = Signal(str)
    progress  = Signal(int, int)
    ask       = Signal(str)
    done      = Signal(int, int, int)


# ═══════════════════════════════════════════════════════════════════════════════
#  ОПИСАНИЯ ТУЛТИПОВ
# ═══════════════════════════════════════════════════════════════════════════════

TOOLTIPS = {
    "drum_cap":
        "Максимальная громкость ударных инструментов.\n"
        "85 — стандарт SS14, при значениях выше\n"
        "ударные будут слишком громкими.",
    "melodic_cap":
        "Максимальная громкость мелодических дорожек.\n"
        "127 — без ограничений. Снижай если\n"
        "общий звук слишком громкий.",
    "yes_all":
        "При обработке повреждённых файлов\n"
        "не спрашивать — обрабатывать автоматически.",
}


# ═══════════════════════════════════════════════════════════════════════════════
#  ГЛАВНОЕ ОКНО
# ═══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MIDI Оптимизатор для SS14")
        self.setMinimumSize(940, 640)
        self.resize(1100, 750)
        self.setWindowIcon(get_app_icon())

        # Состояние
        self._files: list[Path] = []
        self._input_dir:  Path | None = None
        self._output_dir: Path | None = None
        self._running    = False
        self._stop_flag  = False
        self._yes_all    = False
        self._ask_answer_queue: queue.Queue = queue.Queue()

        # Сигналы
        self._signals = WorkerSignals()
        self._signals.log.connect(self._on_log)
        self._signals.progress.connect(self._on_progress)
        self._signals.ask.connect(self._on_ask)
        self._signals.done.connect(self._on_done)

        # Настройки
        self._settings = QSettings("Berber0s", "MIDI_Optimizer_SS14")

        # Нативный DnD
        self.setAcceptDrops(True)

        self._build_ui()
        self._restore_settings()

    # ── Drag & Drop ──────────────────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        folders, files = [], []
        for url in urls:
            p = Path(url.toLocalFile())
            if p.is_dir():
                folders.append(p)
            elif p.is_file() and p.suffix.lower() in ('.mid', '.midi'):
                files.append(p)

        if folders:
            self._input_dir = folders[0]
            self._lbl_in.setText(self._short_path(str(self._input_dir)))
            self._lbl_in.setStyleSheet("color: #e0e0e0; font-size: 11px;")
            self._files.clear()
            self._refresh_files()
            self._save_settings()
        elif files:
            self._files = files
            self._input_dir = files[0].parent
            self._lbl_in.setText(self._short_path(str(self._input_dir)))
            self._lbl_in.setStyleSheet("color: #e0e0e0; font-size: 11px;")
            self._redraw_list()
            self._save_settings()

    # ── Построение UI ────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        main_layout.addWidget(self._make_header())

        body = QWidget()
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(8, 8, 8, 8)
        body_layout.setSpacing(0)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._make_left_panel())
        splitter.addWidget(self._make_right_panel())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        body_layout.addWidget(splitter)

        main_layout.addWidget(body, 1)
        main_layout.addWidget(self._make_footer())

    def _make_header(self):
        header = QFrame()
        header.setObjectName("header")
        header.setFixedHeight(48)

        layout = QHBoxLayout(header)
        layout.setContentsMargins(16, 0, 16, 0)

        title = QLabel("♩  MIDI Оптимизатор для Space Station 14")
        title.setObjectName("header_title")
        layout.addWidget(title)
        layout.addStretch()

        author = QLabel("by Berber0s")
        author.setObjectName("header_author")
        layout.addWidget(author)

        return header

    def _make_left_panel(self):
        panel = QFrame()
        panel.setObjectName("panel")
        pl = QVBoxLayout(panel)
        pl.setContentsMargins(8, 8, 8, 8)
        pl.setSpacing(6)

        # Папки
        folders = QFrame()
        folders.setObjectName("section")
        fl = QGridLayout(folders)
        fl.setContentsMargins(10, 10, 10, 10)
        fl.setSpacing(6)
        fl.setColumnStretch(1, 1)

        fl.addWidget(QLabel("Вход:"), 0, 0)
        self._lbl_in = QLabel("(не выбрана)")
        self._lbl_in.setStyleSheet("color: #777777; font-size: 11px;")
        fl.addWidget(self._lbl_in, 0, 1)
        b = QPushButton("Выбрать")
        b.setObjectName("secondary")
        b.clicked.connect(self._browse_input)
        fl.addWidget(b, 0, 2)

        fl.addWidget(QLabel("Вывод:"), 1, 0)
        self._lbl_out = QLabel("(рядом с исходником)")
        self._lbl_out.setStyleSheet("color: #777777; font-size: 11px;")
        fl.addWidget(self._lbl_out, 1, 1)
        b2 = QPushButton("Выбрать")
        b2.setObjectName("secondary")
        b2.clicked.connect(self._browse_output)
        fl.addWidget(b2, 1, 2)

        pl.addWidget(folders)

        # Список файлов
        self._file_scroll = QScrollArea()
        self._file_scroll.setWidgetResizable(True)
        self._file_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self._file_container = QWidget()
        self._file_container.setStyleSheet("background-color: #141414;")
        self._file_layout = QVBoxLayout(self._file_container)
        self._file_layout.setContentsMargins(4, 4, 4, 4)
        self._file_layout.setSpacing(1)
        self._file_layout.setAlignment(Qt.AlignTop)

        hint = QLabel("Перетащи папку или MIDI файлы сюда")
        hint.setObjectName("drop_hint")
        hint.setAlignment(Qt.AlignCenter)
        hint.setContentsMargins(0, 60, 0, 60)
        self._file_layout.addWidget(hint)

        self._file_scroll.setWidget(self._file_container)
        pl.addWidget(self._file_scroll, 1)

        self._file_vars: list[QCheckBox] = []

        # Кнопки
        br = QHBoxLayout()
        br.setSpacing(4)
        for txt, fn in [
            ("✔ Все",      lambda: self._select_all(True)),
            ("✘ Снять",    lambda: self._select_all(False)),
            ("↺ Обновить", self._refresh_files),
        ]:
            btn = QPushButton(txt)
            btn.setObjectName("small")
            btn.clicked.connect(fn)
            br.addWidget(btn)
        br.addStretch()
        self._lbl_count = QLabel("0 файлов")
        self._lbl_count.setStyleSheet("color: #555555; font-size: 11px;")
        br.addWidget(self._lbl_count)
        pl.addLayout(br)

        return panel

    def _make_right_panel(self):
        panel = QFrame()
        panel.setObjectName("panel")
        pl = QVBoxLayout(panel)
        pl.setContentsMargins(8, 8, 8, 8)
        pl.setSpacing(6)

        # Настройки
        sett = QFrame()
        sett.setObjectName("section")
        sl = QVBoxLayout(sett)
        sl.setContentsMargins(12, 10, 12, 12)
        sl.setSpacing(2)

        t = QLabel("Настройки")
        t.setObjectName("title")
        sl.addWidget(t)
        sl.addSpacing(6)

        self._drum_slider, self._drum_val = self._make_slider(
            sl, "Velocity ударных", 1, 127, DEFAULT_DRUM_CAP, "drum_cap")
        self._melodic_slider, self._melodic_val = self._make_slider(
            sl, "Velocity мелодии", 1, 127, DEFAULT_MELODIC_CAP, "melodic_cap")

        sl.addSpacing(6)
        ya = QHBoxLayout()
        self._yes_all_cb = QCheckBox("Авто-подтверждение битых")
        self._yes_all_cb.setObjectName("toggle")
        self._yes_all_cb.setToolTip(TOOLTIPS["yes_all"])
        ya.addWidget(self._yes_all_cb)
        ya.addStretch()
        sl.addLayout(ya)

        pl.addWidget(sett)

        # Лог
        lh = QHBoxLayout()
        lt = QLabel("Лог обработки")
        lt.setObjectName("title")
        lh.addWidget(lt)
        lh.addStretch()
        bc = QPushButton("Очистить")
        bc.setObjectName("small")
        bc.clicked.connect(self._clear_log)
        lh.addWidget(bc)
        pl.addLayout(lh)

        self._log = QTextEdit()
        self._log.setObjectName("log")
        self._log.setReadOnly(True)
        pl.addWidget(self._log, 1)

        return panel

    def _make_footer(self):
        footer = QFrame()
        footer.setObjectName("footer")
        footer.setFixedHeight(56)
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(8)

        self._btn_all = QPushButton("▶  Обработать ВСЕ")
        self._btn_all.setObjectName("primary")
        self._btn_all.setFixedHeight(36)
        self._btn_all.setMinimumWidth(170)
        self._btn_all.clicked.connect(self._run_all)
        layout.addWidget(self._btn_all)

        self._btn_sel = QPushButton("▶  Выбранные")
        self._btn_sel.setObjectName("secondary")
        self._btn_sel.setFixedHeight(36)
        self._btn_sel.setMinimumWidth(130)
        self._btn_sel.clicked.connect(self._run_selected)
        layout.addWidget(self._btn_sel)

        pw = QWidget()
        pvl = QVBoxLayout(pw)
        pvl.setContentsMargins(0, 0, 0, 0)
        pvl.setSpacing(2)
        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        pvl.addWidget(self._progress)
        self._lbl_status = QLabel("Готов")
        self._lbl_status.setStyleSheet("color: #555555; font-size: 10px;")
        pvl.addWidget(self._lbl_status)
        layout.addWidget(pw, 1)

        self._btn_stop = QPushButton("Стоп")
        self._btn_stop.setObjectName("secondary")
        self._btn_stop.setFixedHeight(36)
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        layout.addWidget(self._btn_stop)

        return footer

    def _make_slider(self, parent_layout, label, from_, to, default, tooltip_key, fmt=None):
        if fmt is None:
            fmt = lambda v: str(v)

        row = QHBoxLayout()
        row.setSpacing(6)

        tip = QLabel("ⓘ")
        tip.setObjectName("tip_icon")
        tip.setCursor(Qt.WhatsThisCursor)
        tip.setToolTip(TOOLTIPS.get(tooltip_key, ""))
        row.addWidget(tip)

        nm = QLabel(label)
        nm.setStyleSheet("font-size: 11px;")
        row.addWidget(nm)
        row.addStretch()

        vl = QLabel(fmt(default))
        vl.setObjectName("slider_value")
        vl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        row.addWidget(vl)

        parent_layout.addLayout(row)

        slider = QSlider(Qt.Horizontal)
        slider.setRange(from_, to)
        slider.setValue(default)
        slider.setSingleStep(1)
        slider.setPageStep(5)
        slider.valueChanged.connect(lambda v: vl.setText(fmt(v)))
        parent_layout.addWidget(slider)
        parent_layout.addSpacing(4)

        return slider, vl

    # ── Настройки ─────────────────────────────────────────────────────────────

    def _restore_settings(self):
        d = self._settings.value("last_input_dir", "")
        if d and Path(d).exists():
            self._input_dir = Path(d)
            self._lbl_in.setText(self._short_path(str(self._input_dir)))
            self._lbl_in.setStyleSheet("color: #e0e0e0; font-size: 11px;")
            self._refresh_files()

        d = self._settings.value("last_output_dir", "")
        if d and Path(d).exists():
            self._output_dir = Path(d)
            self._lbl_out.setText(self._short_path(str(self._output_dir)))
            self._lbl_out.setStyleSheet("color: #e0e0e0; font-size: 11px;")

        self._drum_slider.setValue(int(self._settings.value("drum_cap", DEFAULT_DRUM_CAP)))
        self._melodic_slider.setValue(int(self._settings.value("melodic_cap", DEFAULT_MELODIC_CAP)))
        self._yes_all_cb.setChecked(self._settings.value("yes_all", "false") == "true")

    def _save_settings(self):
        if self._input_dir:
            self._settings.setValue("last_input_dir", str(self._input_dir))
        if self._output_dir:
            self._settings.setValue("last_output_dir", str(self._output_dir))
        self._settings.setValue("drum_cap", self._drum_slider.value())
        self._settings.setValue("melodic_cap", self._melodic_slider.value())
        self._settings.setValue("yes_all", "true" if self._yes_all_cb.isChecked() else "false")

    def closeEvent(self, event):
        self._save_settings()
        super().closeEvent(event)

    # ── Утилиты ───────────────────────────────────────────────────────────────

    @staticmethod
    def _short_path(p: str, n: int = 50) -> str:
        return ("…" + p[-(n-1):]) if len(p) > n else p

    # ── Файлы ─────────────────────────────────────────────────────────────────

    def _browse_input(self):
        d = QFileDialog.getExistingDirectory(self, "Папка с MIDI файлами",
                                              str(self._input_dir or ""))
        if d:
            self._input_dir = Path(d)
            self._lbl_in.setText(self._short_path(str(self._input_dir)))
            self._lbl_in.setStyleSheet("color: #e0e0e0; font-size: 11px;")
            self._files.clear()
            self._refresh_files()
            self._save_settings()

    def _browse_output(self):
        d = QFileDialog.getExistingDirectory(self, "Папка для результатов",
                                              str(self._output_dir or ""))
        if d:
            self._output_dir = Path(d)
            self._lbl_out.setText(self._short_path(str(self._output_dir)))
            self._lbl_out.setStyleSheet("color: #e0e0e0; font-size: 11px;")
        else:
            self._output_dir = None
            self._lbl_out.setText("(рядом с исходником)")
            self._lbl_out.setStyleSheet("color: #777777; font-size: 11px;")
        self._save_settings()

    def _refresh_files(self):
        if not self._input_dir or not self._input_dir.exists():
            return
        files = sorted(
            list(self._input_dir.rglob("*.mid")) +
            list(self._input_dir.rglob("*.midi"))
        )
        self._files = [f for f in files if not f.stem.endswith("_ss14")]
        self._redraw_list()

    def _redraw_list(self):
        while self._file_layout.count():
            item = self._file_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._file_vars.clear()

        if not self._files:
            hint = QLabel("Файлы не найдены" if self._input_dir else "Перетащи папку или MIDI файлы сюда")
            hint.setObjectName("drop_hint")
            hint.setAlignment(Qt.AlignCenter)
            hint.setContentsMargins(0, 60, 0, 60)
            self._file_layout.addWidget(hint)
            self._lbl_count.setText("0 файлов")
            return

        for i, f in enumerate(self._files):
            row = QFrame()
            bg = "#151515" if i % 2 == 0 else "#191919"
            row.setStyleSheet(f"background-color: {bg}; border-radius: 3px;")
            rl = QHBoxLayout(row)
            rl.setContentsMargins(6, 3, 8, 3)
            rl.setSpacing(6)

            cb = QCheckBox()
            cb.setChecked(True)
            rl.addWidget(cb)
            self._file_vars.append(cb)

            done = self._get_out_path(f).exists()
            lbl = QLabel(f.name + ("  ✓" if done else ""))
            if done:
                lbl.setObjectName("file_done")
            else:
                lbl.setStyleSheet("font-size: 11px;")
            rl.addWidget(lbl, 1)

            try:
                rel = str(f.parent.relative_to(self._input_dir)) if self._input_dir else ""
                if rel and rel != ".":
                    r = QLabel(rel)
                    r.setStyleSheet("color: #444444; font-size: 9px;")
                    rl.addWidget(r)
            except Exception:
                pass

            self._file_layout.addWidget(row)

        self._lbl_count.setText(f"{len(self._files)} файлов")

    def _get_out_path(self, src: Path) -> Path:
        name = src.stem + "_ss14.mid"
        return (self._output_dir / name) if self._output_dir else (src.parent / name)

    def _select_all(self, state: bool):
        for cb in self._file_vars:
            cb.setChecked(state)

    # ── Лог ───────────────────────────────────────────────────────────────────

    def _on_log(self, text: str):
        self._log.append(text)
        sb = self._log.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _clear_log(self):
        self._log.clear()

    def _on_progress(self, done: int, total: int):
        self._progress.setValue(int(done / total * 100) if total else 0)
        self._lbl_status.setText(f"{done} / {total}")
        self._lbl_status.setStyleSheet("color: #777777; font-size: 10px;")

    def _on_ask(self, problems: str):
        dlg = QDialog(self)
        dlg.setWindowTitle("Повреждённый файл")
        dlg.setFixedSize(460, 220)
        dlg.setStyleSheet("QDialog { background-color: #1a1a1a; } QLabel { color: #e0e0e0; }")

        vl = QVBoxLayout(dlg)
        vl.setSpacing(8)

        t = QLabel("⚠ Файл повреждён")
        t.setAlignment(Qt.AlignCenter)
        t.setStyleSheet("font-size: 14px; font-weight: bold;")
        vl.addWidget(t)

        d = QLabel(problems)
        d.setWordWrap(True)
        d.setStyleSheet("color: #999999; font-size: 11px;")
        d.setAlignment(Qt.AlignCenter)
        vl.addWidget(d)

        q = QLabel("Результат может быть искажён. Продолжить?")
        q.setStyleSheet("color: #999999; font-size: 11px;")
        q.setAlignment(Qt.AlignCenter)
        vl.addWidget(q)
        vl.addSpacing(8)

        result = ['n']
        def answer(a):
            result[0] = a
            dlg.accept()

        bl = QHBoxLayout()
        bl.addStretch()
        for txt, a, w in [("Да", 'y', 80), ("Да для всех", 'ya', 120), ("Пропустить", 'n', 100)]:
            b = QPushButton(txt)
            b.setObjectName("secondary")
            b.setFixedWidth(w)
            b.clicked.connect(lambda _, x=a: answer(x))
            bl.addWidget(b)
        bl.addStretch()
        vl.addLayout(bl)

        dlg.exec()
        self._ask_answer_queue.put(result[0])

    def _on_done(self, ok: int, skip: int, fail: int):
        self._on_log(f"\n{'─' * 46}")
        self._on_log(f"Готово: {ok} успешно  {skip} пропущено  {fail} ошибок")
        self._set_running(False)
        self._progress.setValue(100 if (ok + skip + fail) > 0 else 0)
        color = "#bbbbbb" if fail == 0 else "#aa5555"
        self._lbl_status.setText(f"Готово: {ok} ✓  {skip} пропущено  {fail} ошибок")
        self._lbl_status.setStyleSheet(f"color: {color}; font-size: 10px;")
        self._redraw_list()

    # ── Запуск ────────────────────────────────────────────────────────────────

    def _set_running(self, state: bool):
        self._running = state
        self._btn_all.setEnabled(not state)
        self._btn_sel.setEnabled(not state)
        self._btn_stop.setEnabled(state)
        if not state:
            self._stop_flag = False

    def _stop(self):
        self._stop_flag = True
        self._signals.log.emit("\nОстановка после текущего файла...")

    def _run_all(self):
        self._select_all(True)
        self._run_selected()

    def _run_selected(self):
        if self._running:
            return
        targets = [f for f, cb in zip(self._files, self._file_vars) if cb.isChecked()]
        if not targets:
            QMessageBox.information(self, "Нет файлов", "Выбери папку и убедись что файлы отмечены.")
            return

        if len(targets) > 1:
            ret = QMessageBox.question(
                self, "Начать обработку?",
                f"Будет обработано файлов: {len(targets)}\n\nИнтерфейс останется отзывчивым.",
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            if ret != QMessageBox.StandardButton.Ok:
                return

        self._stop_flag = False
        self._yes_all = self._yes_all_cb.isChecked()
        self._set_running(True)
        self._progress.setValue(0)
        self._lbl_status.setText(f"0 / {len(targets)}")
        self._lbl_status.setStyleSheet("color: #777777; font-size: 10px;")
        self._save_settings()

        threading.Thread(
            target=self._worker,
            args=(targets, self._drum_slider.value(), self._melodic_slider.value(),
                  self._yes_all, self._output_dir),
            daemon=True).start()

    def _worker(self, files, drum_cap, melodic_cap, yes_all, out_dir):
        log = lambda msg: self._signals.log.emit(msg)
        ok = fail = skip = 0

        def ask_fn(problems):
            if yes_all or self._yes_all:
                return 'y'
            self._signals.ask.emit(problems)
            return self._ask_answer_queue.get()

        for i, f in enumerate(files):
            if self._stop_flag:
                break
            out = (out_dir / (f.stem + "_ss14.mid")) if out_dir \
                  else (f.parent / (f.stem + "_ss14.mid"))
            try:
                success, ya_req = optimize(f, out,
                                           yes_all=yes_all or self._yes_all,
                                           drum_cap=drum_cap, melodic_cap=melodic_cap,
                                           log_fn=log, ask_fn=ask_fn)
                if ya_req:
                    self._yes_all = True
                if success: ok += 1
                else:       skip += 1
            except Exception as e:
                log(f"  ОШИБКА: {e}")
                log(traceback.format_exc())
                fail += 1
            self._signals.progress.emit(i + 1, len(files))

        self._signals.done.emit(ok, skip, fail)


# ═══════════════════════════════════════════════════════════════════════════════
#  ТОЧКА ВХОДА
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    try:
        import mido
    except ImportError:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Зависимость не найдена",
                             "Установи mido:\n  pip install mido\n\nИли запусти install.bat")
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setStyleSheet(STYLESHEET)

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#141414"))
    palette.setColor(QPalette.WindowText, QColor("#e0e0e0"))
    palette.setColor(QPalette.Base, QColor("#1a1a1a"))
    palette.setColor(QPalette.AlternateBase, QColor("#1e1e1e"))
    palette.setColor(QPalette.ToolTipBase, QColor("#1e1e1e"))
    palette.setColor(QPalette.ToolTipText, QColor("#e0e0e0"))
    palette.setColor(QPalette.Text, QColor("#e0e0e0"))
    palette.setColor(QPalette.Button, QColor("#2a2a2a"))
    palette.setColor(QPalette.ButtonText, QColor("#e0e0e0"))
    palette.setColor(QPalette.Highlight, QColor("#555555"))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    QToolTip.setFont(QFont("Segoe UI", 9))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
