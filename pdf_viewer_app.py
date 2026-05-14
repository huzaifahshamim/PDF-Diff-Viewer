# -*- coding: UTF-8 -*-
import difflib
import html
import inspect
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import traceback
import unicodedata
from collections import defaultdict
from tkinter import ttk, filedialog, messagebox

import fitz
import klembord
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

try:
	import win32com.client
	import pythoncom
	from ctypes import windll, wintypes
	on_windows=1
except:
	windll = None
	wintypes = None
	on_windows=0

try:
	import pyautogui
	PYAUTOGUI_AVAILABLE = True
except ImportError:
	PYAUTOGUI_AVAILABLE = False



# python -m venv myenv
# myenv\Scripts\activate
# python -m pip install Pillow klembord tkinterdnd2 pywin32 pyinstaller pymupdf pyautogui
# python myenv\Scripts\pywin32_postinstall.py -install
# ren pdf_viewer_app.py pdf_viewer_app.pyw
# pyinstaller --noconfirm pdf_viewer_app.pyw
#
# the following files can be removed from "dist" created by pinstaller:
#
# del dist\pdf_viewer_app\_internal\libcrypto-3.dll
# del dist\pdf_viewer_app\_internal\libssl-3.dll
# del dist\pdf_viewer_app\_internal\MSVCP140.dll
# del dist\pdf_viewer_app\_internal\unicodedata.pyd
# del dist\pdf_viewer_app\_internal\PIL\_avif.cp312-win_amd64.pyd
# del dist\pdf_viewer_app\_internal\PIL\_webp.cp312-win_amd64.pyd
# rmdir /s /q dist\pdf_viewer_app\_internal\_tcl_data\tzdata
# rmdir /s /q dist\pdf_viewer_app\_internal\_tcl_data\encoding
# rmdir /s /q dist\pdf_viewer_app\_internal\idlelib
# rmdir /s /q dist\pdf_viewer_app\_internal\setuptools
# rmdir /s /q dist\pdf_viewer_app\_internal\tcl8

# copy minimal git and .dll to the main directory (dist\pdf_viewer_app)
# then, 7z the folder

# not working:
# python -m nuitka --mode=standalone --enable-plugin=tk-inter pdf_viewer_app.py


#TEMP_PDF_DIR = os.path.join(os.getcwd(), "temp_pdfs")
TEMP_PDF_DIR = os.path.join(os.path.dirname(__file__), "temp_pdfs")
os.makedirs(TEMP_PDF_DIR, exist_ok=True)
try:
	windll.user32.SetThreadDpiAwarenessContext(wintypes.HANDLE(-2))
except AttributeError:
	pass


class ToolTip:
	"""Lightweight tooltip replacement (does not depend on idlelib)."""

	def __init__(self, widget, text):
		self.widget = widget
		self.text = text
		self.tip_window = None
		widget.bind("<Enter>", self._show)
		widget.bind("<Leave>", self._hide)

	def _show(self, _event=None):
		if self.tip_window:
			return
		x = self.widget.winfo_rootx() + 20
		y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
		self.tip_window = tw = tk.Toplevel(self.widget)
		tw.wm_overrideredirect(True)
		tw.wm_geometry(f"+{x}+{y}")
		label = tk.Label(
			tw, text=self.text, justify=tk.LEFT,
			background="#ffffe0", relief=tk.SOLID, borderwidth=1,
			font=("TkDefaultFont", 9),
		)
		label.pack()

	def _hide(self, _event=None):
		if self.tip_window:
			self.tip_window.destroy()
			self.tip_window = None


def find_git_executable():
	"""Resolve git: bundled PortableGit next to the app/exe, then PATH."""
	if getattr(sys, "frozen", False):
		base = os.path.dirname(sys.executable)
	else:
		base = os.path.dirname(os.path.abspath(__file__))
	for rel in (
		os.path.join("git", "cmd", "git.exe"),
		os.path.join("PortableGit", "cmd", "git.exe"),
		os.path.join("mingw64", "bin", "git.exe"),
		os.path.join("cmd", "git.exe"),
	):
		candidate = os.path.join(base, rel)
		if os.path.isfile(candidate):
			return candidate
	found = shutil.which("git")
	return found if found else "git"


def normalize_token(text, case_insensitive, ignore_quotes):
	"""Normalize a single token for comparison."""
	t = unicodedata.normalize("NFKC", text)
	t = t.replace("\u00a0", " ").replace("\u200b", "").replace("\u00ad", "")
	if ignore_quotes:
		t = (
			t.replace("\u2018", "'").replace("\u2019", "'").replace("\u02bc", "'")
			.replace("\u201c", '"').replace("\u201d", '"')
		)
	if case_insensitive:
		t = t.casefold()
	return t


LEADING_PUNCTUATION = "([{\"'"
TRAILING_PUNCTUATION = ".,;:!?)]}'\""


def _split_bbox_for_segments(x0, y0, x1, y1, segments):
	"""Split a word bounding box proportionally by character count."""
	total_len = sum(len(s) for s in segments)
	if total_len == 0:
		return [(x0, y0, x1, y1)] * len(segments)
	width = x1 - x0
	rects = []
	cursor = x0
	for segment in segments:
		seg_width = width * (len(segment) / total_len)
		rects.append((cursor, y0, cursor + seg_width, y1))
		cursor += seg_width
	rects[-1] = (rects[-1][0], y0, x1, y1)
	return rects


def split_word_punctuation(word_dict):
	"""Split a word into core text and attached punctuation tokens."""
	text = word_dict["text"]
	if not text or len(text) == 1:
		return [word_dict]

	lead_len = 0
	while lead_len < len(text) and text[lead_len] in LEADING_PUNCTUATION:
		lead_len += 1
	rest = text[lead_len:]
	trail_len = 0
	while trail_len < len(rest) and rest[len(rest) - 1 - trail_len] in TRAILING_PUNCTUATION:
		trail_len += 1
	core = rest[: len(rest) - trail_len] if trail_len else rest
	trail = rest[len(rest) - trail_len :] if trail_len else ""

	segments = []
	if lead_len:
		segments.append(text[:lead_len])
	if core:
		segments.append(core)
	if trail:
		segments.append(trail)
	if len(segments) <= 1:
		return [word_dict]

	x0, y0, x1, y1 = word_dict["x0"], word_dict["y0"], word_dict["x1"], word_dict["y1"]
	rects = _split_bbox_for_segments(x0, y0, x1, y1, segments)
	result = []
	for segment, rect in zip(segments, rects):
		new_word = dict(word_dict)
		new_word["text"] = segment
		new_word["x0"], new_word["y0"], new_word["x1"], new_word["y1"] = rect
		result.append(new_word)
	return result


def merge_hyphenated_line_breaks(page_words):
	"""Merge words split across lines by a trailing hyphen or soft hyphen."""
	if not page_words:
		return page_words

	merged = []
	i = 0
	while i < len(page_words):
		current = page_words[i]
		if i + 1 < len(page_words):
			nxt = page_words[i + 1]
			current_text = current["text"]
			ends_hyphen = current_text.endswith("-") or current_text.endswith("\u00ad")
			same_block = current.get("block_idx") == nxt.get("block_idx")
			next_line = nxt.get("line_idx", 0) > current.get("line_idx", 0)
			vertical_gap = nxt["y0"] - current["y1"]
			line_height = max(current.get("font_size", 12.0), 1.0)
			close_lines = 0 <= vertical_gap <= line_height * 1.5
			if ends_hyphen and same_block and next_line and close_lines:
				strip_char = "\u00ad" if current_text.endswith("\u00ad") else "-"
				combined = dict(current)
				combined["text"] = current_text[: -len(strip_char)] + nxt["text"]
				combined["x0"] = min(current["x0"], nxt["x0"])
				combined["y0"] = min(current["y0"], nxt["y0"])
				combined["x1"] = max(current["x1"], nxt["x1"])
				combined["y1"] = max(current["y1"], nxt["y1"])
				merged.append(combined)
				i += 2
				continue
		merged.append(current)
		i += 1
	return merged


def postprocess_page_words(page_words, split_punctuation=True, merge_hyphenation=True):
	"""Apply hyphenation merge and punctuation splitting to one page of words."""
	words = merge_hyphenated_line_breaks(page_words) if merge_hyphenation else list(page_words)
	result = []
	for word in words:
		if split_punctuation:
			result.extend(split_word_punctuation(word))
		else:
			result.append(word)
	return result


def assign_chunk_ids(words_data, paragraph_gap_factor=1.8):
	"""Assign chunk_id boundaries using page/block changes and paragraph gaps."""
	if not words_data:
		return words_data
	chunk_id = 0
	prev = None
	for word in words_data:
		if prev is None:
			word["chunk_id"] = chunk_id
		else:
			new_chunk = False
			if word["page_num"] != prev["page_num"]:
				new_chunk = True
			elif word.get("block_idx") != prev.get("block_idx"):
				new_chunk = True
			else:
				gap = word["y0"] - prev["y1"]
				line_height = max(prev.get("font_size", 12.0), 1.0)
				if gap > line_height * paragraph_gap_factor:
					new_chunk = True
			if new_chunk:
				chunk_id += 1
			word["chunk_id"] = chunk_id
		prev = word
	return words_data


def apply_opcodes_to_words(words_data1, words_data2, opcodes, with_moves=False, reset=True, id_counter_start=0):
	"""Map diff opcodes to highlight colors and sync-scroll ids (index-based)."""
	if reset:
		for w in words_data1:
			w["unique_id"] = None
			w["highlight_color"] = None
			w["move_id"] = None
			w["move_side"] = None
		for w in words_data2:
			w["unique_id"] = None
			w["highlight_color"] = None
			w["move_id"] = None
			w["move_side"] = None

	common_word_id_counter = id_counter_start
	for op in opcodes:
		if with_moves:
			tag, i1, i2, j1, j2, is_moved = op
		else:
			tag, i1, i2, j1, j2 = op[:5]
			is_moved = False

		if tag == "equal":
			for k in range(i2 - i1):
				common_id = f"common-word-{common_word_id_counter}"
				words_data1[i1 + k]["unique_id"] = common_id
				words_data2[j1 + k]["unique_id"] = common_id
				common_word_id_counter += 1
		elif tag == "delete":
			color = "blue" if is_moved else "red"
			for k in range(i1, i2):
				words_data1[k]["highlight_color"] = color
		elif tag == "insert":
			color = "blue" if is_moved else "green"
			for k in range(j1, j2):
				words_data2[k]["highlight_color"] = color
		elif tag == "replace":
			for k in range(i1, i2):
				words_data1[k]["highlight_color"] = "red"
			for k in range(j1, j2):
				words_data2[k]["highlight_color"] = "green"
	return words_data1, words_data2, common_word_id_counter


def _words_to_text(words, start, end):
	return " ".join(words[i]["text"] for i in range(start, end))


def _hunk_signature(words, start, end, case_insensitive, ignore_quotes):
	return " ".join(
		normalize_token(words[i]["text"], case_insensitive, ignore_quotes)
		for i in range(start, end)
	)


def _location_from_words(words, start, end):
	first = words[start]
	last = words[end - 1]
	return {
		"page": first["page_num"],
		"chunk_id": first.get("chunk_id", 0),
		"y0": first["y0"],
		"y1": last["y1"],
		"x0": min(words[i]["x0"] for i in range(start, end)),
		"x1": max(words[i]["x1"] for i in range(start, end)),
	}


def _context_snippet(words, index, direction, max_words=20):
	"""Return nearby unchanged words as context (before index if direction<0, after if >0)."""
	parts = []
	if direction < 0:
		i = index - 1
		while i >= 0 and len(parts) < max_words:
			if not words[i].get("highlight_color"):
				parts.insert(0, words[i]["text"])
			elif parts:
				break
			i -= 1
	else:
		i = index
		while i < len(words) and len(parts) < max_words:
			if not words[i].get("highlight_color"):
				parts.append(words[i]["text"])
			elif parts:
				break
			i += 1
	return " ".join(parts)


def pair_move_ids(words_left, words_right, case_insensitive, ignore_quotes):
	"""Assign shared move_id to paired blue (moved) hunks on left and right."""
	for w in words_left:
		w["move_id"] = None
		w["move_side"] = None
	for w in words_right:
		w["move_id"] = None
		w["move_side"] = None

	def collect_blue_hunks(words):
		hunks = []
		i = 0
		while i < len(words):
			if words[i].get("highlight_color") != "blue":
				i += 1
				continue
			start = i
			while i < len(words) and words[i].get("highlight_color") == "blue":
				i += 1
			hunks.append({
				"start": start,
				"end": i,
				"signature": _hunk_signature(words, start, i, case_insensitive, ignore_quotes),
			})
		return hunks

	right_buckets = defaultdict(list)
	for hunk in collect_blue_hunks(words_right):
		right_buckets[hunk["signature"]].append(hunk)

	move_counter = 0
	for left_hunk in collect_blue_hunks(words_left):
		bucket = right_buckets.get(left_hunk["signature"], [])
		if not bucket:
			continue
		right_hunk = bucket.pop(0)
		move_id = f"move-{move_counter}"
		move_counter += 1
		for idx in range(left_hunk["start"], left_hunk["end"]):
			words_left[idx]["move_id"] = move_id
			words_left[idx]["move_side"] = "from"
		for idx in range(right_hunk["start"], right_hunk["end"]):
			words_right[idx]["move_id"] = move_id
			words_right[idx]["move_side"] = "to"


def _collect_color_hunks(words, color):
	hunks = []
	i = 0
	while i < len(words):
		if words[i].get("highlight_color") != color:
			i += 1
			continue
		start = i
		while i < len(words) and words[i].get("highlight_color") == color:
			i += 1
		hunks.append({"start": start, "end": i, "color": color})
	return hunks


def build_change_records(
	words_left,
	words_right,
	left_name="left",
	right_name="right",
	case_insensitive=True,
	ignore_quotes=True,
	compare_options=None,
	group_by_chunk=True,
):
	"""
	Build structured change records from aligned word lists.
	Returns a dict suitable for JSON export and agent consumption.
	Alignment is always full-document; group_by_chunk only affects how
	added/deleted/modified hunks are paired in the export.
	"""
	pair_move_ids(words_left, words_right, case_insensitive, ignore_quotes)

	changes = []
	change_counter = 0
	seen_move_ids = set()

	# Moved hunks (paired by move_id)
	move_groups = defaultdict(lambda: {"from": None, "to": None})
	for side_key, words in (("from", words_left), ("to", words_right)):
		for hunk in _collect_color_hunks(words, "blue"):
			move_id = words[hunk["start"]].get("move_id")
			if not move_id:
				continue
			move_groups[move_id][side_key] = hunk

	for move_id, group in sorted(move_groups.items(), key=lambda x: x[0]):
		if move_id in seen_move_ids:
			continue
		seen_move_ids.add(move_id)
		left_hunk = group["from"]
		right_hunk = group["to"]
		left_text = _words_to_text(words_left, left_hunk["start"], left_hunk["end"]) if left_hunk else None
		right_text = _words_to_text(words_right, right_hunk["start"], right_hunk["end"]) if right_hunk else None
		loc_source = left_hunk or right_hunk
		words_ref = words_left if left_hunk else words_right
		changes.append({
			"id": f"chg-{change_counter:04d}",
			"type": "moved",
			"move_id": move_id,
			"left_text": left_text,
			"right_text": right_text,
			"left_location": _location_from_words(words_left, left_hunk["start"], left_hunk["end"]) if left_hunk else None,
			"right_location": _location_from_words(words_right, right_hunk["start"], right_hunk["end"]) if right_hunk else None,
			"context_before": _context_snippet(words_ref, loc_source["start"], -1) if loc_source else "",
			"context_after": _context_snippet(words_ref, loc_source["end"], 1) if loc_source else "",
		})
		change_counter += 1

	def append_red_green_changes(left_red, right_green, chunk_id=None):
		nonlocal change_counter
		pairs = min(len(left_red), len(right_green))
		for i in range(pairs):
			lh, rh = left_red[i], right_green[i]
			entry = {
				"id": f"chg-{change_counter:04d}",
				"type": "modified",
				"left_text": _words_to_text(words_left, lh["start"], lh["end"]),
				"right_text": _words_to_text(words_right, rh["start"], rh["end"]),
				"left_location": _location_from_words(words_left, lh["start"], lh["end"]),
				"right_location": _location_from_words(words_right, rh["start"], rh["end"]),
				"left_word_range": [lh["start"], lh["end"]],
				"right_word_range": [rh["start"], rh["end"]],
				"context_before": _context_snippet(words_left, lh["start"], -1),
				"context_after": _context_snippet(words_left, lh["end"], 1),
			}
			if chunk_id is not None:
				entry["chunk_id"] = chunk_id
			changes.append(entry)
			change_counter += 1
		for lh in left_red[pairs:]:
			entry = {
				"id": f"chg-{change_counter:04d}",
				"type": "deleted",
				"left_text": _words_to_text(words_left, lh["start"], lh["end"]),
				"right_text": None,
				"left_location": _location_from_words(words_left, lh["start"], lh["end"]),
				"right_location": None,
				"left_word_range": [lh["start"], lh["end"]],
				"right_word_range": None,
				"context_before": _context_snippet(words_left, lh["start"], -1),
				"context_after": _context_snippet(words_left, lh["end"], 1),
			}
			if chunk_id is not None:
				entry["chunk_id"] = chunk_id
			changes.append(entry)
			change_counter += 1
		for rh in right_green[pairs:]:
			entry = {
				"id": f"chg-{change_counter:04d}",
				"type": "added",
				"left_text": None,
				"right_text": _words_to_text(words_right, rh["start"], rh["end"]),
				"left_location": None,
				"right_location": _location_from_words(words_right, rh["start"], rh["end"]),
				"left_word_range": None,
				"right_word_range": [rh["start"], rh["end"]],
				"context_before": _context_snippet(words_right, rh["start"], -1),
				"context_after": _context_snippet(words_right, rh["end"], 1),
			}
			if chunk_id is not None:
				entry["chunk_id"] = chunk_id
			changes.append(entry)
			change_counter += 1

	if group_by_chunk:
		all_chunk_ids = sorted({
			w.get("chunk_id", 0) for w in words_left + words_right
		})
		for chunk_id in all_chunk_ids:
			left_red = [
				h for h in _collect_color_hunks(words_left, "red")
				if words_left[h["start"]].get("chunk_id", 0) == chunk_id
			]
			right_green = [
				h for h in _collect_color_hunks(words_right, "green")
				if words_right[h["start"]].get("chunk_id", 0) == chunk_id
			]
			append_red_green_changes(left_red, right_green, chunk_id=chunk_id)
	else:
		append_red_green_changes(
			_collect_color_hunks(words_left, "red"),
			_collect_color_hunks(words_right, "green"),
		)

	# Sort changes by left location (fallback to right) for document order
	def sort_key(chg):
		loc = chg.get("left_location") or chg.get("right_location") or {}
		return (loc.get("page", 0), loc.get("chunk_id", 0), loc.get("y0", 0.0))

	changes.sort(key=sort_key)
	for i, chg in enumerate(changes):
		chg["id"] = f"chg-{i:04d}"

	stats = {
		"added": sum(1 for c in changes if c["type"] == "added"),
		"deleted": sum(1 for c in changes if c["type"] == "deleted"),
		"modified": sum(1 for c in changes if c["type"] == "modified"),
		"moved": sum(1 for c in changes if c["type"] == "moved"),
		"total": len(changes),
		"unchanged_words": sum(1 for w in words_left if not w.get("highlight_color")),
	}

	return {
		"version": 1,
		"left_document": left_name,
		"right_document": right_name,
		"compare_options": compare_options or {},
		"stats": stats,
		"changes": changes,
	}


def write_change_records_json(records, output_path):
	"""Write change records to a JSON file."""
	output_path = os.path.abspath(output_path)
	os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
	payload = dict(records)
	payload["export_path"] = output_path
	with open(output_path, "w", encoding="utf-8") as f:
		json.dump(payload, f, ensure_ascii=False, indent=2)
	print(f"Change records written to: {output_path}")
	return output_path


def _sanitize_filename_part(name):
	"""Make a string safe for use in a filename (Windows-compatible)."""
	clean = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
	return clean or "document"


def default_change_records_json_path(left_name, right_name):
	"""Default auto-export path under temp_pdfs/."""
	n1 = _sanitize_filename_part(os.path.splitext(left_name or "left")[0])
	n2 = _sanitize_filename_part(os.path.splitext(right_name or "right")[0])
	return os.path.join(TEMP_PDF_DIR, f"{n1}_vs_{n2}_changes.json")


def auto_save_change_records(records, left_name, right_name):
	"""Write change records JSON to the default export path."""
	path = default_change_records_json_path(left_name, right_name)
	return write_change_records_json(records, path)


def compare_files_to_change_records(
	left_path,
	right_path,
	case_insensitive=True,
	ignore_quotes=True,
	ignore_ligatures=True,
	split_punctuation=True,
	merge_hyphenation=True,
	use_chunked=True,
):
	"""Headless: open two files, align, and return change records dict."""
	left_path = os.path.abspath(left_path)
	right_path = os.path.abspath(right_path)
	doc_left = fitz.open(left_path)
	doc_right = fitz.open(right_path)
	try:
		words_left = extract_words_with_styles(
			doc_left,
			ignore_ligatures=ignore_ligatures,
			split_punctuation=split_punctuation,
			merge_hyphenation=merge_hyphenation,
		)
		words_right = extract_words_with_styles(
			doc_right,
			ignore_ligatures=ignore_ligatures,
			split_punctuation=split_punctuation,
			merge_hyphenation=merge_hyphenation,
		)
		words_left = [dict(w) for w in words_left]
		words_right = [dict(w) for w in words_right]
		words_left, words_right = align_words(
			words_left, words_right,
			case_insensitive, ignore_quotes,
			use_chunked=use_chunked,
		)
		compare_options = {
			"case_insensitive": case_insensitive,
			"ignore_quotes": ignore_quotes,
			"ignore_ligatures": ignore_ligatures,
			"split_punctuation": split_punctuation,
			"merge_hyphenation": merge_hyphenation,
			"group_by_chunk": use_chunked,
			"aligner": "git" if _git_available else "difflib",
		}
		return build_change_records(
			words_left, words_right,
			left_name=os.path.basename(left_path),
			right_name=os.path.basename(right_path),
			case_insensitive=case_insensitive,
			ignore_quotes=ignore_quotes,
			compare_options=compare_options,
			group_by_chunk=use_chunked,
		)
	finally:
		doc_left.close()
		doc_right.close()


def convert_clipboard_to_pdf(output_filename="clipboard_content.pdf"):
	"""
	Converts the HTML content from the clipboard to a PDF.
	If no HTML is found, it uses the plain text content.
	"""
	try:
		klembord.init()
	except RuntimeError:
		print("Error: Could not initialize klembord. Make sure a display server is running (e.g., X server on Linux).", file=sys.stderr)
		return None
	html_content = None
	plain_text_content = None
	try:
		plain_text_content, html_content = klembord.get_with_rich_text()
	except Exception as e:
		print(f"Warning: Could not retrieve rich text from clipboard: {e}", file=sys.stderr)
		print("Attempting to get plain text only.", file=sys.stderr)
		plain_text_content = klembord.get_text()
	content_to_use = ""
	if html_content:
		content_to_use = html_content
		if content_to_use.lower().find("<html")!=-1:
			content_to_use=content_to_use[content_to_use.lower().find("<html"):]
		elif content_to_use.lower().find("<head")!=-1:
			content_to_use=content_to_use[content_to_use.lower().find("<head"):]
		print("Using HTML content from clipboard.")
		#
		# Regex to find 'style="..."' or 'style='...'
		# And then replace 'background:' within that capture group
		# This is a more complex but more precise regex approach.
		# It looks for style attributes and then performs a sub-replacement inside the matched style content.
		def replace_style_content(match):
			style_content = match.group(1) # The content inside the style attribute quotes
			# Now, replace background: with background-color: within this specific style content
			# using a nested re.sub, case-insensitively
			new_style_content = re.sub(r'(?i)background:', 'background-color:', style_content)
			return f'style="{new_style_content}"' # Reconstruct the style attribute
		# This regex captures the content of the style attribute (between the quotes).
		# We handle both single and double quotes for the style attribute value.
		# It's still not perfect for all edge cases (e.g., mismatched quotes or escaped quotes within style)
		# but is much better than a global replace.
		content_to_use = re.sub(
			r'style=["\'](.*?)["\']', # Capture everything inside style="..." or style='...'
			replace_style_content,	# Use a function to process the captured content
			content_to_use,
			flags=re.DOTALL | re.IGNORECASE # DOTALL to match across newlines, IGNORECASE for 'style' itself
		)
	elif plain_text_content:
		escaped = html.escape(plain_text_content)
		content_to_use = f"""
		<html>
		<head>
			<style>
				body {{
					font-family: monospace;
					white-space: pre-wrap;
					word-wrap: break-word;
				}}
			</style>
		</head>
		<body>
			<div>{escaped}</div>
		</body>
		</html>
		"""
		print("Using plain text content from clipboard (wrapped in <pre> tags).")
	else:
		print("Clipboard is empty or contains no readable content.", file=sys.stderr)
		return None
	try:
		pathlib.Path(output_filename).parent.mkdir(parents=True, exist_ok=True)
		story = fitz.Story(html=content_to_use)  
		writer = fitz.DocumentWriter(output_filename)
		mediabox = fitz.paper_rect("a4")  
		where = mediabox + (36, 36, -36, -36)  
		more = True
		page_number = 0
		while more:  
			page_number += 1
			dev = writer.begin_page(mediabox)  
			more, filled = story.place(where)  
			story.draw(dev)  
			writer.end_page()
		writer.close()
		print(f"Clipboard content converted to PDF: {output_filename}")
		return output_filename
	except Exception as e:
		print(f"Error converting clipboard content to PDF: {e}", file=sys.stderr)
		return None


def convert_word_to_pdf_no_markup(input_file_path, output_pdf_path=None):
	"""
	Converts a Word .docx, .doc, .rtf, or .txt file to a PDF by setting various view and print options
	to hide markup, then using the SaveAs method.
	The original file is not modified.
	Crucially, this function aims to ensure it does NOT interfere with any existing
	user-opened Word instances or unsaved user documents.
	Requires Microsoft Word to be installed on a Windows system.

	Args:
		input_file_path (str): The full path to the input file.
		output_pdf_path (str, optional): The full path for the output PDF file.
										  If None, a temporary name will be generated
										  in the TEMP_PDF_DIR.
	Returns:
		str: The path to the generated PDF file, or None if conversion fails.
	"""
	if not on_windows:
		print("Only working on Windows and requiring pythoncom and win32com.client")
		return
	
	input_file_path = input_file_path.replace("/", "\\")
	if output_pdf_path:
		output_pdf_path = output_pdf_path.replace("/", "\\")

	if not os.path.exists(input_file_path):
		print(f"Error: Input file not found at {input_file_path}")
		return None

	if output_pdf_path is None:
		base_name = os.path.splitext(os.path.basename(input_file_path))[0]
		output_pdf_path = os.path.join(TEMP_PDF_DIR, f"{base_name}_temp_{os.urandom(4).hex()}.pdf")

	os.makedirs(TEMP_PDF_DIR, exist_ok=True)

	wdFormatPDF = 17
	wdRevisionsViewFinal = 0

	word_app = None
	doc = None

	try:
		print(f"--- Starting conversion for '{input_file_path}' to '{output_pdf_path}' ---")
		pythoncom.CoInitialize()

		# Always dispatch a new instance of Word.
		word_app = win32com.client.DispatchEx("Word.Application")
		print("Attempted to launch a new Word instance for conversion.")
		
		word_app.Visible = False
		word_app.DisplayAlerts = False

		doc = word_app.Documents.Open(str(input_file_path))

		# Set the WarnBeforeSavingPrintingSendingMarkup option
		#
		# still to fix this
		# An error occurred during conversion: (-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The WarnBeforeSavingPrintingSendingMarkup method or property is not available because the current document is read-only.', 'wdmain11.chm', 37373, -2146823683), None)
		# (when opening a read-only file) 
		#
		if hasattr(word_app.Options, 'WarnBeforeSavingPrintingSendingMarkup'):
			word_app.Options.WarnBeforeSavingPrintingSendingMarkup = False
			print("1. Set word_app.Options.WarnBeforeSavingPrintingSendingMarkup to False.")
		else:
			print("Warning: 'WarnBeforeSavingPrintingSendingMarkup' property not found. Skipping.")

		# Set revision view for the document opened in this specific instance
		if doc.ActiveWindow:
			doc.ActiveWindow.View.RevisionsView = wdRevisionsViewFinal
			print("2. Set document view to 'No Markup' (ActiveWindow.View.RevisionsView).")
		else:
			print("Warning: ActiveWindow not found for the document. Could not set revision view.")

		if hasattr(doc, 'ShowRevisions'):
			doc.ShowRevisions = False
			print("3. Set doc.ShowRevisions to False.")
		else:
			print("Warning: 'ShowRevisions' property not found on document. Skipping.")

		# Set print options for this specific Word instance
		if hasattr(word_app.Options, 'PrintRevisions'):
			word_app.Options.PrintRevisions = False
			print("4. Set word_app.Options.PrintRevisions to False.")
		else:
			print("Warning: 'PrintRevisions' property not found on word_app.Options. Skipping.")

		if hasattr(word_app.Options, 'PrintComments'):
			word_app.Options.PrintComments = False
			print("5. Set word_app.Options.PrintComments to False.")
		else:
			print("Warning: 'PrintComments' property not found on word_app.Options. Skipping.")
		
		if hasattr(word_app.Options, 'PrintHiddenText'):
			word_app.Options.PrintHiddenText = False
			print("6. Set word_app.Options.PrintHiddenText to False.")
		else:
			print("Warning: 'PrintHiddenText' property not found. Skipping.")

		if hasattr(word_app.Options, 'PrintDrawingObjects'):
			word_app.Options.PrintDrawingObjects = True # Usually want drawings
			print("7. Set word_app.Options.PrintDrawingObjects to True.")
		else:
			print("Warning: 'PrintDrawingObjects' property not found. Skipping.")

		doc.SaveAs(str(output_pdf_path), FileFormat=wdFormatPDF)
		print(f"Document saved as PDF to: {output_pdf_path}")
		
		# Close only the document that was opened by this script instance.
		doc.Close(SaveChanges=False) # SaveChanges=False is crucial
		print("Document closed within the isolated Word instance.")
		print("--- Conversion successful! ---")
		return output_pdf_path

	except Exception as e:
		print(f"An error occurred during conversion: {e}")
		print(f"Error details (e.args): {e.args}")
		try:
			excepinfo = pythoncom.GetErrorInfo()
			if excepinfo:
				print(f"COM Error Info: Source={excepinfo[0]}, Description={excepinfo[2]}")
		except Exception:
			pass
		return None

	finally:
		if word_app:
			try:
				word_app.Quit(SaveChanges=0) # wdDoNotSaveChanges = 0
				print("Isolated Word application instance quit.")
			except Exception as e:
				print(f"Error quitting Word application: {e}")

		pythoncom.CoUninitialize()


def extract_words_with_styles(
	pdf_document,
	ignore_ligatures=False,
	split_punctuation=True,
	merge_hyphenation=True,
):
	"""
	Extracts all words from a PyMuPDF document with their coordinates.
	Uses dict block/line order for reading order, with scaled line tolerance.
	"""
	all_words_data = []

	for page_num, page in enumerate(pdf_document):
		page.remove_rotation()
		if ignore_ligatures:
			words_raw = page.get_text("words", flags=0)
			text_dict = page.get_text("dict", flags=0)
		else:
			words_raw = page.get_text("words")
			text_dict = page.get_text("dict")
		if not words_raw:
			continue

		ordered_lines = []
		for block_idx, block in enumerate(text_dict.get("blocks", [])):
			if block.get("type") != 0:
				continue
			for line_idx, line in enumerate(block.get("lines", [])):
				bbox = line["bbox"]
				ordered_lines.append({
					"block_idx": block_idx,
					"line_idx": line_idx,
					"bbox": fitz.Rect(bbox),
					"sort_y": bbox[1],
					"sort_x": bbox[0],
				})

		heights = [w[3] - w[1] for w in words_raw if w[3] > w[1]]
		median_height = sorted(heights)[len(heights) // 2] if heights else 12.0
		line_tolerance = max(3.0, median_height * 0.35)

		if not ordered_lines:
			ordered_lines = [{
				"block_idx": 0,
				"line_idx": 0,
				"bbox": fitz.Rect(words_raw[0][0], words_raw[0][1], words_raw[0][2], words_raw[0][3]),
				"sort_y": words_raw[0][1],
				"sort_x": words_raw[0][0],
			}]

		line_words = defaultdict(list)
		for word_info in words_raw:
			x0, y0, x1, y1, word_text = word_info[:5]
			word_cy = (y0 + y1) / 2
			best_line = None
			best_dist = float("inf")
			for li, line in enumerate(ordered_lines):
				lb = line["bbox"]
				if lb.y0 - line_tolerance <= word_cy <= lb.y1 + line_tolerance:
					dist = abs(word_cy - (lb.y0 + lb.y1) / 2)
					if dist < best_dist:
						best_dist = dist
						best_line = li
			if best_line is None:
				ordered_lines.append({
					"block_idx": 9999,
					"line_idx": len(ordered_lines),
					"bbox": fitz.Rect(x0, y0, x1, y1),
					"sort_y": y0,
					"sort_x": x0,
				})
				best_line = len(ordered_lines) - 1
			line_words[best_line].append(word_info)

		def line_order_key(li):
			line = ordered_lines[li]
			return (line["sort_y"], line["sort_x"], line["block_idx"], line["line_idx"])

		page_words = []
		for li in sorted(line_words.keys(), key=line_order_key):
			line_meta = ordered_lines[li]
			for word_info in sorted(line_words[li], key=lambda w: w[0]):
				x0, y0, x1, y1, word_text = word_info[:5]
				page_words.append({
					"text": word_text,
					"x0": x0, "y0": y0, "x1": x1, "y1": y1,
					"page_num": page_num,
					"block_idx": line_meta["block_idx"],
					"line_idx": line_meta["line_idx"],
					"font_family": "",
					"font_size": median_height,
					"font_color": "#000000",
					"font_weight": "normal",
					"font_style": "normal",
					"unique_id": None,
					"highlight_color": None,
					"move_id": None,
					"move_side": None,
					"chunk_id": 0,
				})

		page_words = postprocess_page_words(
			page_words,
			split_punctuation=split_punctuation,
			merge_hyphenation=merge_hyphenation,
		)
		all_words_data.extend(page_words)

	assign_chunk_ids(all_words_data)
	return all_words_data

def helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes):
	a_compare = [
		normalize_token(word_info["text"], case_insensitive, ignore_quotes)
		for word_info in words_data1
	]
	b_compare = [
		normalize_token(word_info["text"], case_insensitive, ignore_quotes)
		for word_info in words_data2
	]
	return a_compare, b_compare

def align_words_with_difflib(words_data1, words_data2, case_insensitive, ignore_quotes):
	"""Aligns two word sequences using difflib.SequenceMatcher (Ratcliff-Obershelp)."""
	a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
	s = difflib.SequenceMatcher(None, a_compare, b_compare)
	opcodes = list(s.get_opcodes())
	words_data1, words_data2, _ = apply_opcodes_to_words(words_data1, words_data2, opcodes, with_moves=False)
	return words_data1, words_data2
def apply_annotations_to_pdf_pages(pdf_document, words_data):
	if not pdf_document or pdf_document.is_closed:
		return
	words_by_page = defaultdict(list)
	for word in words_data:
		if word["highlight_color"]:
			words_by_page[word["page_num"]].append(word)
	for page_num in range(pdf_document.page_count):
		page = pdf_document.load_page(page_num)
		annotations_to_delete = [
			annot for annot in page.annots()
			if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get("title") == "PDFComparer"
		]
		for annot in annotations_to_delete:
			try:
				page.delete_annot(annot)
			except Exception as e:
				print(f"Error deleting old annotation on page {page_num}: {e}")
		page_words = words_by_page[page_num]
		if not page_words:
			print(time.time(), f"page {page_num}, number of annotations: 0 (no words to highlight)")
			continue
		#words_by_page is already sorted (with a smarter logic than this one commmented, for better detection of rows)
		#page_words.sort(key=lambda w: (w["y0"], w["x0"]))
		highlights_by_color = defaultdict(list)
		for word in page_words:
			rect = fitz.Rect(word["x0"], word["y0"], word["x1"], word["y1"])
			highlights_by_color[word["highlight_color"]].append(rect)
		total_annotations_added = 0
		for color, rects_to_merge in highlights_by_color.items():
			if not rects_to_merge:
				continue
			merged_rects = []
			if rects_to_merge: 
				current_merged_rect = rects_to_merge[0]
				for i in range(1, len(rects_to_merge)):
					next_rect = rects_to_merge[i]
					y_tolerance = 10 
					x_tolerance = 10 
					if (abs(current_merged_rect.y0 - next_rect.y0) < y_tolerance and
						abs(current_merged_rect.y1 - next_rect.y1) < y_tolerance and
						next_rect.x0 <= current_merged_rect.x1 + x_tolerance): 
						current_merged_rect = current_merged_rect | next_rect 
					else:
						merged_rects.append(current_merged_rect)
						current_merged_rect = next_rect
				merged_rects.append(current_merged_rect) 
			highlight_color_rgb_float = (0.0, 0.0, 0.0)
			if color == "red":
				highlight_color_rgb_float = (1.0, 0.0, 0.0)
			elif color == "green":
				highlight_color_rgb_float = (0.0, 1.0, 0.0)
			elif color == "blue": 
				highlight_color_rgb_float = (0.0, 0.5, 1.0)
			for merged_rect in merged_rects:
				try:
					annot = page.add_highlight_annot(merged_rect)
					annot.set_colors(stroke=highlight_color_rgb_float)
					annot.set_opacity(0.3)
					annot.set_info(title="PDFComparer")
					annot.update()
					total_annotations_added += 1
				except Exception as e:
					print(f"Error adding merged highlight annotation on page {page_num}: {e}")
		print(time.time(), f"page {page_num}, number of annotations: {total_annotations_added} (from {len(page_words)} original words)")
	print(time.time(), "fine di " + inspect.currentframe().f_code.co_name)


class GitSequenceMatcher:
	GIT_COLOR_CONFIG = [
		"-c", "color.ui=always",
		"-c", "color.diff.old=red",
		"-c", "color.diff.new=green",
		"-c", "color.diff.meta=yellow",
	]

	def __init__(self, a, b, temp_dir=None, git_executable=None):
		self.a = a
		self.b = b
		self.temp_file_a = None
		self.temp_file_b = None
		self.temp_dir = temp_dir or TEMP_PDF_DIR
		self.git_executable = git_executable or find_git_executable()

	def _create_temp_files(self):
		"""Creates temporary files with one JSON-encoded token per line."""
		with tempfile.NamedTemporaryFile(mode="w+", delete=False, encoding="utf-8", dir=self.temp_dir) as f_a:
			self.temp_file_a = f_a.name
			for item in self.a:
				f_a.write(json.dumps(item, ensure_ascii=True) + "\n")

		with tempfile.NamedTemporaryFile(mode="w+", delete=False, encoding="utf-8", dir=self.temp_dir) as f_b:
			self.temp_file_b = f_b.name
			for item in self.b:
				f_b.write(json.dumps(item, ensure_ascii=True) + "\n")

	def _cleanup_temp_files(self):
		"""Removes the temporary files."""
		if self.temp_file_a and os.path.exists(self.temp_file_a):
			os.remove(self.temp_file_a)
		if self.temp_file_b and os.path.exists(self.temp_file_b):
			os.remove(self.temp_file_b)

	def get_opcodes(self):
		"""
		Generates a list of 6-tuple opcodes (tag, i1, i2, j1, j2, is_moved)
		similar to difflib.SequenceMatcher, with 'is_moved' flag for delete/insert.
		"""
		self._create_temp_files()
		process = None
		start_time=time.time()
		try:
			command = [
				self.git_executable,
				*self.GIT_COLOR_CONFIG,
				"--no-pager",
				"diff",
				"--no-index",
				"--no-ext-diff",
				"--diff-algorithm=histogram",
				"--color=always",
				"--color-moved",
				"--unified=99999999",
				self.temp_file_a,
				self.temp_file_b,
			]
			print(f"\nRunning command: {' '.join(command)}")
			process = subprocess.run(
				command,
				capture_output=True,
				text=True,
				encoding='utf-8',
				errors='replace'
			)

			# print(f"Git diff return code: {process.returncode}")
			# print("--- Raw Git Diff Output (repr) ---")
			# print(repr(process.stdout))
			# print("--- End Raw Git Diff Output ---")
			# print("Stderr from git (if any):")
			# print(process.stderr)
			# print("--- End Stderr ---")

			diff_output = process.stdout


			if process.returncode == 0 and not diff_output.strip():# in case the two files are equal and therefore git diff returns 0 and empty stdout
				with open(self.temp_file_a, 'r', encoding='utf-8', errors='replace') as f:
					num_lines = sum(1 for _ in f)
				return [('equal', 0, num_lines, 0, num_lines, False)]


			COLOR_RED_FG = r'\x1b\[31m'# deletions
			COLOR_GREEN_FG = r'\x1b\[32m'# insertions
			COLOR_BOLD_MAGENTA_FG = r'\x1b\[1;35m'# deletions (move)
			COLOR_BLUE_FG =		 r'\x1b\[1;34m'# deletions (move)
			COLOR_BOLD_CYAN_FG = r'\x1b\[1;36m'# insertions (move)
			COLOR_BOLD_YELLOW_FG = r'\x1b\[1;33m'# insertions (move)
			COLOR_RED_BG = r'\x1b\[41m'

			current_a_idx = 0
			current_b_idx = 0

			lines = diff_output.splitlines()
			in_hunk = False

			# Stores granular changes with internal tags and content
			# (internal_tag, content, a_start, a_end, b_start, b_end)
			granular_changes = []

			print("\n--- Line-by-Line Parsing Debug ---")
			for line_num, line in enumerate(lines):
				line_without_ansi = re.sub(r'\x1b\[[0-9;]*m', '', line)

				if line.startswith('\x1b[1mdiff --git'):
					in_hunk = True
					continue
				if not in_hunk:
					continue

				if line_without_ansi.strip().startswith('index ') or \
				   line_without_ansi.strip().startswith('--- a/') or \
				   line_without_ansi.strip().startswith('+++ b/'):
					continue
				
				if line_without_ansi.strip().startswith('@@'):
					match = re.match(r'@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@', line_without_ansi.strip())
					if match:
						current_a_idx = int(match.group(1)) - 1
						current_b_idx = int(match.group(3)) - 1
					else:
						print(f"  -> ERROR: '@@' line did not match regex: {repr(line_without_ansi.strip())}")
					continue
				
				tag = None
				content_to_match = ''

				if re.search(f'^{COLOR_BOLD_MAGENTA_FG}-', line) or re.search(f'^{COLOR_BLUE_FG}-', line):
					tag = 'moved_delete'
					content_to_match = line_without_ansi[1:].strip()
				elif re.search(f'^{COLOR_BOLD_CYAN_FG}\\+', line) or re.search(f'^{COLOR_BOLD_YELLOW_FG}\\+', line) or line.strip().endswith(f'{COLOR_RED_BG}'):
					tag = 'moved_insert'
					content_to_match = line_without_ansi[1:].strip()
				elif re.search(f'^{COLOR_RED_FG}-', line):
					tag = 'delete'
					content_to_match = line_without_ansi[1:].strip()
				elif re.search(f'^{COLOR_GREEN_FG}\\+', line):
					tag = 'insert'
					content_to_match = line_without_ansi[1:].strip()
				elif line_without_ansi.startswith(' '):
					tag = 'equal'
					content_to_match = line_without_ansi[1:].strip()
				else:
					if not line_without_ansi.strip():
						continue
					else:
						print(f"  -> WARNING: Line not classified by any rule: {repr(line)}")
						continue

				if tag and content_to_match:
					if tag == 'delete' or tag == 'moved_delete':
						granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx + 1, current_b_idx, current_b_idx))
						current_a_idx += 1
					elif tag == 'insert' or tag == 'moved_insert':
						granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx, current_b_idx, current_b_idx + 1))
						current_b_idx += 1
					elif tag == 'equal':
						granular_changes.append((tag, content_to_match, current_a_idx, current_a_idx + 1, current_b_idx, current_b_idx + 1))
						current_a_idx += 1
						current_b_idx += 1

			#print("\n--- End Line-by-Line Parsing Debug ---")
			#print(f"DEBUG: granular_changes after first pass: {granular_changes}")

			# --- Move Detection and Marking (before consolidation) ---
			# Create a dictionary to map content to a list of its occurrences in source/dest
			moved_candidates = {} # content -> [(a_idx, b_idx, 'moved_delete'/'moved_insert', original_granular_idx)]

			for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
				if g_tag in ['moved_delete', 'moved_insert']:
					if g_content not in moved_candidates:
						moved_candidates[g_content] = []
					# Store (a_start, b_start, original_tag, original_index_in_granular_changes)
					# We use start indices (g_a1, g_b1) for simpler matching
					moved_candidates[g_content].append((g_a1, g_b1, g_tag, idx))

			# Keep track of granular indices that are part of a detected 'moved' pair
			# These will be marked with is_moved=True and treated as delete/insert ops
			is_moved_flags = {} # (original_granular_idx) -> True

			for content, candidates in moved_candidates.items():
				deletes = sorted(
					[c for c in candidates if c[2] == "moved_delete"],
					key=lambda c: c[3],
				)
				inserts = sorted(
					[c for c in candidates if c[2] == "moved_insert"],
					key=lambda c: c[3],
				)
				for d_entry, i_entry in zip(deletes, inserts):
					is_moved_flags[d_entry[3]] = True
					is_moved_flags[i_entry[3]] = True

			# --- Consolidation into difflib-style opcodes (with is_moved flag) ---
			final_opcodes_pre_replace = []
			
			current_tag = None
			current_i1, current_i2, current_j1, current_j2 = -1, -1, -1, -1
			current_is_moved_flag = False

			for idx, (g_tag, g_content, g_a1, g_a2, g_b1, g_b2) in enumerate(granular_changes):
				# Determine the actual output tag and its moved status
				# If a 'moved_delete' or 'moved_insert' was paired, its is_moved_flag is True
				# Otherwise, they become regular delete/insert.
				actual_tag = g_tag
				if actual_tag in ['moved_delete', 'moved_insert']:
					actual_tag = 'delete' if g_tag == 'moved_delete' else 'insert'

				is_moved_for_this_item = is_moved_flags.get(idx, False)

				# Initialize for the first valid change or a new type of change
				if current_tag is None:
					current_tag = actual_tag
					current_i1, current_i2 = g_a1, g_a2
					current_j1, current_j2 = g_b1, g_b2
					current_is_moved_flag = is_moved_for_this_item
					continue

				# Check if the current granular change can extend the current block
				can_extend = False
				# A block can only extend if the tags are the same, AND the is_moved_flag is the same
				if actual_tag == current_tag and is_moved_for_this_item == current_is_moved_flag:
					if actual_tag == 'equal':
						if g_a1 == current_i2 and g_b1 == current_j2:
							can_extend = True
					elif actual_tag == 'delete':
						if g_a1 == current_i2:
							can_extend = True
					elif actual_tag == 'insert':
						if g_b1 == current_j2:
							can_extend = True
				
				if can_extend:
					# Corrected lines:
					current_i2 = g_a2
					current_j2 = g_b2
				else:
					# Current block ends, add it to final opcodes
					final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))
					# Start a new block
					current_tag = actual_tag
					current_i1, current_i2 = g_a1, g_a2
					current_j1, current_j2 = g_b1, g_b2
					current_is_moved_flag = is_moved_for_this_item
			
			# Add the last block if any
			if current_tag is not None:
				final_opcodes_pre_replace.append((current_tag, current_i1, current_i2, current_j1, current_j2, current_is_moved_flag))

			# --- Post-consolidation for 'replace' operations ---
			# Now, merge adjacent delete and insert opcodes into 'replace' where appropriate.
			# This must happen after the initial consolidation of same-type operations.
			# Note: We do NOT convert 'moved' deletes/inserts into 'replace' if they are flagged as moved.
			# This is because 'replace' means content changed *in place*, while 'moved' means it changed location.

			consolidated_opcodes = []
			i = 0
			while i < len(final_opcodes_pre_replace):
				current_op = final_opcodes_pre_replace[i]
				tag, i1, i2, j1, j2, is_moved = current_op

				# Check for replace only if the current delete/insert is NOT marked as 'moved'
				if (tag == 'delete' and not is_moved) and i + 1 < len(final_opcodes_pre_replace):
					next_op = final_opcodes_pre_replace[i+1]
					next_tag, next_i1, next_i2, next_j1, next_j2, next_is_moved = next_op

					# If delete is immediately followed by an insert, and neither are marked as moved
					if (next_tag == 'insert' and not next_is_moved) and i2 == next_i1 and j2 == next_j1:
						# Create a 'replace' opcode. The is_moved flag for 'replace' is always False.
						consolidated_opcodes.append(('replace', i1, i2, j1, next_j2, False))
						i += 2 # Skip both current delete and the next insert
						continue
				
				consolidated_opcodes.append(current_op)
				i += 1

			opcodes = consolidated_opcodes

		except Exception as e:
			print(f"An unexpected error occurred during parsing: {e}")
			traceback.print_exc()
			if process:
				print(f"Stdout from git: {process.stdout}")
				print(f"Stderr from git: {process.stderr}")
			return []
		finally:
			self._cleanup_temp_files()
		print("\n\n\n\n****time to run git diff and parse (in sec): ",time.time()-start_time)
		return opcodes



def align_words_with_git_diff(words_data1, words_data2, case_insensitive, ignore_quotes):
	a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
	s = GitSequenceMatcher(a_compare, b_compare, temp_dir=TEMP_PDF_DIR)
	opcodes = s.get_opcodes()
	if not opcodes:
		print("Git diff produced no opcodes; falling back to difflib.")
		return align_words_with_difflib(words_data1, words_data2, case_insensitive, ignore_quotes)
	words_data1, words_data2, _ = apply_opcodes_to_words(words_data1, words_data2, opcodes, with_moves=True)
	return words_data1, words_data2

def is_git_diff_available():
	"""Checks if git diff --no-index is available."""
	git_exe = find_git_executable()
	if git_exe == "git" and not shutil.which("git"):
		return False
	try:
		result = subprocess.run(
			[git_exe, "--version"],
			capture_output=True,
			text=True,
			timeout=10,
		)
		return result.returncode == 0
	except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
		return False


_git_available = is_git_diff_available()
if _git_available:
	print("git diff command available")
else:
	print("git diff command not available; defaulting to difflib")


def get_word_level_opcodes(a_compare, b_compare):
	"""Return word-level diff opcodes and whether move flags are present."""
	if _git_available:
		opcodes = GitSequenceMatcher(a_compare, b_compare, temp_dir=TEMP_PDF_DIR).get_opcodes()
		if opcodes:
			return opcodes, True
	s = difflib.SequenceMatcher(None, a_compare, b_compare)
	return [(*opcode, False) for opcode in s.get_opcodes()], False


def align_words_single_pass(words_data1, words_data2, case_insensitive, ignore_quotes):
	"""Single-pass word alignment using git diff or difflib."""
	a_compare, b_compare = helper_case_quotes(words_data1, words_data2, case_insensitive, ignore_quotes)
	opcodes, with_moves = get_word_level_opcodes(a_compare, b_compare)
	words_data1, words_data2, _ = apply_opcodes_to_words(
		words_data1, words_data2, opcodes, with_moves=with_moves,
	)
	return words_data1, words_data2


def align_words(words_data1, words_data2, case_insensitive, ignore_quotes, use_chunked=True):
	"""
	Align two documents with one full-document word diff.
	use_chunked is accepted for API compatibility; chunk boundaries are
	assigned at extraction and used only when grouping export change records.
	"""
	return align_words_single_pass(words_data1, words_data2, case_insensitive, ignore_quotes)



class PDFViewerPane:
	PAGE_PADDING = 10 
	BUFFER_PAGES = 3  
	def __init__(self, master, parent_app, pane_id):
		self.sorted=None
		self.words_by_unique_id = None
		self.master = master
		self.parent_app = parent_app
		self.pane_id = pane_id
		self.pdf_document = None 
		self.words_data = []	 
		self.zoom_level = 1.0
		self.rendered_page_cache = {} 
		self.page_layout_info = {} 
		self.total_document_height = 0 
		self.max_document_width = 0
		self.last_mouse_x = 0
		self.last_mouse_y = 0
		self.file_name = None
		self.pan_start_x = 0
		self.pan_start_y = 0
		self.canvas_start_x_offset = 0
		self.canvas_start_y_offset = 0
		self.panning = False
		self.render_job_id = None   
		self.resize_job_id = None   
		self.ignore_scroll_events_counter = 0 
		self.temp_pdf_path = None   
		self.loading_message_id = None 
		self.setup_ui()
	def setup_ui(self):
		"""Sets up the UI elements for the PDF viewer pane."""
		self.canvas_frame = ttk.Frame(self.master, relief=tk.SUNKEN, borderwidth=1)
		self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
		self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL)
		self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
		self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL)
		self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
		self.canvas = tk.Canvas(self.canvas_frame, bg="gray",
								yscrollcommand=self.v_scrollbar.set,
								xscrollcommand=self.h_scrollbar.set)
		self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
		self.v_scrollbar.config(command=self.on_vertical_scroll)
		self.h_scrollbar.config(command=self.on_horizontal_scroll)
		self.canvas.bind("<Configure>", self.on_canvas_configure) 
		self.canvas.bind("<MouseWheel>", self.on_mousewheel)	 
		self.canvas.bind("<Button-4>", self.on_mousewheel)	   
		self.canvas.bind("<Button-5>", self.on_mousewheel)	   
		self.canvas.bind("<ButtonPress-1>", self.start_pan)
		self.canvas.bind("<B1-Motion>", self.do_pan)
		self.canvas.bind("<ButtonRelease-1>", self.stop_pan)
		self.canvas.bind('<Up>', self.on_key_scroll)
		self.canvas.bind('<Down>', self.on_key_scroll)
		self.canvas.bind('<Left>', self.on_key_scroll)
		self.canvas.bind('<Right>', self.on_key_scroll)
		self.canvas.bind('<Prior>', self.on_key_scroll) 
		self.canvas.bind('<Next>', self.on_key_scroll)  
		self.canvas.bind('<Home>', self.on_key_scroll) 
		self.canvas.bind('<End>', self.on_key_scroll)  
		self.canvas.bind("<<UserCanvasScrolled>>", lambda event, pane=self: self.parent_app.on_pane_scrolled(event, pane))
		self.canvas.drop_target_register(DND_FILES)
		self.canvas.dnd_bind('<<Drop>>', self.on_drop)
		self.canvas.bind("<Button-3>", self.on_right_click)
		self.context_menu = tk.Menu(self.master, tearoff=0)
		self.canvas.bind("<Double-Button-1>", self._toggle_pan_mode)
		self.canvas.bind("<Motion>", self._on_pan_move)
		self._pan_mode_active = False
		self._cursor_start_pos = None
		self._after_id = None
	def _toggle_pan_mode(self, event):
		"""Toggles the panning mode on or off."""
		if self._pan_mode_active:
			self._deactivate_pan_mode()
		else:
			self._activate_pan_mode()
	def _activate_pan_mode(self):
		"""Activates the clickless pan mode and starts the cursor snap-back timer."""
		if not PYAUTOGUI_AVAILABLE:
			print("Cannot activate pan mode: pyautogui is not installed.")
			return
			
		self._pan_mode_active = True
		self.canvas.config(cursor="hand2")
		self._cursor_start_pos = pyautogui.position()
		
		# Set the initial scan mark
		canvas_x = self.canvas.winfo_pointerx() - self.canvas.winfo_rootx()
		canvas_y = self.canvas.winfo_pointery() - self.canvas.winfo_rooty()
		self.canvas.scan_mark(canvas_x, canvas_y)

		print(f"Pan mode activated. Cursor locked at {self._cursor_start_pos}")
		self._snap_back_timer()
	def _deactivate_pan_mode(self):
		"""Deactivates the clickless pan mode."""
		self._pan_mode_active = False
		self.canvas.config(cursor="")
		if self._after_id:
			self.master.after_cancel(self._after_id)
			self._after_id = None
		print("Pan mode deactivated.")
	def _on_pan_move(self, event):#with timer continuosly postponed
		"""Drags the canvas view, as the mouse moves and without click, if pan mode is active."""
		if self._pan_mode_active:
			#print("event: ",event.x, event.y, "self._cursor_start_pos.x: ",self._cursor_start_pos.x,self.canvas.winfo_rootx())
			#self.canvas.scan_dragto(event.x, event.y, gain=1)
			self.canvas.scan_dragto(self._cursor_start_pos.x- self.canvas.winfo_rootx(), event.y, gain=3)#instead ov event.x we stick to original x (where the user double clicked)
			self.schedule_render_visible_pages() 
			if self.ignore_scroll_events_counter == 0:
				self.canvas.event_generate("<<UserCanvasScrolled>>")
			if self._after_id:
				self.master.after_cancel(self._after_id)
				self._after_id = None
				self._after_id = self.master.after(40, self._snap_back_timer)
	def _snap_back_timer(self):
		"""Periodically snaps the cursor back and resets the scan mark."""
		if not self._pan_mode_active:
			return
		# Move cursor back to the starting point
		pyautogui.moveTo(self._cursor_start_pos.x, self._cursor_start_pos.y)
		# Immediately after moving, we must reset the canvas's scan mark
		# to this position to prevent the canvas from jumping.
		canvas_x = self._cursor_start_pos.x - self.canvas.winfo_rootx()
		canvas_y = self._cursor_start_pos.y - self.canvas.winfo_rooty()
		self.canvas.scan_mark(canvas_x, canvas_y)

		# Schedule the next snap-back
		self._after_id = self.master.after(400, self._snap_back_timer)
	def on_right_click(self, event):
		"""Displays a context menu on right-click."""
		self.context_menu.delete(0, tk.END) 
		if self.pdf_document and not self.pdf_document.is_closed:
			self.context_menu.add_command(
				label="Save PDF with Annotations...",
				command=self.save_pdf_with_annotations
			)
			self.context_menu.add_command(
				label="Toggle light/dark mode",
				command=self.toggle_light_dark_mode
			)
			self.context_menu.add_separator() 
		self.context_menu.add_command(
			label="Paste from Clipboard",
			command=self.paste_from_clipboard_action
		)
		try:
			self.context_menu.tk_popup(event.x_root, event.y_root)
		finally:
			self.context_menu.grab_release()
	def toggle_light_dark_mode(self):
		"""Toggles the blend mode of PDFComparer highlights between Multiply and Exclusion."""
		if not self.pdf_document or self.pdf_document.is_closed:
			return

		current_mode = None
		# First, determine the current mode by checking the first relevant annotation
		for page in self.pdf_document:
			if current_mode:
				break
			for annot in page.annots():
				if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get("title") == "PDFComparer":
					current_mode = annot.blendmode
					break
		
		if current_mode is None:
			# No relevant annotations found, nothing to do.
			return

		# Determine the target mode
		new_mode = "Exclusion" if current_mode == "Multiply" else "Multiply"

		# Now, update all relevant annotations
		for page in self.pdf_document:
			for annot in page.annots():
				if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get("title") == "PDFComparer":
					annot.set_blendmode(new_mode)
					if new_mode=="Exclusion":# changing to dark mode
						annot.set_opacity(1)
					else:#changing to light mode
						annot.set_opacity(0.3)
					annot.update()
		
		self._clear_all_rendered_pages() 
		self.calculate_document_layout() 
		self.render_visible_pages() 
		self.canvas.focus_set() 
		
		
		
		
		print(f"Toggled to {new_mode} mode.")
	def paste_from_clipboard_action(self):
		"""
		Handles the "Paste from Clipboard" action.
		Converts clipboard content to a temporary PDF and loads it into the current pane.
		"""
		temp_output_filename = os.path.join(TEMP_PDF_DIR, f"clipboard_temp_{os.urandom(8).hex()}.pdf")
		self.display_loading_message("Pasting from clipboard...")
		load_thread = threading.Thread(target=self._paste_from_clipboard_threaded,
									   args=(temp_output_filename,))
		load_thread.daemon = True
		load_thread.start()
	def _paste_from_clipboard_threaded(self, temp_output_filename):
		"""
		Performs the clipboard to PDF conversion in a background thread.
		Schedules the GUI update after conversion.
		"""
		converted_file_path = convert_clipboard_to_pdf(temp_output_filename)
		self.master.after(1, self._on_paste_from_clipboard_complete_gui_update,
						  converted_file_path, temp_output_filename)
	def _on_paste_from_clipboard_complete_gui_update(self, converted_file_path, original_temp_filename):
		"""
		Updates the UI after clipboard content has been converted to PDF.
		Runs on the main Tkinter thread.
		"""
		self.hide_loading_message()
		if converted_file_path:
			self.parent_app._initiate_load_process(converted_file_path,
												  0 if self.pane_id == 'left' else 1,
												  "Clipboard Content")
			print("converted_file_path: ",converted_file_path)
			self.temp_pdf_path = converted_file_path 
		else:
			messagebox.showerror("Clipboard Error", "Could not convert clipboard content to PDF. It might be empty or contain unsupported content.")
			self._clear_all_rendered_pages()
	def save_pdf_with_annotations(self):
		"""Saves the current PDF document, including annotations, to a new file."""
		if not self.pdf_document or self.pdf_document.is_closed:
			messagebox.showinfo("Save PDF", "No PDF document is currently open in this pane to save.")
			return
		initial_file = self.file_name if self.file_name else "document"
		base_name, ext = os.path.splitext(initial_file)
		if ext.lower() != ".pdf": 
			initial_file = base_name + ".pdf"
		if "_diff" not in base_name.lower():
			initial_file = f"{base_name}_diff{ext}"
		file_path = filedialog.asksaveasfilename(
			defaultextension=".pdf",
			filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
			initialfile=initial_file
		)
		if file_path:
			try:
				self.pdf_document.save(file_path)
				messagebox.showinfo("Save PDF", f"PDF saved successfully to:\n{file_path}")
			except Exception as e:
				messagebox.showerror("Save PDF Error", f"Failed to save PDF: {e}")
	def on_drop(self, event):
		"""Handler for drag-and-drop file events."""
		file_path = event.data
		if file_path.startswith('{') and file_path.endswith('}'):
			file_path = file_path[1:-1]
		self.parent_app.open_pdf_from_drop(file_path, self.pane_id)
	def display_loading_message(self, message="Loading..."):
		"""Displays a loading message on the canvas."""
		self.hide_loading_message() 
		self.canvas.delete("all") 
		canvas_center_x = self.canvas.winfo_width() / 2
		canvas_center_y = self.canvas.winfo_height() / 2
		self.loading_message_id = self.canvas.create_text(
			canvas_center_x, canvas_center_y,
			text=message, fill="black", font=("Helvetica", 24, "bold"),
			tags="loading_message"
		)
		self.canvas.config(scrollregion=(0,0,0,0)) 
	def hide_loading_message(self):
		"""Hides the loading message from the canvas."""
		if self.loading_message_id:
			self.canvas.delete(self.loading_message_id)
			self.loading_message_id = None
	def load_pdf_internal(self, file_path):
		"""
		Internal method to load a PDF (or converted document) and extract words.
		This runs in a worker thread. It returns document and words data, or None on failure.
		"""
		temp_pdf_path_used = None
		pdf_document_obj = None
		words_data_obj = []
		try:
			original_file_extension = os.path.splitext(file_path)[1].lower()
			if original_file_extension in ['.doc', '.docx', '.rtf', '.txt']:
				print(f"Attempting to convert {original_file_extension} file to PDF...")
				converted_pdf_path = convert_word_to_pdf_no_markup(file_path)
				if converted_pdf_path:
					file_path = converted_pdf_path
					temp_pdf_path_used = converted_pdf_path
					print(f"Successfully converted to temporary PDF: {converted_pdf_path}")
				else:
					return None, [], None, "Conversion Failed"
			pdf_document_obj = fitz.open(file_path)
			ignore_ligatures = self.parent_app.ignore_ligatures.get()
			words_data_obj = extract_words_with_styles(
				pdf_document_obj,
				ignore_ligatures=ignore_ligatures,
				split_punctuation=self.parent_app.split_punctuation.get(),
				merge_hyphenation=self.parent_app.merge_hyphenation.get(),
			)
			if file_path.find("clipboard_temp_")!=-1: temp_pdf_path_used=file_path
			return pdf_document_obj, words_data_obj, temp_pdf_path_used, None 
		except Exception as e:
			print(f"Error in load_pdf_internal: {e}")
			if temp_pdf_path_used and os.path.exists(temp_pdf_path_used):
				try:
					os.remove(temp_pdf_path_used)
					print(f"Cleaned up temporary PDF on error: {temp_pdf_path_used}")
				except Exception as cleanup_e:
					print(f"Error cleaning up temp PDF: {cleanup_e}")
			if pdf_document_obj:
				pdf_document_obj.close()
			return None, [], None, f"Could not open PDF: {e}" 
	def get_current_view_coords(self):
		"""Returns the content coordinates of the top-left corner of the canvas viewport."""
		return self.canvas.canvasx(0), self.canvas.canvasy(0)
	def get_current_view_height_in_content_coords(self):
		"""Returns the height of the current view in content coordinates."""
		return self.canvas.winfo_height()
	def calculate_document_layout(self):
		"""Calculates the layout (dimensions and positions) of all pages based on the current zoom level."""
		self.page_layout_info.clear()
		y_offset = 0
		max_width = 0
		if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
			self.total_document_height = 0
			self.max_document_width = 0
			self.canvas.config(scrollregion=(0,0,0,0)) 
			return
		for i in range(self.pdf_document.page_count):
			page = self.pdf_document.load_page(i)
			base_width = int(page.mediabox.width)
			base_height = int(page.mediabox.height)
			self.page_layout_info[i] = {
				"base_width": base_width,
				"base_height": base_height,
				"y_start_offset": y_offset
			}
			y_offset += base_height + self.PAGE_PADDING 
			max_width = max(max_width, base_width) 
		self.total_document_height = y_offset
		self.max_document_width = max_width
		self.canvas.config(scrollregion=(0, 0,
										 self.max_document_width * self.zoom_level,
										 self.total_document_height * self.zoom_level))
	def schedule_render_visible_pages(self, event=None):
		"""Debounces rendering of visible pages to prevent excessive redraws."""
		if self.render_job_id:
			self.master.after_cancel(self.render_job_id) 
		if self.pdf_document and not self.pdf_document.is_closed:
			self.render_job_id = self.master.after(50, self.render_visible_pages) 
	def schedule_fit_to_width(self, event=None):
		"""Debounces the fit_to_width operation, typically on canvas resize."""
		if self.resize_job_id:
			self.master.after_cancel(self.resize_job_id)
		self.resize_job_id = self.master.after(150, self.fit_to_width) 
	def _clear_all_rendered_pages(self):
		"""Clears all rendered pages from the canvas and cache."""
		self.canvas.delete("all")
		self.rendered_page_cache.clear()
		self.hide_loading_message() 
	def render_visible_pages(self):
		"""Renders pages that are currently visible within the canvas viewport, including a buffer."""
		if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
			return
		if self.canvas.winfo_width() == 0 or self.canvas.winfo_height() == 0:
			return 
		current_view_x_content_coord = self.canvas.canvasx(0)
		current_view_y_content_coord = self.canvas.canvasy(0)
		canvas_width = self.canvas.winfo_width()
		canvas_height = self.canvas.winfo_height()
		visible_y_start = current_view_y_content_coord
		visible_y_end = current_view_y_content_coord + canvas_height
		pages_to_render_now = set()
		for page_num in range(self.pdf_document.page_count):
			page_info = self.page_layout_info.get(page_num)
			if not page_info: continue
			scaled_y_start = page_info["y_start_offset"] * self.zoom_level
			scaled_height = page_info["base_height"] * self.zoom_level
			scaled_height_with_padding = scaled_height + self.PAGE_PADDING * self.zoom_level
			buffer_height_px = self.BUFFER_PAGES * scaled_height_with_padding
			page_top_buffered = scaled_y_start - buffer_height_px
			page_bottom_buffered = scaled_y_start + scaled_height_with_padding + buffer_height_px
			if (page_bottom_buffered >= visible_y_start and
				page_top_buffered <= visible_y_end):
				pages_to_render_now.add(page_num)
		pages_currently_cached = set(self.rendered_page_cache.keys())
		pages_to_remove = pages_currently_cached - pages_to_render_now
		for page_num in pages_to_remove:
			if page_num in self.rendered_page_cache:
				data = self.rendered_page_cache[page_num]
				self.canvas.delete(data["canvas_id"]) 
				del self.rendered_page_cache[page_num] 
		for page_num in pages_to_render_now:
			if page_num not in self.rendered_page_cache:
				try:
					if self.pdf_document.is_closed: 
						continue
					page = self.pdf_document.load_page(page_num)
					matrix = fitz.Matrix(self.zoom_level, self.zoom_level) 
					pix = page.get_pixmap(matrix=matrix)
					img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
					tk_img = ImageTk.PhotoImage(img)
					content_width_at_zoom = self.max_document_width * self.zoom_level
					page_width_at_zoom = page.rect.width * self.zoom_level
					page_x_offset_on_canvas = (content_width_at_zoom - page_width_at_zoom) / 2
					y_pos_on_canvas = self.page_layout_info[page_num]["y_start_offset"] * self.zoom_level
					canvas_id = self.canvas.create_image(page_x_offset_on_canvas, y_pos_on_canvas, anchor=tk.NW, image=tk_img)
					self.rendered_page_cache[page_num] = {"image": tk_img, "canvas_id": canvas_id}
				except Exception as e:
					print(f"Error rendering page {page_num} in pane {self.pane_id}: {e}")
					break
	def fit_to_width(self):
		"""
		Calculates and sets the zoom level so that the first page of the PDF
		fits the width of the canvas. This is usually called on initial load or resize.
		"""
		if not self.pdf_document or self.pdf_document.is_closed or self.pdf_document.page_count == 0:
			return
		if self.canvas.winfo_width() == 0:
			self.master.after(100, self.fit_to_width) 
			return
		page_rect = self.pdf_document.load_page(0).mediabox 
		page_width = page_rect.width
		canvas_width = self.canvas.winfo_width() - self.v_scrollbar.winfo_width()
		if canvas_width <= 0:
			canvas_width = 1 
		new_zoom = canvas_width / page_width
		self.set_zoom(new_zoom) 
	def set_zoom(self, new_zoom_level, mouse_x_canvas_pixel=None, mouse_y_canvas_pixel=None, from_sync=False):
		"""
		Adjusts the zoom level of the PDF viewer.
		Can optionally zoom around a specific mouse coordinate.
		"""
		old_zoom = self.zoom_level
		if abs(new_zoom_level - old_zoom) < 0.001:
			return
		if not self.pdf_document or self.pdf_document.is_closed:
			return
		if self.canvas.winfo_width() == 0 or self.canvas.winfo_height() == 0:
			self.master.after(100, lambda: self.set_zoom(new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel, from_sync))
			return
		if mouse_x_canvas_pixel is None:
			mouse_x_canvas_pixel = int(self.canvas.winfo_width() / 2)
		if mouse_y_canvas_pixel is None:
			mouse_y_canvas_pixel = int(self.canvas.winfo_height() / 2)
		mouse_x_content_coord_old_zoom = self.canvas.canvasx(mouse_x_canvas_pixel)
		mouse_y_content_coord_old_zoom = self.canvas.canvasy(mouse_y_canvas_pixel)
		mouse_x_doc_coord = mouse_x_content_coord_old_zoom / old_zoom if old_zoom != 0 else 0
		mouse_y_doc_coord = mouse_y_content_coord_old_zoom / old_zoom if old_zoom != 0 else 0
		self._clear_all_rendered_pages() 
		self.zoom_level = new_zoom_level 
		self.calculate_document_layout() 
		new_mouse_x_content_coord = mouse_x_doc_coord * self.zoom_level
		new_mouse_y_content_coord = mouse_y_doc_coord * self.zoom_level
		new_x_scroll_pixels = new_mouse_x_content_coord - mouse_x_canvas_pixel
		new_y_scroll_pixels = new_mouse_y_content_coord - mouse_y_canvas_pixel
		self._apply_scroll(new_x_scroll_pixels, new_y_scroll_pixels) 
		self.render_visible_pages() 
		self.canvas.focus_set() 
		if not from_sync and self.parent_app.sync_zoom_enabled.get():
			self.parent_app.sync_zoom(self, new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel)
		self.parent_app.update_zoom_label(self.pane_id, self.zoom_level) 
	def set_zoom_from_scale_widget(self, val):
		"""Callback for the zoom scale widget."""
		self.set_zoom(float(val), from_sync=False)
	def on_vertical_scroll(self, *args):
		"""Handles vertical scrollbar movements and calls for rendering."""
		self.canvas.yview(*args)
		self.schedule_render_visible_pages() 
		if self.ignore_scroll_events_counter == 0:
			self.canvas.event_generate("<<UserCanvasScrolled>>")
	def on_horizontal_scroll(self, *args):
		"""Handles horizontal scrollbar movements and calls for rendering."""
		self.canvas.xview(*args)
		self.schedule_render_visible_pages() 
		if self.ignore_scroll_events_counter == 0:
			self.canvas.event_generate("<<UserCanvasScrolled>>")
	def on_mousewheel(self, event):
		"""Handles mouse wheel scrolling for both vertical scroll and zoom (with Ctrl key)."""
		if not self.pdf_document or self.pdf_document.is_closed:
			return "break" 
		scroll_delta = 0
		if event.delta: 
			scroll_delta = -int(event.delta/120) 
		elif event.num == 4: 
			scroll_delta = -1
		elif event.num == 5: 
			scroll_delta = 1
		if event.state & 0x4: 
			old_zoom = self.zoom_level
			zoom_factor = 1.1 if scroll_delta < 0 else (1/1.1) 
			new_zoom_level = self.zoom_level * zoom_factor
			min_zoom = 0.25
			max_zoom = 4.0
			new_zoom_level = max(min_zoom, min(max_zoom, new_zoom_level))
			if abs(new_zoom_level - old_zoom) > 0.001:
				self.set_zoom(new_zoom_level, event.x, event.y, from_sync=False)
		elif event.state == 9 or (event.state & 0x1): 
			self.canvas.xview_scroll(scroll_delta, "units") 
			if self.ignore_scroll_events_counter == 0:
				self.canvas.event_generate("<<UserCanvasScrolled>>")
		else:
			self.canvas.yview_scroll(scroll_delta, "units") 
			if self.ignore_scroll_events_counter == 0:
				self.canvas.event_generate("<<UserCanvasScrolled>>")
		self.schedule_render_visible_pages()
		return "break" 
	def on_key_scroll(self, event):
		"""Handles keyboard-initiated scrolling."""
		if not self.pdf_document or self.pdf_document.is_closed:
			return "break"
		scroll_amount_units = 3
		scroll_amount_pages = 1
		if event.keysym == 'Up':
			self.canvas.yview_scroll(-scroll_amount_units, "units")
		elif event.keysym == 'Down':
			self.canvas.yview_scroll(scroll_amount_units, "units")
		elif event.keysym == 'Left':
			self.canvas.xview_scroll(-scroll_amount_units, "units")
		elif event.keysym == 'Right':
			self.canvas.xview_scroll(scroll_amount_units, "units")
		elif event.keysym == 'Prior': 
			self.canvas.yview_scroll(-scroll_amount_pages, "pages")
		elif event.keysym == 'Next':  
			self.canvas.yview_scroll(scroll_amount_pages, "pages")
		elif event.keysym == 'Home': 
			self.canvas.yview_moveto(0.0)
		elif event.keysym == 'End':  
			self.canvas.yview_moveto(1.0)
		else:
			return 
		if self.ignore_scroll_events_counter == 0:
			self.canvas.event_generate("<<UserCanvasScrolled>>")
		self.schedule_render_visible_pages()
		return "break"
	def _apply_scroll(self, x_scroll_pixels, y_scroll_pixels):
		"""
		Applies a scroll to the canvas programmatically.
		Increments ignore_scroll_events_counter to prevent sync-scroll loops.
		"""
		self.ignore_scroll_events_counter += 1
		try:
			if not self.pdf_document or self.pdf_document.is_closed:
				return
			total_width_at_zoom = self.max_document_width * self.zoom_level
			total_height_at_zoom = self.total_document_height * self.zoom_level
			max_x_scroll = max(0, total_width_at_zoom - self.canvas.winfo_width())
			max_y_scroll = max(0, total_height_at_zoom - self.canvas.winfo_height())
			x_scroll_pixels = max(0, min(x_scroll_pixels, max_x_scroll))
			y_scroll_pixels = max(0, min(y_scroll_pixels, max_y_scroll))
			x_prop = x_scroll_pixels / total_width_at_zoom if total_width_at_zoom > 0 else 0
			y_prop = y_scroll_pixels / total_height_at_zoom if total_height_at_zoom > 0 else 0
			self.canvas.xview_moveto(x_prop)
			self.canvas.yview_moveto(y_prop)
			self.schedule_render_visible_pages()
		finally:
			self.ignore_scroll_events_counter -= 1 
	def start_pan(self, event):
		"""Initiates panning functionality by storing initial mouse and canvas positions."""
		if not self.pdf_document or self.pdf_document.is_closed:
			return
		self.panning = True
		self.pan_start_x = event.x
		self.pan_start_y = event.y
		self.canvas_start_x_offset = self.canvas.canvasx(0)
		self.canvas_start_y_offset = self.canvas.canvasy(0)
		self.canvas.config(cursor="fleur") 
		self.canvas.focus_set() 
	def do_pan(self, event):
		"""Performs panning movement based on initial and current mouse positions."""
		if self.panning and self.pdf_document and not self.pdf_document.is_closed:
			dx = event.x - self.pan_start_x
			dy = event.y - self.pan_start_y
			target_x_scroll = self.canvas_start_x_offset - dx
			target_y_scroll = self.canvas_start_y_offset - dy
			self.ignore_scroll_events_counter += 1 
			try:
				total_width_at_zoom = self.max_document_width * self.zoom_level
				total_height_at_zoom = self.total_document_height * self.zoom_level
				x_prop = target_x_scroll / total_width_at_zoom if total_width_at_zoom > 0 else 0
				y_prop = target_y_scroll / total_height_at_zoom if total_height_at_zoom > 0 else 0
				x_prop = max(0.0, min(x_prop, 1.0))
				y_prop = max(0.0, min(y_prop, 1.0))
				self.canvas.xview_moveto(x_prop)
				self.canvas.yview_moveto(y_prop)
				#print("%.2f yview_moveto: %s"%(time.time(), y_prop))
			finally:
				self.ignore_scroll_events_counter -= 1
			self.schedule_render_visible_pages() 
			if self.ignore_scroll_events_counter == 0:
				self.canvas.event_generate("<<UserCanvasScrolled>>")
	def stop_pan(self, event):
		"""Stops panning functionality."""
		if not self.pdf_document or self.pdf_document.is_closed:
			return
		self.panning = False
		self.canvas.config(cursor="") 
		self.schedule_render_visible_pages()
		self.canvas.focus_set()
	def on_canvas_configure(self, event):
		"""Handles canvas resizing and reconfigures the scroll region and rendering."""
		if self.pdf_document and not self.pdf_document.is_closed:
			current_x_prop = self.canvas.xview()[0]
			current_y_prop = self.canvas.yview()[0]
			self.calculate_document_layout() 
			self.schedule_fit_to_width()
			total_width_at_zoom = self.max_document_width * self.zoom_level
			total_height_at_zoom = self.total_document_height * self.zoom_level
			self._apply_scroll(current_x_prop * total_width_at_zoom,
							   current_y_prop * total_height_at_zoom)
			self.render_visible_pages() 
		self.canvas.focus_set()
	def close_pdf(self):
		"""Closes the PDF document and clears associated data and canvas, including temporary file."""
		if self.pdf_document:
			try:
				if 0: 
					for page_num in range(self.pdf_document.page_count):
						page = self.pdf_document.load_page(page_num)
						annots_to_delete = [
							annot for annot in page.annots()
							if annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT and annot.info.get("title") == "PDFComparer"
						]
						for annot in annots_to_delete:
							try:
								page.delete_annot(annot)
							except Exception as e:
								print(f"Error deleting annotation during close on page {page_num}: {e}")
				self.pdf_document.close()
				print(f"Pane {self.pane_id}: PDF document closed successfully.")
			except Exception as e:
				print(f"Pane {self.pane_id}: Error closing PDF document: {e}")
			self.pdf_document = None 
		self.file_name = None 
		self.sorted = None
		self.words_by_unique_id = None
		self.rendered_page_cache.clear()
		self.words_data = [] 
		self.page_layout_info = {} 
		self.canvas.delete("all") 
		self.canvas.config(scrollregion=(0,0,0,0))
		if self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
			try:
				os.remove(self.temp_pdf_path)
				print(f"Pane {self.pane_id}: Deleted temporary PDF: {self.temp_pdf_path}")
			except Exception as e:
				print(f"Pane {self.pane_id}: Error deleting temporary PDF {self.temp_pdf_path}: {e}")
			self.temp_pdf_path = None 
		self.hide_loading_message() 
		print(f"Pane {self.pane_id}: PDF closed and resources cleared.")
class PDFViewerApp:
	def __init__(self, master):
		self.master = master
		self.master.geometry("1200x800") 
		self.pdf_documents = [None, None] 
		self.words_data_list = [None, None] 
		self.current_active_pane = None 
		self.pane1 = None 
		self.pane2 = None 
		self.zoom_scale_1 = None
		self.zoom_scale_2 = None
		self.zoom_percent_label_1 = None
		self.zoom_percent_label_2 = None
		self.sync_scroll_enabled = tk.BooleanVar(value=True)
		self.sync_zoom_enabled = tk.BooleanVar(value=True)
		self.case_insensitive = tk.BooleanVar(value=True)
		self.ignore_quotes = tk.BooleanVar(value=True)
		self.ignore_ligatures = tk.BooleanVar(value=True)
		self.split_punctuation = tk.BooleanVar(value=True)
		self.merge_hyphenation = tk.BooleanVar(value=True)
		self.chunk_compare = tk.BooleanVar(value=True)
		self.setup_ui() 
		self.update_window_title() 
		self.master.after_idle(self.update_ui_state)
		self.master.after_idle(lambda: self.pane1.canvas.focus_set())
		self._process_command_line_args()
		#variable used in the sync_scroll()
		self.scroll_time=0
		self.scroll_y=0
		self.scroll_pane=None
		self.scroll_height=0
		self.scroll_target_y=0
		self.scroll_distance=0
		self._compare_in_progress = False
		self._compare_thread = None
		self.words_data_original = [None, None]
		self.change_records = None
		self.change_records_path = None

	def setup_ui(self):
		"""Sets up the main application UI, including control frame and viewer panes."""
		control_frame = ttk.Frame(self.master, padding="10")
		control_frame.pack(fill=tk.X, side=tk.TOP)
		self.open_button_1 = ttk.Button(control_frame, text="Open (L)", command=lambda: self.open_pdf(0))
		self.open_button_1.pack(side=tk.LEFT, padx=5)
		self.open_button_2 = ttk.Button(control_frame, text="Open (R)", command=lambda: self.open_pdf(1))
		self.open_button_2.pack(side=tk.LEFT, padx=5)
		ttk.Label(control_frame, text="Zoom (L):").pack(side=tk.LEFT, padx=(15, 0))
		self.zoom_scale_1 = ttk.Scale(control_frame, from_=0.33, to_=3.0, orient=tk.HORIZONTAL, length=100)
		self.zoom_scale_1.set(1.0)
		self.zoom_scale_1.pack(side=tk.LEFT, padx=5)
		self.zoom_percent_label_1 = ttk.Label(control_frame, text="100%")
		self.zoom_percent_label_1.pack(side=tk.LEFT)
		ttk.Label(control_frame, text="Zoom (R):").pack(side=tk.LEFT, padx=(15, 0))
		self.zoom_scale_2 = ttk.Scale(control_frame, from_=0.33, to_=3.0, orient=tk.HORIZONTAL, length=100)
		self.zoom_scale_2.set(1.0)
		self.zoom_scale_2.pack(side=tk.LEFT, padx=5)
		self.zoom_percent_label_2 = ttk.Label(control_frame, text="100%")
		self.zoom_percent_label_2.pack(side=tk.LEFT)
		self.prev_change_button = ttk.Button(control_frame, text="Prev.", command=self.go_to_prev_change, underline=0)
		self.prev_change_button.pack(side=tk.LEFT, padx=(20, 5))
		self.next_change_button = ttk.Button(control_frame, text="Next", command=self.go_to_next_change, underline=0)
		self.next_change_button.pack(side=tk.LEFT, padx=5)
		self.recompare_button = ttk.Button(control_frame, text="Re-compare", command=self.recompare_documents)
		self.recompare_button.pack(side=tk.LEFT, padx=5)
		self.export_changes_button = ttk.Button(
			control_frame, text="Export changes", command=self.export_change_records_dialog,
		)
		self.export_changes_button.pack(side=tk.LEFT, padx=5)
		ToolTip(
			self.export_changes_button,
			"Save a copy of the change JSON. A file is auto-created in temp_pdfs/ after each compare.",
		)
		self.sync_scroll_checkbox = ttk.Checkbutton(control_frame, text="Sync Scroll",
													variable=self.sync_scroll_enabled, onvalue=True, offvalue=False)
		self.sync_scroll_checkbox.pack(side=tk.LEFT, padx=(20, 5))
		self.sync_zoom_checkbox = ttk.Checkbutton(control_frame, text="Sync Zoom",
												  variable=self.sync_zoom_enabled, onvalue=True, offvalue=False)
		self.sync_zoom_checkbox.pack(side=tk.LEFT, padx=5)
		self.case_insensitive_checkbox = ttk.Checkbutton(control_frame, text="Case Insensitive",
												  variable=self.case_insensitive, onvalue=True, offvalue=False)
		self.case_insensitive_checkbox.pack(side=tk.LEFT, padx=5)
		self.tip_case_insensitive = ToolTip(
			self.case_insensitive_checkbox,
			"Applied when comparing. Use Re-compare after changing.",
		)
		self.ignore_quotes_checkbox = ttk.Checkbutton(control_frame, text="Ignore quotes type",
												  variable=self.ignore_quotes, onvalue=True, offvalue=False)
		self.ignore_quotes_checkbox.pack(side=tk.LEFT, padx=5)
		self.tip_ignore_quotes = ToolTip(
			self.ignore_quotes_checkbox,
			"Applied when comparing. Use Re-compare after changing.",
		)
		self.ignore_ligatures_checkbox = ttk.Checkbutton(control_frame, text="Ignore 'f' ligatures",
												  variable=self.ignore_ligatures, onvalue=True, offvalue=False)
		self.ignore_ligatures_checkbox.pack(side=tk.LEFT, padx=5)
		self.tip_ignore_ligatures = ToolTip(
			self.ignore_ligatures_checkbox,
			"Applied when extracting text. Reload files after changing.",
		)
		self.split_punctuation_checkbox = ttk.Checkbutton(
			control_frame, text="Split punct.",
			variable=self.split_punctuation, onvalue=True, offvalue=False,
		)
		self.split_punctuation_checkbox.pack(side=tk.LEFT, padx=5)
		ToolTip(self.split_punctuation_checkbox, "Split attached punctuation when extracting. Reload files after changing.")
		self.merge_hyphenation_checkbox = ttk.Checkbutton(
			control_frame, text="Merge hyphens",
			variable=self.merge_hyphenation, onvalue=True, offvalue=False,
		)
		self.merge_hyphenation_checkbox.pack(side=tk.LEFT, padx=5)
		ToolTip(self.merge_hyphenation_checkbox, "Merge line-break hyphenations when extracting. Reload files after changing.")
		self.chunk_compare_checkbox = ttk.Checkbutton(
			control_frame, text="Group by chunk",
			variable=self.chunk_compare, onvalue=True, offvalue=False,
		)
		self.chunk_compare_checkbox.pack(side=tk.LEFT, padx=5)
		ToolTip(
			self.chunk_compare_checkbox,
			"Group exported change records by paragraph chunk. "
			"Does not affect compare speed. Use Re-compare after changing.",
		)
		self.panes_container = ttk.Frame(self.master)
		self.panes_container.pack(fill=tk.BOTH, expand=True)
		self.pane1 = PDFViewerPane(self.panes_container, self, 'left')
		self.pane1.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
		self.pane1.canvas.bind("<FocusIn>", lambda e: self.set_active_pane(self.pane1))
		self.pane2 = PDFViewerPane(self.panes_container, self, 'right')
		self.pane2.canvas_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
		self.pane2.canvas.bind("<FocusIn>", lambda e: self.set_active_pane(self.pane2))
		self.zoom_scale_1.config(command=self.pane1.set_zoom_from_scale_widget)
		self.zoom_scale_2.config(command=self.pane2.set_zoom_from_scale_widget)
		self.master.bind('p', lambda event: self.go_to_prev_change())
		self.master.bind('P', lambda event: self.go_to_prev_change())
		self.master.bind('n', lambda event: self.go_to_next_change())
		self.master.bind('N', lambda event: self.go_to_next_change())
	def _process_command_line_args(self):
		"""Processes command-line arguments to load initial PDF files."""
		if len(sys.argv) > 1:
			file_path1 = sys.argv[1]
			self.master.after_idle(lambda: self._initiate_load_process(file_path1, 0, os.path.basename(file_path1)))
		if len(sys.argv) > 2:
			file_path2 = sys.argv[2]
			self.master.after_idle(lambda: self._initiate_load_process(file_path2, 1, os.path.basename(file_path2)))
	def update_window_title(self):
		"""Updates the main window's title to show the names of the loaded PDF files."""
		name1 = self.pane1.file_name if self.pane1.file_name else "Panel 1"
		name2 = self.pane2.file_name if self.pane2.file_name else "Panel 2"
		self.master.title(f"PDF Diff Viewer - {name1} vs {name2}")
	def set_active_pane(self, pane):
		"""Sets the currently active pane (the one with keyboard focus)."""
		self.current_active_pane = pane
	def update_zoom_label(self, pane_id, zoom_level):
		"""Updates the zoom percentage label and slider for a given pane."""
		if pane_id == 'left' and self.zoom_scale_1 and self.zoom_percent_label_1:
			self.zoom_percent_label_1.config(text=f"{int(zoom_level * 100)}%")
			if abs(self.zoom_scale_1.get() - zoom_level) > 0.001:
				self.zoom_scale_1.config(command=lambda *args: None) 
				self.zoom_scale_1.set(zoom_level)
				self.zoom_scale_1.config(command=self.pane1.set_zoom_from_scale_widget) 
		elif pane_id == 'right' and self.zoom_scale_2 and self.zoom_percent_label_2:
			self.zoom_percent_label_2.config(text=f"{int(zoom_level * 100)}%")
			if abs(self.zoom_scale_2.get() - zoom_level) > 0.001:
				self.zoom_scale_2.config(command=lambda *args: None)
				self.zoom_scale_2.set(zoom_level)
				self.zoom_scale_2.config(command=self.pane2.set_zoom_from_scale_widget)
	def update_ui_state(self):
		"""Updates the state of UI elements (e.g., enable/disable zoom sliders)."""
		doc1_loaded = self.pdf_documents[0] and not self.pdf_documents[0].is_closed if self.pdf_documents[0] else False
		doc2_loaded = self.pdf_documents[1] and not self.pdf_documents[1].is_closed if self.pdf_documents[1] else False
		self.zoom_scale_1.config(state=tk.NORMAL if doc1_loaded else tk.DISABLED)
		self.zoom_scale_2.config(state=tk.NORMAL if doc2_loaded else tk.DISABLED)
		self.prev_change_button.config(state=tk.NORMAL if (doc1_loaded and doc2_loaded) else tk.DISABLED)
		self.next_change_button.config(state=tk.NORMAL if (doc1_loaded and doc2_loaded) else tk.DISABLED)
		self.recompare_button.config(state=tk.NORMAL if (doc1_loaded and doc2_loaded and not self._compare_in_progress) else tk.DISABLED)
		self.export_changes_button.config(
			state=tk.NORMAL if (doc1_loaded and doc2_loaded and self.change_records is not None) else tk.DISABLED
		)
		self.update_zoom_label('left', self.pane1.zoom_level if doc1_loaded else 1.0)
		self.update_zoom_label('right', self.pane2.zoom_level if doc2_loaded else 1.0)
	def open_pdf(self, pane_index):
		"""
		Opens a PDF file or converts a document to PDF and then opens it in the specified pane.
		This is triggered by the "Open PDF/Doc" buttons.
		"""
		file_types = [
			("All supported files", "*.pdf .docx *.doc *.rtf *.txt"),
			("PDF files", "*.pdf"),
			("Word Documents", "*.docx *.doc"),
			("Rich Text Format", "*.rtf"),
			("Text files", "*.txt"),
			("All files", "*.*")
		]
		file_path = filedialog.askopenfilename(filetypes=file_types)
		if file_path:
			self._initiate_load_process(file_path, pane_index, os.path.basename(file_path))
	def open_pdf_from_drop(self, file_path, pane_id):
		"""
		Opens a PDF from a drag-and-drop event in the specified pane.
		"""
		pane_index = 0 if pane_id == 'left' else 1
		self._initiate_load_process(file_path, pane_index, os.path.basename(file_path))
	def _initiate_load_process(self, file_path, pane_index, display_file_name):
		"""
		Initiates the PDF loading and processing in a separate thread.
		This method is called by both open_pdf and open_pdf_from_drop.
		"""
		pane = self.pane1 if pane_index == 0 else self.pane2
		pane.display_loading_message(f"Loading '{display_file_name}'...")
		self.pdf_documents[pane_index] = None
		self.words_data_list[pane_index] = None
		pane.words_data = [] 
		pane.close_pdf() 
		pane._clear_all_rendered_pages() 
		load_thread = threading.Thread(target=self._load_and_process_pdf_threaded,
									   args=(file_path, pane_index, display_file_name))
		load_thread.daemon = True 
		load_thread.start()
	def _load_and_process_pdf_threaded(self, file_path, pane_index, display_file_name):
		"""
		This method runs in a separate thread. It performs file conversion,
		PDF opening, and word extraction. It then schedules the GUI update
		back on the main thread.
		"""
		pane = self.pane1 if pane_index == 0 else self.pane2
		pdf_doc, words_data, temp_path, error_message = pane.load_pdf_internal(file_path)
		self.master.after(1, self._on_pdf_load_complete_gui_update,
						  pane_index, pdf_doc, words_data, temp_path, error_message, display_file_name)
	def _on_pdf_load_complete_gui_update(self, pane_index, pdf_doc, words_data, temp_path, error_message, display_file_name):
		"""
		This method runs on the main Tkinter thread. It updates the UI
		after a PDF has been loaded and processed in a background thread.
		"""
		pane = self.pane1 if pane_index == 0 else self.pane2
		pane.hide_loading_message() 
		if error_message:
			messagebox.showerror("Error", f"Failed to open/process file in {pane.pane_id} pane: {error_message}")
			self.pdf_documents[pane_index] = None
			self.words_data_list[pane_index] = None
			pane.temp_pdf_path = None
			pane.canvas.config(scrollregion=(0,0,0,0)) 
			pane.canvas.delete("all")
			self.update_ui_state()
			self.update_window_title() 
			return
		pane.pdf_document = pdf_doc
		pane.words_data = words_data 
		pane.temp_pdf_path = temp_path
		pane.file_name = display_file_name 
		self.pdf_documents[pane_index] = pdf_doc
		self.words_data_list[pane_index] = [dict(w) for w in words_data]
		self.words_data_original[pane_index] = [dict(w) for w in words_data]
		pane.calculate_document_layout()
		pane.canvas.yview_moveto(0) 
		pane.canvas.xview_moveto(0) 
		pane.fit_to_width() 
		self.update_ui_state() 
		self.update_window_title() 
		self.perform_comparison_if_ready() 
		pane.canvas.focus_set() 
	def perform_comparison_if_ready(self):
		"""Starts word-by-word comparison in a background thread when both PDFs are loaded."""
		doc1_ready = self.pdf_documents[0] and not self.pdf_documents[0].is_closed if self.pdf_documents[0] else False
		doc2_ready = self.pdf_documents[1] and not self.pdf_documents[1].is_closed if self.pdf_documents[1] else False
		if doc1_ready and doc2_ready:
			self._start_comparison_thread()
		else:
			print("Waiting for both documents to be ready for comparison.")
		self.update_ui_state()

	def recompare_documents(self):
		"""Re-runs comparison using cached word data and current compare options."""
		doc1_ready = self.pdf_documents[0] and not self.pdf_documents[0].is_closed if self.pdf_documents[0] else False
		doc2_ready = self.pdf_documents[1] and not self.pdf_documents[1].is_closed if self.pdf_documents[1] else False
		if doc1_ready and doc2_ready:
			self._start_comparison_thread(reset_from_cache=True)
		self.update_ui_state()

	def _start_comparison_thread(self, reset_from_cache=False):
		if self._compare_in_progress:
			return
		if reset_from_cache:
			if self.words_data_list[0] is None or self.words_data_list[1] is None:
				return
		self._compare_in_progress = True
		self.change_records = None
		self.change_records_path = None
		self.pane1.sorted = None
		self.pane2.sorted = None
		self.pane1.display_loading_message("Comparing...")
		self.pane2.display_loading_message("Comparing...")
		self.update_ui_state()
		self._compare_thread = threading.Thread(
			target=self._compare_threaded,
			args=(reset_from_cache,),
			daemon=True,
		)
		self._compare_thread.start()

	def _compare_threaded(self, reset_from_cache=False):
		try:
			if reset_from_cache and self.words_data_original[0] and self.words_data_original[1]:
				words1_copy = [dict(w) for w in self.words_data_original[0]]
				words2_copy = [dict(w) for w in self.words_data_original[1]]
			else:
				words1_copy = [dict(w) for w in self.words_data_list[0]] if self.words_data_list[0] else []
				words2_copy = [dict(w) for w in self.words_data_list[1]] if self.words_data_list[1] else []
			words1_aligned, words2_aligned = align_words(
				words1_copy, words2_copy,
				self.case_insensitive.get(),
				self.ignore_quotes.get(),
				use_chunked=self.chunk_compare.get(),
			)
			self.master.after(0, lambda: self._on_compare_complete(words1_aligned, words2_aligned))
		except Exception as e:
			traceback.print_exc()
			self.master.after(0, lambda: self._on_compare_failed(str(e)))

	def _on_compare_complete(self, words1_aligned, words2_aligned):
		self._compare_in_progress = False
		self.pane1.hide_loading_message()
		self.pane2.hide_loading_message()
		self.pane1.sorted = None
		self.pane2.sorted = None
		self.pane1.words_by_unique_id = None
		self.pane2.words_by_unique_id = None
		self.words_data_list[0] = words1_aligned
		self.words_data_list[1] = words2_aligned
		self.pane1.words_data = words1_aligned
		self.pane2.words_data = words2_aligned
		apply_annotations_to_pdf_pages(self.pdf_documents[0], self.pane1.words_data)
		apply_annotations_to_pdf_pages(self.pdf_documents[1], self.pane2.words_data)
		self.pane1._clear_all_rendered_pages()
		self.pane2._clear_all_rendered_pages()
		self.pane1.render_visible_pages()
		self.pane2.render_visible_pages()
		if self.current_active_pane:
			self.sync_scroll(self.current_active_pane)
		else:
			self.sync_scroll(self.pane1)
		self._refresh_change_records(words1_aligned, words2_aligned)
		self.update_ui_state()
		print("Comparison complete.")

	def _compare_options_dict(self):
		return {
			"case_insensitive": self.case_insensitive.get(),
			"ignore_quotes": self.ignore_quotes.get(),
			"ignore_ligatures": self.ignore_ligatures.get(),
			"split_punctuation": self.split_punctuation.get(),
			"merge_hyphenation": self.merge_hyphenation.get(),
			"group_by_chunk": self.chunk_compare.get(),
			"aligner": "git" if _git_available else "difflib",
		}

	def _refresh_change_records(self, words_left, words_right):
		"""Build structured change records and auto-export JSON after compare."""
		left_name = self.pane1.file_name or "left"
		right_name = self.pane2.file_name or "right"
		self.change_records = build_change_records(
			words_left, words_right,
			left_name=left_name,
			right_name=right_name,
			case_insensitive=self.case_insensitive.get(),
			ignore_quotes=self.ignore_quotes.get(),
			compare_options=self._compare_options_dict(),
			group_by_chunk=self.chunk_compare.get(),
		)
		try:
			self.change_records_path = auto_save_change_records(
				self.change_records, left_name, right_name,
			)
		except Exception as e:
			self.change_records_path = None
			print(f"Warning: could not auto-save change records JSON: {e}", file=sys.stderr)
		print(
			f"Change records: {self.change_records['stats']['total']} hunks "
			f"({self.change_records['stats']['added']} added, "
			f"{self.change_records['stats']['deleted']} deleted, "
			f"{self.change_records['stats']['modified']} modified, "
			f"{self.change_records['stats']['moved']} moved)"
		)
		if self.change_records_path:
			print(f"Change records JSON: {self.change_records_path}")

	def export_change_records_dialog(self):
		"""Save a copy of change records JSON (auto-export already wrote the default file)."""
		if not self.change_records:
			messagebox.showinfo("Export Changes", "No comparison results to export. Compare two documents first.")
			return
		name1 = os.path.splitext(self.pane1.file_name or "left")[0]
		name2 = os.path.splitext(self.pane2.file_name or "right")[0]
		initial = f"{name1}_vs_{name2}_changes.json"
		file_path = filedialog.asksaveasfilename(
			defaultextension=".json",
			filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
			initialfile=initial,
			initialdir=TEMP_PDF_DIR,
		)
		if file_path:
			try:
				write_change_records_json(self.change_records, file_path)
				messagebox.showinfo(
					"Export Changes",
					f"Change records saved to:\n{file_path}"
					+ (f"\n\n(Auto-export remains at:\n{self.change_records_path})" if self.change_records_path else ""),
				)
			except Exception as e:
				messagebox.showerror("Export Changes", f"Failed to save change records:\n{e}")

	def _on_compare_failed(self, error_message):
		self._compare_in_progress = False
		self.pane1.hide_loading_message()
		self.pane2.hide_loading_message()
		messagebox.showerror("Compare Error", f"Comparison failed:\n{error_message}")
		self.update_ui_state()
	def on_pane_scrolled(self, event, source_pane):
		"""Callback for when a user scrolls one of the PDF panes."""
		if self.sync_scroll_enabled.get() and source_pane.pdf_document and not source_pane.pdf_document.is_closed:
			self.sync_scroll(source_pane)
	def sync_scroll(self, source_pane):
		"""Synchronizes the scroll position of the target pane with the source pane."""
		if not self.sync_scroll_enabled.get():
			return
		target_pane = self.pane2 if source_pane.pane_id == 'left' else self.pane1
		if not (source_pane.pdf_document and not source_pane.pdf_document.is_closed and
				target_pane.pdf_document and not target_pane.pdf_document.is_closed):
			return
		if target_pane.ignore_scroll_events_counter > 0:
			return
		source_x, source_y = source_pane.get_current_view_coords()
		source_canvas_height = source_pane.canvas.winfo_height()
		
		# read previous scroll
		prev_scroll_time=self.scroll_time
		prev_scroll_y=self.scroll_y
		prev_scroll_pane=self.scroll_pane
		prev_scroll_height=self.scroll_height
		#set for next one
		time_scroll=time.time()
		self.scroll_time=time_scroll
		self.scroll_y=source_y
		self.scroll_pane=source_pane
		self.scroll_height=source_canvas_height
		
		if source_canvas_height == 0:
			return
		first_common_word_in_view = None
		if not source_pane.sorted:
			source_pane.sorted=sorted(source_pane.words_data, key=lambda x: (x["page_num"], x["y0"], x["x0"]))
		for word_info in source_pane.sorted:
			if word_info["unique_id"] is not None: 
				page_num = word_info["page_num"]
				word_y0_doc = word_info["y0"]
				page_info = source_pane.page_layout_info.get(page_num)
				if not page_info: continue
				word_y_content_coord = (page_info["y_start_offset"] + word_y0_doc) * source_pane.zoom_level
				if word_y_content_coord >= source_y - (source_canvas_height * 0.01): 
					if word_y_content_coord < source_y + source_canvas_height:
						first_common_word_in_view = word_info
						break 
		if first_common_word_in_view:
			common_word_id = first_common_word_in_view["unique_id"]
			if not target_pane.words_by_unique_id:
				target_pane.words_by_unique_id = {
					w["unique_id"]: w for w in target_pane.words_data if w["unique_id"]
				}
			target_word_info = target_pane.words_by_unique_id.get(common_word_id)
			#print(f"\nscroll direction: {source_y-prev_scroll_y}")# positive=we are scrolling down
			#print("source y: ",source_y)
			#print(f"source: {first_common_word_in_view["text"]}, {first_common_word_in_view["page_num"]}, {first_common_word_in_view["x0"]}, {first_common_word_in_view["y0"]},\ntarget: {target_word_info["text"]}, {target_word_info["page_num"]}, {target_word_info["x0"]}, {target_word_info["y0"]}")
			if target_word_info:
				target_page_num = target_word_info["page_num"]
				target_word_y0_doc = target_word_info["y0"]
				target_page_info = target_pane.page_layout_info.get(target_page_num)
				if not target_page_info: return
				source_word_y_content_coord_exact = (first_common_word_in_view["y0"] + source_pane.page_layout_info[first_common_word_in_view["page_num"]]["y_start_offset"]) * source_pane.zoom_level
				y_offset_in_source_view = source_word_y_content_coord_exact - source_y
				target_word_y_content_coord = (target_page_info["y_start_offset"] + target_word_y0_doc) * target_pane.zoom_level
				target_y_scroll_pixels = target_word_y_content_coord - y_offset_in_source_view
				prev_distance=self.scroll_distance
				distance=target_word_y_content_coord-source_y
				prev_target_y=self.scroll_target_y
				is_target_word_visible= (target_word_y_content_coord>target_pane.get_current_view_coords()[1]  and target_word_y_content_coord<target_pane.get_current_view_coords()[1]+target_pane.canvas.winfo_height())
				#print("target_y_scroll_delta: ",target_y_scroll_pixels-prev_target_y)
				#print("target y", target_y_scroll_pixels)
				#print("same direction? ", (source_y-prev_scroll_y)*(target_y_scroll_pixels-prev_target_y)>0)
				#print("is_target_word_visible?",is_target_word_visible)
				if (source_y-prev_scroll_y)*(target_y_scroll_pixels-prev_target_y)>0 or  not is_target_word_visible:
					#print("scrolled!")
					self.scroll_distance=distance
					self.scroll_target_y=target_y_scroll_pixels
					source_x_prop = source_pane.canvas.xview()[0]
					target_x_scroll_pixels = source_x_prop * target_pane.max_document_width * target_pane.zoom_level
					target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll_pixels)
		#else:
		elif 0:#don't scroll target pane if no common words are found in the source pane
			source_x_prop, source_y_prop = source_pane.canvas.xview()[0], source_pane.canvas.yview()[0]
			target_pane._apply_scroll(
				source_x_prop * target_pane.max_document_width * target_pane.zoom_level,
				source_y_prop * target_pane.total_document_height * target_pane.zoom_level
			)
	def sync_zoom(self, source_pane, new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel):
		"""Synchronizes the zoom level of the target pane with the source pane."""
		if not self.sync_zoom_enabled.get():
			return
		target_pane = self.pane2 if source_pane.pane_id == 'left' else self.pane1
		if not (source_pane.pdf_document and not source_pane.pdf_document.is_closed and
				target_pane.pdf_document and not target_pane.pdf_document.is_closed):
			return
		target_pane.set_zoom(new_zoom_level, mouse_x_canvas_pixel, mouse_y_canvas_pixel, from_sync=True)
	def get_word_content_y(self, pane, word_info):
		"""Calculates the word's y-coordinate in content space (document coordinates * zoom)."""
		if not pane.pdf_document or pane.pdf_document.is_closed:
			return -1
		page_num = word_info["page_num"]
		page_info = pane.page_layout_info.get(page_num)
		if not page_info:
			return -1
		return (page_info["y_start_offset"] + word_info["y0"]) * pane.zoom_level
	def is_word_visible(self, pane, word_info):
		"""
		Checks if a word is currently visible in the pane's canvas viewport.
		This checks if ANY part of the word is visible.
		"""
		if not pane.pdf_document or pane.pdf_document.is_closed:
			return False
		view_x, view_y = pane.get_current_view_coords()
		canvas_width = pane.canvas.winfo_width()
		canvas_height = pane.canvas.winfo_height()
		word_x0_content = (word_info["x0"] + (pane.max_document_width - pane.page_layout_info[word_info["page_num"]]["base_width"]) / 2) * pane.zoom_level
		word_y0_content = self.get_word_content_y(pane, word_info)
		word_x1_content = (word_info["x1"] + (pane.max_document_width - pane.page_layout_info[word_info["page_num"]]["base_width"]) / 2) * pane.zoom_level
		word_y1_content = (pane.page_layout_info[word_info["page_num"]]["y_start_offset"] + word_info["y1"]) * pane.zoom_level
		horizontal_overlap = not (word_x1_content < view_x or word_x0_content > (view_x + canvas_width))
		vertical_overlap = not (word_y1_content < view_y or word_y0_content > (view_y + canvas_height))
		return horizontal_overlap and vertical_overlap
	def _find_closest_change(self, direction):#direction=1 -> search down; direction=-1 -> search up
		"""
		Finds the closest (next or previous) highlighted change not currently visible.
		Args:
			direction (int): 1 for next, -1 for previous.
		Returns:
			dict: {pane: PDFViewerPane, word_info: dict, target_y_scroll_pixels: float} or None
		"""
		if not self.pdf_documents[0] or self.pdf_documents[0].is_closed or \
		   not self.pdf_documents[1] or self.pdf_documents[1].is_closed:
			messagebox.showinfo("Navigation Error", "Both PDF documents must be loaded to navigate changes.")
			return None
		panes = [self.pane1, self.pane2]
		current_view_y_pane1 = self.pane1.get_current_view_coords()[1]
		current_view_y_pane2 = self.pane2.get_current_view_coords()[1]
		current_view_height_pane1 = self.pane1.get_current_view_height_in_content_coords()
		current_view_height_pane2 = self.pane2.get_current_view_height_in_content_coords()
		all_highlighted_words = []
		for pane_idx, pane in enumerate(panes):
			for word_idx, word_info in enumerate(pane.words_data):
				if word_info.get("highlight_color"):
					abs_y_pos = self.get_word_content_y(pane, word_info)
					all_highlighted_words.append({
						"pane": pane,
						"word_info": word_info,
						"abs_y_pos": abs_y_pos,
						"pane_index": pane_idx,
						"word_index": word_idx
					})
		if not all_highlighted_words:
			messagebox.showinfo("No Changes", "No highlighted changes found in the documents.")
			return None
		all_highlighted_words.sort(key=lambda x: (x["word_info"]["page_num"], x["abs_y_pos"]))
		mid_y_pane1 = current_view_y_pane1 + current_view_height_pane1 / 2
		mid_y_pane2 = current_view_y_pane2 + current_view_height_pane2 / 2
		closest_unseen_change = None
		min_distance = float('inf')
		for change in all_highlighted_words:
			pane = change["pane"]
			word_info = change["word_info"]
			abs_y_pos = change["abs_y_pos"]
			is_visible = self.is_word_visible(pane, word_info)
			if direction == 1:
				if abs_y_pos > pane.get_current_view_coords()[1] + pane.get_current_view_height_in_content_coords() : 
					distance = abs_y_pos - (pane.get_current_view_coords()[1] + pane.get_current_view_height_in_content_coords())
					if distance +0.00001 < min_distance and distance > 0:
						min_distance = distance
						target_y_scroll = abs_y_pos 
						closest_unseen_change = {"pane": pane, "word_info": word_info, "target_y_scroll": target_y_scroll}
			else: 
				if abs_y_pos < pane.get_current_view_coords()[1] : 
					distance = pane.get_current_view_coords()[1] - abs_y_pos
					if distance +0.00001 < min_distance  and distance > 0:
						min_distance = distance
						word_height_at_zoom = (word_info["y1"] - word_info["y0"]) * pane.zoom_level
						target_y_scroll = abs_y_pos - (pane.get_current_view_height_in_content_coords() - word_height_at_zoom)
						closest_unseen_change = {"pane": pane, "word_info": word_info, "target_y_scroll": target_y_scroll}
		if not closest_unseen_change and all_highlighted_words:
			print("No more changes")
			return None#following code was not working properly (at the beginning/end was creating a loop and not displaying the msessage); when reaching the fist/latest change just do nothing
			if direction == 1: 
				for change in reversed(all_highlighted_words):
					pane = change["pane"]
					word_info = change["word_info"]
					if not self.is_word_visible(pane, word_info):
						abs_y_pos = self.get_word_content_y(pane, word_info)
						target_y_scroll = abs_y_pos 
						return {"pane": pane, "word_info": word_info, "target_y_scroll": target_y_scroll}
				messagebox.showinfo("No More Changes", "You are at the end of the document or all changes are currently visible.")
				return None
			else: 
				for change in all_highlighted_words:
					pane = change["pane"]
					word_info = change["word_info"]
					if not self.is_word_visible(pane, word_info):
						abs_y_pos = self.get_word_content_y(pane, word_info)
						word_height_at_zoom = (word_info["y1"] - word_info["y0"]) * pane.zoom_level
						target_y_scroll = abs_y_pos - (pane.get_current_view_height_in_content_coords() - word_height_at_zoom) 
						return {"pane": pane, "word_info": word_info, "target_y_scroll": target_y_scroll}
				messagebox.showinfo("No More Changes", "You are at the beginning of the document or all changes are currently visible.")
				return None
		return closest_unseen_change
	def go_to_next_change(self):
		"""Moves the view to the next closest highlighted change."""
		print("Attempting to go to next change.")
		change_info = self._find_closest_change(direction=1)
		if change_info:
			target_pane = change_info["pane"]
			target_word_info = change_info["word_info"]
			target_y_scroll = change_info["target_y_scroll"]
			current_x_prop = target_pane.canvas.xview()[0]
			target_x_scroll_pixels = current_x_prop * target_pane.max_document_width * target_pane.zoom_level
			target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll)
			self.sync_scroll(target_pane)
		else:
			print("No next change found or all changes are visible.")
	def go_to_prev_change(self):
		"""Moves the view to the previous closest highlighted change."""
		print("Attempting to go to previous change.")
		change_info = self._find_closest_change(direction=-1)
		if change_info:
			target_pane = change_info["pane"]
			target_word_info = change_info["word_info"]
			target_y_scroll = change_info["target_y_scroll"]
			current_x_prop = target_pane.canvas.xview()[0]
			target_x_scroll_pixels = current_x_prop * target_pane.max_document_width * target_pane.zoom_level
			target_pane._apply_scroll(target_x_scroll_pixels, target_y_scroll)
			self.sync_scroll(target_pane)
		else:
			print("No previous change found or all changes are visible.")
	def on_closing(self):
		"""Handles the application closing event, ensuring PDFs are properly closed and temp files deleted."""
		print("PDFViewerApp: Closing application.")
		self.pane1.close_pdf() 
		self.pane2.close_pdf() 
		self.master.destroy() 


if __name__ == "__main__":
	if len(sys.argv) >= 5 and sys.argv[1] == "--export-changes":
		output_json = sys.argv[2]
		left_pdf = sys.argv[3]
		right_pdf = sys.argv[4]
		try:
			records = compare_files_to_change_records(left_pdf, right_pdf)
			write_change_records_json(records, output_json)
			print(
				f"Exported {records['stats']['total']} change(s): "
				f"{records['stats']['added']} added, {records['stats']['deleted']} deleted, "
				f"{records['stats']['modified']} modified, {records['stats']['moved']} moved."
			)
		except Exception as exc:
			print(f"Export failed: {exc}", file=sys.stderr)
			traceback.print_exc()
			sys.exit(1)
		sys.exit(0)

	root = TkinterDnD.Tk()
	app = PDFViewerApp(root)
	root.protocol("WM_DELETE_WINDOW", app.on_closing)

	root.mainloop()

