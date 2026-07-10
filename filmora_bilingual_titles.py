#!/usr/bin/env python3
"""
filmora_bilingual_titles.py
============================
Automation for turning a bilingual (Chinese/Indonesian) Word script into
Wondershare Filmora "z1"/"Z2" title clips, automatically:

  1. Reading the Word doc and pairing up each Chinese line with its
     Indonesian translation (auto-skips headings/links/summary at the top --
     no need to know how many intro paragraphs there are).
  2. Choosing the z1 or Z2 custom title template per line, based on the
     Indonesian line's character count (default rule: <=58 -> z1, >=59 -> Z2).
  3. Writing brand-new title clips directly into a copy of a Filmora .wfp
     project file -- no clicking inside Filmora required.

HOW IT WORKS / LIMITATIONS (read this before relying on it):
  - Needs a "shell" .wfp project that already contains at least one z1 title
    clip and one Z2 title clip (any project where you've used both once is
    fine -- their exact text/timing doesn't matter, they're just donors).
  - Timing is NOT read from audio. New clips are placed back-to-back starting
    either at 0:00 (--mode replace) or right after whatever's already in the
    shell (--mode append), with duration estimated from reading speed
    (default 15 Indonesian characters/second, 2s minimum). You will still
    need to nudge clip timing in Filmora to match the actual spoken audio.
  - The z1/Z2 split is a character-count approximation of true on-screen
    pixel width, so it won't be 100% perfect on unusually wide/narrow text --
    expect the rare miss, same as doing it by hand.
  - Always run this on a COPY of your project. Keep your real .wfp safe.

USAGE
-----
Replace all titles in a fresh project (starts at 0:00):
    python filmora_bilingual_titles.py script.docx shell.wfp output.wfp

Append new titles after whatever's already in the project:
    python filmora_bilingual_titles.py script.docx shell.wfp output.wfp --mode append

Custom threshold or reading speed:
    python filmora_bilingual_titles.py script.docx shell.wfp output.wfp --threshold 60 --cps 13
"""
import re
import json
import uuid
import zipfile
import shutil
import os
import argparse
import sys

try:
    import docx
except ImportError:
    docx = None


# ---------------------------------------------------------------------------
# Word doc reading
# ---------------------------------------------------------------------------
CJK_RE = re.compile(r'[\u4e00-\u9fff]')


def has_cjk(s: str) -> bool:
    return bool(CJK_RE.search(s))


def find_content_start(paragraphs, window=16):
    """
    Auto-detect the index where (source_line, Indonesian_line) content begins.
    Uses a local window (not the whole remainder) so a one-off exception later
    in the document -- e.g. a line the speaker said in English instead of
    Chinese -- doesn't prevent detecting the real start.
    """
    n = len(paragraphs)
    for i in range(n):
        w = min(window, n - i)
        if w < 10 or w % 2 != 0:
            continue
        segment = paragraphs[i:i + w]
        if all(has_cjk(p) == (j % 2 == 0) for j, p in enumerate(segment)):
            return i
    raise ValueError(
        "Could not find a clean alternating Chinese/Indonesian block in this "
        "document. Check that every Chinese line is immediately followed by "
        "its Indonesian translation, one per paragraph."
    )


def extract_pairs(docx_path, verbose=True):
    if docx is None:
        sys.exit("Missing dependency: pip install python-docx")
    d = docx.Document(docx_path)
    non_empty = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    start = find_content_start(non_empty)
    content = non_empty[start:]
    if len(content) % 2 != 0:
        raise ValueError(
            f"After skipping the header (paragraph 0-{start-1}), {len(content)} "
            f"content paragraphs remain, which is an ODD number -- one line is "
            f"missing or two got merged somewhere. Last paragraph found: "
            f"{content[-1]!r}. Check the end of the document."
        )
    pairs = list(zip(content[0::2], content[1::2]))
    if verbose:
        for idx, (src, ind) in enumerate(pairs):
            if not has_cjk(src):
                para_no = start + idx * 2
                print(f"  [info] pair {idx} (paragraph #{para_no}) has a non-Chinese "
                      f"source line -- please double check: {src!r} -> {ind!r}")
    return pairs


# ---------------------------------------------------------------------------
# .wfp raw-text editing (surgical splicing -- see conversation for why)
# ---------------------------------------------------------------------------
def hex_utf16be_encode(text: str) -> str:
    return text.encode('utf-16-be').hex().upper()


def hex_utf16be_decode(hexstr: str) -> str:
    return bytes.fromhex(hexstr).decode('utf-16-be')


def find_toplevel_blocks(raw):
    tag_re = re.compile(r'<TimeLine>|</TimeLine>')
    stack, spans, depth = [], [], 0
    for m in tag_re.finditer(raw):
        if m.group() == '<TimeLine>':
            stack.append((m.start(), depth)); depth += 1
        else:
            start, d = stack.pop(); depth -= 1
            spans.append((start, m.end(), d))
    return spans


def children_of(spans, start, end, depth):
    kids = [s for s in spans if s[2] == depth + 1 and start < s[0] and s[1] <= end]
    kids.sort(key=lambda s: s[0])
    return kids


def sub1(pattern, repl_fn, text, label):
    matches = list(re.finditer(pattern, text, re.DOTALL))
    if len(matches) != 1:
        raise ValueError(f"{label}: expected exactly 1 match in this scope, got {len(matches)}")
    m = matches[0]
    return text[:m.start(1)] + repl_fn(m.group(1)) + text[m.end(1):]


def set_child_text_and_duration(child_text, new_text, new_end_frame):
    def caption_sub(escaped_val):
        unescaped = (escaped_val.replace('&quot;', '"').replace('&gt;', '>')
                                 .replace('&lt;', '<').replace('&amp;', '&'))
        new_hex = hex_utf16be_encode(new_text)
        new_unescaped = re.sub(r'(Title=")[0-9A-Fa-f]+(")', lambda m: m.group(1) + new_hex + m.group(2),
                                unescaped, count=1)
        caption_sub.new_len = len(new_unescaped)
        return (new_unescaped.replace('&', '&amp;').replace('<', '&lt;')
                              .replace('>', '&gt;').replace('"', '&quot;'))
    child_text = sub1(r'<Property key="Text.Section.CaptionData" type="std::wstring" value="(.*?)"/>',
                       caption_sub, child_text, "child CaptionData")
    child_text = sub1(r'<Property key="Text.Section.CaptionDataLength" type="int" value="(\d+)"/>',
                       lambda old: str(caption_sub.new_len), child_text, "child CaptionDataLength")
    child_text = sub1(r'<Property key="Render.RangeFrameNumber" type="NLERange" value="Start:0, End:(\d+)"/>',
                       lambda old: str(new_end_frame), child_text, "child RangeFrameNumber")
    return child_text


def clone_clip(donor_block, chinese_text, indonesian_text, new_position, new_end_frame):
    spans = find_toplevel_blocks(donor_block)
    root2 = max(spans, key=lambda s: s[1] - s[0])
    kids = children_of(spans, root2[0], root2[1], root2[2])
    if len(kids) != 2:
        raise ValueError(f"expected 2 child text elements in donor clip, found {len(kids)}")

    header = donor_block[:kids[0][0]]
    child0 = donor_block[kids[0][0]:kids[0][1]]
    gap    = donor_block[kids[0][1]:kids[1][0]]
    child1 = donor_block[kids[1][0]:kids[1][1]]
    tail   = donor_block[kids[1][1]:]

    new_guid = "{" + str(uuid.uuid4()) + "}"
    header = sub1(r'<Property key="Key_TimelineGuid" type="std::wstring" value="(.*?)"/>',
                  lambda old: new_guid, header, "parent Key_TimelineGuid")
    header = sub1(r'<Property key="Render.Position" type="int" value="(\d+)"/>',
                  lambda old: str(new_position), header, "parent Render.Position")
    header = sub1(r'<Property key="Render.RangeFrameNumber" type="NLERange" value="Start:0, End:(\d+)"/>',
                  lambda old: str(new_end_frame), header, "parent RangeFrameNumber")
    header = sub1(r'<Property key="Render.ContentRange" type="NLERange" value="Start:0, End:(\d+)"/>',
                  lambda old: str(new_end_frame), header, "parent ContentRange")

    child0 = set_child_text_and_duration(child0, chinese_text, new_end_frame)
    child1 = set_child_text_and_duration(child1, indonesian_text, new_end_frame)
    return header + child0 + gap + child1 + tail


def get_clip_name(block):
    m = re.search(r'<Property key="key_timeline_resource_info" type="std::wstring" value="(.*?)"/>', block, re.DOTALL)
    return json.loads(m.group(1).replace('&quot;', '"')).get('name')


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def generate(docx_path, shell_wfp, output_wfp, mode='replace',
             threshold=58, cps=15.0, min_duration=2.0, gap_frames=0, fps=None):
    line_pairs = extract_pairs(docx_path)

    workdir = "/tmp/_wfp_build_final"
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(shell_wfp) as z:
        z.extractall(workdir)

    proj_path = os.path.join(workdir, "WSVEFolder", "Project", "project.xml")
    info_path = os.path.join(workdir, "WSVEFolder", "project_info.json")
    raw = open(proj_path, 'r', encoding='utf-8', newline='').read()
    info = json.load(open(info_path, encoding='utf-8'))
    fps = fps or info.get('document_framerate', 25)

    spans = find_toplevel_blocks(raw)
    root_start, root_end, root_depth = spans[-1]
    clips = children_of(spans, root_start, root_end, root_depth)

    donor_z1 = donor_z2 = None
    for cs, ce, _ in clips:
        name = get_clip_name(raw[cs:ce])
        if name == 'z1' and donor_z1 is None:
            donor_z1 = raw[cs:ce]
        elif name == 'Z2' and donor_z2 is None:
            donor_z2 = raw[cs:ce]
    if not donor_z1 or not donor_z2:
        raise ValueError("shell .wfp must contain at least one existing z1 clip and one Z2 clip to use as donors")

    if mode == 'append' and clips:
        last_block = raw[clips[-1][0]:clips[-1][1]]
        cursor = int(re.search(r'Render.Position" type="int" value="(\d+)"', last_block).group(1))
        cursor += int(re.search(r'Render.RangeFrameNumber" type="NLERange" value="Start:0, End:(\d+)"', last_block).group(1)) + 1
        keep_existing = True
    else:
        cursor = 0
        keep_existing = False

    new_blocks, report = [], []
    for chinese, indonesian in line_pairs:
        template = 'z1' if len(indonesian) <= threshold else 'Z2'
        donor = donor_z1 if template == 'z1' else donor_z2
        duration_frames = max(round(min_duration * fps), round(len(indonesian) / cps * fps))
        end_frame = duration_frames - 1
        new_blocks.append(clone_clip(donor, chinese, indonesian, cursor, end_frame))
        report.append((template, len(indonesian), cursor, duration_frames, chinese, indonesian))
        cursor += duration_frames + gap_frames

    separator = raw[clips[-2][1]:clips[-1][0]] if len(clips) >= 2 else '\r\n'
    if keep_existing:
        new_raw = raw[:clips[-1][1]] + separator + separator.join(new_blocks) + raw[clips[-1][1]:]
        total_children = len(clips) + len(line_pairs)
    else:
        new_raw = raw[:clips[0][0]] + separator.join(new_blocks) + raw[clips[-1][1]:] if clips else raw
        total_children = len(line_pairs)

    new_total_frames = cursor - gap_frames
    new_raw = re.sub(r'(<Property key="Children.Count" type="int" value=")\d+(")',
                      lambda m: m.group(1) + str(total_children) + m.group(2), new_raw, count=1)
    new_raw = re.sub(r'(<Property key="Render.ContentRange" type="NLERange" value="Start:0, End:)\d+(")',
                      lambda m: m.group(1) + str(new_total_frames - 1) + m.group(2), new_raw, count=1)
    new_raw = re.sub(r'(<Property key="Render.RangeFrameNumber" type="NLERange" value="Start:0, End:)\d+(")',
                      lambda m: m.group(1) + str(new_total_frames - 1) + m.group(2), new_raw, count=1)

    open(proj_path, 'w', encoding='utf-8', newline='').write(new_raw)
    info['document_duration'] = new_total_frames
    json.dump(info, open(info_path, 'w', encoding='utf-8'), indent=4)

    if os.path.exists(output_wfp):
        os.remove(output_wfp)
    with zipfile.ZipFile(output_wfp, 'w', zipfile.ZIP_STORED) as zf:
        for foldername, _, filenames in os.walk(workdir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                zf.write(filepath, os.path.relpath(filepath, workdir))

    return report


def main():
    ap = argparse.ArgumentParser(description="Generate Filmora z1/Z2 titles from a bilingual Word script.")
    ap.add_argument("docx", help="Path to the bilingual .docx script")
    ap.add_argument("shell_wfp", help="Path to a .wfp project containing at least one z1 and one Z2 title clip")
    ap.add_argument("output_wfp", help="Where to save the generated .wfp")
    ap.add_argument("--mode", choices=["replace", "append"], default="replace",
                     help="replace: fresh timeline from 0:00 (default). append: add after existing clips.")
    ap.add_argument("--threshold", type=int, default=58, help="Max Indonesian char count for z1 (default 58)")
    ap.add_argument("--cps", type=float, default=15.0, help="Assumed reading speed, chars/sec (default 15)")
    args = ap.parse_args()

    report = generate(args.docx, args.shell_wfp, args.output_wfp, mode=args.mode,
                       threshold=args.threshold, cps=args.cps)
    z1n = sum(1 for r in report if r[0] == 'z1')
    z2n = sum(1 for r in report if r[0] == 'Z2')
    print(f"Done: {len(report)} title clips written ({z1n} x z1, {z2n} x Z2) -> {args.output_wfp}")


if __name__ == '__main__':
    main()
