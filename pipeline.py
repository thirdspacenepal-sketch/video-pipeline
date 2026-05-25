#!/usr/bin/env python3
"""
Weekly video pipeline — Nepali script + voiceover -> YouTube / Facebook / Instagram / TikTok

HOW IT WORKS
  Script (.docx, Nepali)  ->  Gemini reads content, plans scenes, writes English image prompts
  Voice  (audio, Nepali)  ->  FFprobe measures total duration; each scene duration is
                               proportioned by its Nepali word count vs. total word count

FREE  pipeline  (USE_AI_IMAGES=false):  PIL generates Devanagari scene cards  (no cost)
PAID  pipeline  (USE_AI_IMAGES=true):   Together.ai FLUX Schnell generates illustrations  (~$1-3/mo)

Everything else — Drive fetch, Gemini planning, FFmpeg assembly, all publishing — is identical.
"""

import os, json, subprocess, tempfile, textwrap, time, requests
from pathlib import Path
from io import BytesIO

import google.generativeai as genai
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from docx import Document
from PIL import Image, ImageDraw, ImageFont


# ── Config ────────────────────────────────────────────────────────────────────

USE_AI_IMAGES     = os.getenv("USE_AI_IMAGES", "false").lower() == "true"
DRIVE_FOLDER_ID   = os.getenv("DRIVE_FOLDER_ID", "")
GEMINI_API_KEY    = os.getenv("GEMINI_API_KEY", "")
TOGETHER_API_KEY  = os.getenv("TOGETHER_API_KEY", "")      # paid pipeline only
YT_CLIENT_ID      = os.getenv("YT_CLIENT_ID", "")
YT_CLIENT_SECRET  = os.getenv("YT_CLIENT_SECRET", "")
YT_REFRESH_TOKEN  = os.getenv("YT_REFRESH_TOKEN", "")
FB_ACCESS_TOKEN   = os.getenv("FB_ACCESS_TOKEN", "")
FB_PAGE_ID        = os.getenv("FB_PAGE_ID", "")
IG_USER_ID        = os.getenv("IG_USER_ID", "")            # Instagram Business account ID

VIDEO_W, VIDEO_H  = 1920, 1080
FPS               = 25
SCENE_BG          = (12, 14, 28)                            # dark navy background
ACCENT_COLORS     = [(80, 100, 220), (60, 180, 140), (220, 100, 60), (160, 80, 200)]

# Noto Sans Devanagari — installed by the GitHub Actions workflow (apt)
DEVANAGARI_FONT_PATHS = [
    "/usr/share/fonts/truetype/noto/NotoSansDevanagari-Regular.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansDevanagari[wdth,wght].ttf",
    "/usr/share/fonts/noto/NotoSansDevanagari-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",     # fallback (no Devanagari, won't crash)
]


# ── 1. Google Drive: fetch latest .docx and audio ────────────────────────────

def _drive_credentials(scopes):
    info = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON", "{}"))
    creds = Credentials(
        token=None,
        refresh_token=info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        scopes=scopes,
    )
    creds.refresh(Request())
    return creds


def fetch_drive_files(work_dir):
    """Returns (docx_path, audio_path) — most recently modified of each type."""
    svc = build("drive", "v3", credentials=_drive_credentials(
        ["https://www.googleapis.com/auth/drive.readonly"]
    ))

    def latest(mime_query):
        res = svc.files().list(
            q=f"'{DRIVE_FOLDER_ID}' in parents and {mime_query} and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name)",
        ).execute()
        files = res.get("files", [])
        if not files:
            raise FileNotFoundError(f"No file matching: {mime_query}")
        return files[0]

    def download(meta, dest):
        req = svc.files().get_media(fileId=meta["id"])
        with open(dest, "wb") as fh:
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
        print(f"  Downloaded: {meta['name']}")
        return dest

    docx_meta  = latest("mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'")
    audio_meta = latest("mimeType contains 'audio'")
    docx_path  = download(docx_meta,  work_dir / "script.docx")
    audio_path = download(audio_meta, work_dir / f"voice{Path(audio_meta['name']).suffix}")
    return docx_path, audio_path


# ── 2. Parse .docx (Nepali text) ─────────────────────────────────────────────

def parse_docx(docx_path):
    doc = Document(docx_path)
    return "\n\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())


# ── 3. Scene planning with Gemini Flash ──────────────────────────────────────
#
#   The script is in Nepali. Gemini handles it natively.
#   IMPORTANT: image_prompt must be in English — Together.ai / FLUX give far
#   better results with English prompts. Gemini translates the visual intent.
#   video_title and video_description stay in Nepali for upload metadata.

GEMINI_PROMPT = """\
You are a video editor's assistant. The script below is written in Nepali.
Read it carefully and break it into 8-12 scenes.

Rules:
- "title"          : Nepali scene title, max 6 words (shown as overlay text on the frame)
- "narration_text" : the exact Nepali script excerpt for this scene
- "image_prompt"   : MUST BE IN ENGLISH — vivid illustration description for an AI image model
                     Include style, subject, mood, lighting, and color palette
- "video_title"    : full video title in Nepali (for YouTube/Facebook)
- "video_description": video description in Nepali (for YouTube/Facebook)

Return ONLY valid JSON — no markdown fences, no extra text:
{
  "video_title": "...",
  "video_description": "...",
  "scenes": [
    {
      "title": "नेपाली शीर्षक",
      "narration_text": "exact Nepali script excerpt",
      "image_prompt": "English illustration prompt for AI image generation"
    }
  ]
}

Script:
"""

def plan_scenes(nepali_script):
    genai.configure(api_key=GEMINI_API_KEY)
    model  = genai.GenerativeModel("gemini-1.5-flash")
    result = model.generate_content(GEMINI_PROMPT + nepali_script)
    raw    = result.text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
    data   = json.loads(raw)
    print(f"  Planned {len(data['scenes'])} scenes: \"{data['video_title']}\"")
    return data


# ── 4. Frame timing from Nepali voiceover ────────────────────────────────────
#
#   Total video length  = actual audio duration (FFprobe on the Nepali voice file)
#   Each scene duration = (scene Nepali word count / total words) * audio duration
#
#   Nepali words in Devanagari are space-separated, so .split() works correctly.

def get_audio_duration(audio_path):
    r = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=noprint_wrappers=1:nokey=1", str(audio_path)],
        capture_output=True, text=True, check=True,
    )
    return float(r.stdout.strip())


def compute_scene_durations(scenes, audio_path):
    """
    Returns list of float durations (one per scene) that sum to audio length.
    Proportioned by each scene's Nepali word count.
    """
    total_audio = get_audio_duration(audio_path)
    word_counts = [len(s["narration_text"].split()) for s in scenes]
    total_words = sum(word_counts) or 1

    raw = [max(1.5, (wc / total_words) * total_audio) for wc in word_counts]

    # Normalise so total exactly matches audio length (rounding may drift)
    scale     = total_audio / sum(raw)
    durations = [d * scale for d in raw]

    print(f"  Audio: {total_audio:.1f}s across {len(scenes)} scenes "
          f"({min(durations):.1f}s - {max(durations):.1f}s per scene)")
    return durations


# ── 5a. Free visuals: PIL Devanagari scene cards ─────────────────────────────

def _load_font(size):
    for path in DEVANAGARI_FONT_PATHS:
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def generate_scene_image_free(scene, index, work_dir):
    img  = Image.new("RGB", (VIDEO_W, VIDEO_H), color=SCENE_BG)
    draw = ImageDraw.Draw(img)

    # Colored accent strip at top
    draw.rectangle([(0, 0), (VIDEO_W, 12)], fill=ACCENT_COLORS[index % len(ACCENT_COLORS)])

    # Nepali title (large, upper-center)
    draw.text(
        (VIDEO_W // 2, VIDEO_H // 2 - 110),
        scene.get("title", f"दृश्य {index + 1}"),
        font=_load_font(84), fill=(240, 240, 255), anchor="mm",
    )

    # Nepali narration excerpt (wrapped, lower-center)
    excerpt = scene.get("narration_text", "")[:220]
    draw.multiline_text(
        (VIDEO_W // 2, VIDEO_H // 2 + 90),
        textwrap.fill(excerpt, width=52),
        font=_load_font(40), fill=(160, 165, 195),
        anchor="mm", align="center", spacing=14,
    )

    # Scene number (bottom-right, subtle)
    draw.text(
        (VIDEO_W - 48, VIDEO_H - 36), str(index + 1),
        font=_load_font(28), fill=(70, 75, 105), anchor="rm",
    )

    out = work_dir / f"scene_{index:03d}.png"
    img.save(out, format="PNG")
    return out


# ── 5b. Paid visuals: Together.ai FLUX Schnell ───────────────────────────────

def generate_scene_image_ai(scene, index, work_dir):
    """
    Uses the English image_prompt that Gemini generated from the Nepali scene text.
    English prompts give significantly better results with FLUX.
    """
    r = requests.post(
        "https://api.together.xyz/v1/images/generations",
        headers={"Authorization": f"Bearer {TOGETHER_API_KEY}",
                 "Content-Type": "application/json"},
        json={
            "model":  "black-forest-labs/FLUX.1-schnell-Free",
            "prompt": scene["image_prompt"],
            "width":  1792,
            "height": 1024,
            "steps":  4,
            "n":      1,
        },
        timeout=90,
    )
    r.raise_for_status()
    url  = r.json()["data"][0]["url"]
    raw  = requests.get(url, timeout=30).content
    out  = work_dir / f"scene_{index:03d}.png"
    Image.open(BytesIO(raw)).convert("RGB").resize((VIDEO_W, VIDEO_H)).save(out)
    return out


# ── 6. Assemble with FFmpeg ───────────────────────────────────────────────────

def assemble_video(image_paths, durations, audio_path, work_dir):
    """
    Builds the concat manifest (each image held for its computed duration),
    then runs FFmpeg to mux visuals + Nepali voiceover into final MP4.
    """
    concat_file = work_dir / "concat.txt"
    with open(concat_file, "w") as f:
        for img, dur in zip(image_paths, durations):
            f.write(f"file '{img.resolve()}'\n")
            f.write(f"duration {dur:.4f}\n")
        # FFmpeg concat demuxer requires the last file repeated without a duration line
        f.write(f"file '{image_paths[-1].resolve()}'\n")

    out = work_dir / "final_video.mp4"
    subprocess.run([
        "ffmpeg", "-y",
        "-f", "concat", "-safe", "0", "-i", str(concat_file),
        "-i", str(audio_path),
        "-vf", (
            f"fps={FPS},"
            f"scale={VIDEO_W}:{VIDEO_H}:force_original_aspect_ratio=decrease,"
            f"pad={VIDEO_W}:{VIDEO_H}:(ow-iw)/2:(oh-ih)/2:color=black"
        ),
        "-c:v", "libx264", "-preset", "fast", "-crf", "22",
        "-c:a", "aac", "-b:a", "192k",
        "-pix_fmt", "yuv420p",
        "-shortest",
        str(out),
    ], check=True)

    print(f"  Output: {out.name}  ({out.stat().st_size / 1e6:.1f} MB)")
    return out


# ── 7. Publish ────────────────────────────────────────────────────────────────

def _youtube_service():
    creds = Credentials(
        token=None, refresh_token=YT_REFRESH_TOKEN,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=YT_CLIENT_ID, client_secret=YT_CLIENT_SECRET,
        scopes=["https://www.googleapis.com/auth/youtube.upload"],
    )
    creds.refresh(Request())
    return build("youtube", "v3", credentials=creds)


def publish_youtube(video_path, title, description):
    print("  Uploading to YouTube...")
    svc   = _youtube_service()
    media = MediaFileUpload(str(video_path), mimetype="video/mp4", resumable=True)
    req   = svc.videos().insert(
        part="snippet,status",
        body={"snippet":  {"title": title, "description": description, "categoryId": "22"},
              "status":   {"privacyStatus": "public"}},
        media_body=media,
    )
    resp = None
    while resp is None:
        _, resp = req.next_chunk()
    print(f"    YouTube: https://youtu.be/{resp['id']}")


def publish_facebook(video_path, title, description):
    print("  Uploading to Facebook...")
    with open(video_path, "rb") as vf:
        r = requests.post(
            f"https://graph-video.facebook.com/{FB_PAGE_ID}/videos",
            data={"title": title, "description": description, "access_token": FB_ACCESS_TOKEN},
            files={"source": vf},
            timeout=300,
        )
    r.raise_for_status()
    print(f"    Facebook video ID: {r.json().get('id')}")


def publish_instagram_reel(caption, public_video_url):
    """
    Instagram Reels requires a publicly accessible video URL.
    Upload the MP4 to a public CDN first, then pass the URL here.
    See README for CDN options.
    """
    print("  Publishing to Instagram Reels...")
    r1 = requests.post(
        f"https://graph.facebook.com/v19.0/{IG_USER_ID}/media",
        params={"media_type": "REELS", "video_url": public_video_url,
                "caption": caption[:2200], "access_token": FB_ACCESS_TOKEN},
        timeout=60,
    )
    r1.raise_for_status()
    container_id = r1.json()["id"]

    for _ in range(10):
        time.sleep(15)
        status = requests.get(
            f"https://graph.facebook.com/v19.0/{container_id}",
            params={"fields": "status_code", "access_token": FB_ACCESS_TOKEN},
        ).json().get("status_code")
        if status == "FINISHED":
            break
        print(f"    Waiting for IG processing... ({status})")

    r2 = requests.post(
        f"https://graph.facebook.com/v19.0/{IG_USER_ID}/media_publish",
        params={"creation_id": container_id, "access_token": FB_ACCESS_TOKEN},
        timeout=60,
    )
    r2.raise_for_status()
    print(f"    Instagram Reel ID: {r2.json().get('id')}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    mode = "Together.ai FLUX (paid)" if USE_AI_IMAGES else "PIL Devanagari cards (free)"
    print(f"\n=== Weekly Video Pipeline  |  visuals: {mode} ===\n")

    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)

        print("1. Fetching files from Google Drive...")
        docx_path, audio_path = fetch_drive_files(work)

        print("2. Parsing Nepali script from .docx...")
        script = parse_docx(docx_path)
        print(f"  {len(script.split())} words extracted")

        print("3. Planning scenes with Gemini Flash (Nepali -> English image prompts)...")
        plan      = plan_scenes(script)
        scenes    = plan["scenes"]

        print("4. Computing frame durations from Nepali voiceover...")
        durations = compute_scene_durations(scenes, audio_path)

        gen_fn = generate_scene_image_ai if USE_AI_IMAGES else generate_scene_image_free
        print(f"5. Generating {len(scenes)} visuals ({'Together.ai' if USE_AI_IMAGES else 'PIL'})...")
        image_paths = []
        for i, scene in enumerate(scenes):
            print(f"  [{i+1}/{len(scenes)}] {scene['title']}")
            image_paths.append(gen_fn(scene, i, work))

        print("6. Assembling video with FFmpeg...")
        video_path = assemble_video(image_paths, durations, audio_path, work)

        print("7. Publishing...")
        publish_youtube(video_path, plan["video_title"], plan["video_description"])
        if FB_ACCESS_TOKEN and FB_PAGE_ID:
            publish_facebook(video_path, plan["video_title"], plan["video_description"])
        # Uncomment when Instagram CDN is configured:
        # publish_instagram_reel(plan["video_description"], upload_to_public_cdn(video_path))

    print("\nDone.")


if __name__ == "__main__":
    main()
