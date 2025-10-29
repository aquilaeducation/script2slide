def build_vtt_files(blocks):
    out = {}
    idx = 1
    for b in blocks:
        if b["type"] == "slide" and (b.get("narration") or "").strip():
            name = f"slide_{idx:02d}.vtt"
            body = "WEBVTT\n\n00:00.000 --> 00:06.000\n" + b["narration"].strip() + "\n"
            out[name] = body
            idx += 1
    return out
