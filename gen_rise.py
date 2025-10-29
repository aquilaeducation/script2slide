def build_rise_blocks(blocks, course_title):
    lines = [f"# Lesson: {course_title}", ""]
    for b in blocks:
        if b["type"] == "slide":
            lines += [
                "## Text Block",
                f"**{b.get('title','')}**",
                b.get("narration","").strip(),
                ""
            ]
            if b.get("bullets"):
                for bullet in b["bullets"]:
                    lines.append(f"- {bullet}")
                lines.append("")
            if b.get("image"):
                alt = b.get("alt","").strip()
                lines.append(f"_Image (upload {b['image']}). Alt: \"{alt}\"_")
                lines.append("")
            lines.append("---")
            lines.append("")
        elif b["type"] == "quiz_single":
            lines += [
                "## Knowledge Check (Single Choice)",
                f"**{b.get('question','')}**"
            ]
            choices = b.get("choices") or []
            ci = b.get("correctIndex")
            for i, c in enumerate(choices):
                if i == ci:
                    lines.append(f"- **{c}**   ← correct")
                else:
                    lines.append(f"- {c}")
            lines += [
                "",
                f"Feedback (Correct): {b.get('feedbackCorrect','')}",
                f"Feedback (Incorrect): {b.get('feedbackIncorrect','')}",
                "",
                "---",
                ""
            ]
    return "\n".join(lines).strip() + "\n"
