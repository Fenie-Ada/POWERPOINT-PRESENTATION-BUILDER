import os
import re
import sys
from datetime import datetime
from io import BytesIO

import requests
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# colours used on the slides
DARK_BLUE = RGBColor(31, 56, 100)
DARK_GREY = RGBColor(64, 64, 64)
LIGHT_GREY = RGBColor(210, 215, 225)
WHITE = RGBColor(255, 255, 255)

MAX_BULLETS = 6  # most bullet points on one slide, the rest go on a "(cont.)" slide


def clean_line(line):
    # removes bullet symbols (- * • #) and extra spaces from the start of a line
    return line.strip().lstrip("-*•#").strip()


def clean_title(name):
    # turns a file name like "meeting_notes" into "Meeting Notes"
    return name.replace("_", " ").replace("-", " ").strip().title()


def make_filename(title):
    # turns a title into a safe file name, e.g. "Meeting Notes" -> "Meeting_Notes.pptx"
    safe = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")
    return (safe or "presentation") + ".pptx"


# ---------- reading the input ----------

def read_text_notes(text):
    # turns plain text into a list of slides: [{"title": ..., "points": [...]}]
    # each block of lines (separated by a blank line) becomes one slide
    # the first line of a block is the slide title, the other lines are bullet points
    slides = []
    block = []

    for line in text.splitlines() + [""]:  # the extra "" makes sure the last block is saved
        if line.strip() == "":
            if block:
                slides.append({"title": block[0], "points": block[1:]})
                block = []
        else:
            cleaned = clean_line(line)
            if cleaned:
                block.append(cleaned)

    return slides


def read_word_file(source, fallback_title):
    # reads a Word file (a file path, or a file held in memory) and returns (title, slides)
    try:
        document = Document(source)
    except PackageNotFoundError:
        raise ValueError(
            "Could not open that Word file. Make sure it is a real .docx file "
            "(old .doc files are not supported; save it as .docx first)."
        )

    title = document.core_properties.title or fallback_title
    paragraphs = document.paragraphs
    has_headings = any(p.style.name.startswith("Heading") for p in paragraphs)

    if not has_headings:
        # no headings, so treat it like plain notes (a blank line starts a new slide)
        text = "\n".join(p.text for p in paragraphs)
        return title, read_text_notes(text)

    slides = []
    current = None  # the slide we are adding points to

    for paragraph in paragraphs:
        text = paragraph.text.strip()
        style = paragraph.style.name
        if not text:
            continue

        if style == "Title":
            title = text  # the document's title becomes the presentation title
        elif style.startswith("Heading"):
            current = {"title": text, "points": []}  # every heading starts a new slide
            slides.append(current)
        else:
            if current is None:  # text that comes before the first heading
                current = {"title": "Introduction", "points": []}
                slides.append(current)
            cleaned = clean_line(text)
            if cleaned:
                current["points"].append(cleaned)

    return title, slides


def download_google_doc(link):
    # downloads a Google Doc as a Word file and returns it in memory
    # the doc must be shared as "Anyone with the link can view"
    match = re.search(r"/document/d/([\w-]+)", link)
    if not match:
        raise ValueError("That doesn't look like a Google Docs link.")

    export_url = f"https://docs.google.com/document/d/{match.group(1)}/export?format=docx"
    response = requests.get(export_url, timeout=15)

    # a real Word file starts with "PK", anything else means Google sent a sign-in page
    if response.status_code != 200 or not response.content.startswith(b"PK"):
        raise ValueError(
            "Could not download the Google Doc. Open it, click Share, set General access "
            "to 'Anyone with the link' (Viewer), then try again."
        )

    return BytesIO(response.content)


def read_source(source):
    # works out what kind of input it is and returns (title, slides)
    lower = source.lower()

    if lower.startswith("http"):
        if "docs.google.com/document" not in lower:
            raise ValueError("Only Google Docs links are supported.")
        return read_word_file(download_google_doc(source), "Presentation")

    if not os.path.isfile(source):
        raise FileNotFoundError(source)

    name = clean_title(os.path.splitext(os.path.basename(source))[0])

    if lower.endswith(".txt"):
        # utf-8-sig also handles files saved by Notepad with a hidden marker at the start
        with open(source, "r", encoding="utf-8-sig", errors="replace") as file:
            return name, read_text_notes(file.read())

    if lower.endswith(".docx"):
        return read_word_file(source, name)

    raise ValueError("Unsupported file type. Use a .txt file, a .docx file or a Google Docs link.")


# ---------- building the presentation ----------

def split_points(points):
    # splits a long list of bullet points into groups of MAX_BULLETS
    if not points:
        return [[]]  # one slide with a title only
    return [points[i:i + MAX_BULLETS] for i in range(0, len(points), MAX_BULLETS)]


def font_size_for(points):
    # slides with a lot of text get smaller text so it fits
    total = sum(len(point) for point in points)
    if total > 700:
        return 16
    if total > 450:
        return 20
    return 24


def place(shape, left, top, width, height):
    # puts a shape at a position on the slide (measurements in inches)
    shape.left = Inches(left)
    shape.top = Inches(top)
    shape.width = Inches(width)
    shape.height = Inches(height)


def style_text(frame, text, size, color, bold, align):
    # writes one piece of text into a text box and formats it
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def write_bullets(frame, points):
    # writes each point as a bullet in the text box
    frame.word_wrap = True
    size = font_size_for(points)

    for index, point in enumerate(points):
        if index == 0:
            paragraph = frame.paragraphs[0]
        else:
            paragraph = frame.add_paragraph()
        paragraph.space_after = Pt(10)

        run = paragraph.add_run()
        run.text = point
        run.font.size = Pt(size)
        run.font.color.rgb = DARK_GREY


def add_title_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # layout 0 = title slide

    # dark blue background
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = DARK_BLUE

    title_size = 44 if len(title) <= 40 else 32  # long titles need smaller text
    place(slide.shapes.title, 1, 2.3, 11.33, 1.6)
    style_text(slide.shapes.title.text_frame, title, title_size, WHITE, True, PP_ALIGN.CENTER)

    date_box = slide.placeholders[1]
    place(date_box, 1, 4.1, 11.33, 0.8)
    today = datetime.now().strftime("%d-%m-%Y")
    style_text(date_box.text_frame, today, 20, LIGHT_GREY, False, PP_ALIGN.CENTER)


def add_content_slide(prs, title, points):
    if points:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # layout 1 = title and content
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # layout 5 = title only

    # blue bar across the top of the slide
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.3))
    bar.fill.solid()
    bar.fill.fore_color.rgb = DARK_BLUE
    bar.line.fill.background()  # no outline
    bar.shadow.inherit = False  # no shadow

    place(slide.shapes.title, 0.7, 0.55, 11.9, 1.1)
    style_text(slide.shapes.title.text_frame, title, 32, DARK_BLUE, True, PP_ALIGN.LEFT)

    if points:
        body = slide.placeholders[1]
        place(body, 0.7, 1.9, 11.9, 4.9)
        write_bullets(body.text_frame, points)


def build_presentation(title, slides, filename):
    # creates the PowerPoint file and returns how many slides it has
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # widescreen (16:9)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs, title)

    for section in slides:
        groups = split_points(section["points"])
        for number, group in enumerate(groups):
            slide_title = section["title"] if number == 0 else section["title"] + " (cont.)"
            add_content_slide(prs, slide_title, group)

    prs.save(filename)
    return len(prs.slides)


# ---------- running the program ----------

def yes_no(question):
    # keeps asking until the user types y or n
    while True:
        answer = input(question).strip().lower()
        if answer == "y":
            return True
        if answer == "n":
            return False
        print("Please type y or n.")


def main():
    # take the input from the command line if given, otherwise ask for it
    if len(sys.argv) > 1:
        source = sys.argv[1]
    else:
        source = input("Enter a .txt file, a .docx file or a Google Docs link: ")

    source = source.strip().strip('"')  # removes the quotes Windows adds when you copy a file path

    if not source:
        print("Nothing entered.")
        return

    try:
        title, slides = read_source(source)
    except FileNotFoundError:
        print(f"File not found: {source}")
        return
    except ValueError as error:
        print(error)
        return
    except requests.exceptions.RequestException:
        print("Could not reach Google. Check your internet connection and try again.")
        return

    if not slides:
        print("No text was found to put on slides.")
        return

    answer = input(f"Presentation title (press Enter to use '{title}'): ").strip()
    if answer:
        title = answer

    filename = make_filename(title)
    try:
        count = build_presentation(title, slides, filename)
    except PermissionError:
        print(f"Could not save {filename}. If it is open in PowerPoint, close it and try again.")
        return

    print(f"\nDone! Created {filename} with {count} slides.")

    # os.startfile only exists on Windows
    if hasattr(os, "startfile") and yes_no("Open it now? (y/n): "):
        os.startfile(filename)


if __name__ == "__main__":
    main()
