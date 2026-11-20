import os
import re
from datetime import datetime
from io import BytesIO

import requests
from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.xmlchemy import OxmlElement

WIDTH = 13.333    # slide width in inches (widescreen 16:9)
HEIGHT = 7.5      # slide height in inches
MAX_POINTS = 7    # more main points than this go on a "(cont.)" slide
MAX_LETTERS = 400 # more letters than this on one slide go on a "(cont.)" slide

# the designs: (name, dark colour, accent colour)
THEMES = [("Business", "1F3864", "2E75B6"),
          ("Finance", "14213D", "B8860B"),
          ("Education", "2D3A8C", "F29F05"),
          ("Health", "0B5D5E", "3BA99C"),
          ("Technology", "1E2A38", "00A3C4"),
          ("Nature", "1E4D2B", "6AA84F"),
          ("Creative", "4A1D6B", "E4572E")]

# words that point to each design
KEYWORDS = {
    "Business": ["business", "company", "meeting", "project", "strategy", "management", "sales",
                 "customer", "team", "plan", "report", "organisation", "organization"],
    "Finance": ["budget", "financ", "revenue", "profit", "bank", "invest", "audit", "income",
                "accounting", "expenditure", "loan", "naira", "dollar", "economy", "economic"],
    "Education": ["student", "school", "university", "lecture", "course", "learning", "teacher",
                  "curriculum", "education", "exam", "research", "academic", "lesson"],
    "Health": ["health", "patient", "medical", "hospital", "clinic", "disease", "nurse",
               "doctor", "treatment", "medicine", "wellness", "nutrition"],
    "Technology": ["software", "technology", "data", "system", "network", "computer", "digital",
                   "internet", "algorithm", "programming", "database", "cyber", "website", "app"],
    "Nature": ["environment", "climate", "farm", "agricultur", "plant", "energy", "forest",
               "wildlife", "sustainab", "water", "pollution", "solar", "crop"],
    "Creative": ["marketing", "brand", "design", "creative", "campaign", "event", "media",
                 "advert", "social", "art", "fashion", "music", "festival"],
}


# ---------- small helpers ----------

def rgb(hex_code):
    # turns a colour like "1F3864" into a colour PowerPoint understands
    return RGBColor.from_string(hex_code)


def clean(line):
    # removes bullet symbols (- * bullet #) and extra spaces from the start of a line
    return line.strip().lstrip("-*•#").strip()


def remove_bad_characters(text):
    # removes hidden control characters, which are not allowed inside a PowerPoint file
    # (the pattern means: characters with codes 0-8, 11, 12 and 14-31)
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


# ---------- reading the input ----------
# every reader makes a list of slides like this:
# {"title": "Hardware", "parent": "Budget", "section": False, "points": [("Laptops", 0), ("Dell", 1)]}
# "parent" is the topic a subtopic belongs to
# "section" is True for a topic that only introduces its subtopics
# each point is (text, level): level 0 is a main point, level 1 is a sub-point

def make_slides(items):
    # items are (kind, level, text): kind is "heading" or "point"
    # heading level 1 = topic, heading level 2 = subtopic, anything else becomes a point
    slides = []
    topic = None
    current = None

    for kind, level, text in items:
        if kind == "heading" and level in (1, 2):
            parent = ""
            if level == 2 and topic is not None:
                parent = topic["title"]
                if len(topic["points"]) == 0:
                    topic["section"] = True   # a topic with only subtopics becomes a divider slide
            current = {"title": text, "parent": parent, "section": False, "points": []}
            if level == 1:
                topic = current
            slides.append(current)
        else:
            if current is None:   # text that comes before the first heading
                current = {"title": "Introduction", "parent": "", "section": False, "points": []}
                topic = current
                slides.append(current)
            if kind == "heading":
                level = 0         # a smaller heading becomes a main point
            current["points"].append((text, level))

    return slides


def read_text(text):
    # "# Topic" and "## Subtopic" lines are headings
    # if there are no # lines, each block separated by a blank line is one slide
    # and the first line of the block is the title
    lines = text.splitlines()

    marked = False
    for line in lines:
        if re.match(r"#{1,6}\s+\S", line):   # one to six # symbols, a space, then some text
            marked = True

    items = []
    new_block = True
    for line in lines:
        stripped = line.strip()
        if stripped == "":
            new_block = True
            continue

        words = clean(line)
        if marked and stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))   # how many # symbols
            items.append(("heading", level, words))
        elif not marked and new_block:
            items.append(("heading", 1, words))
        else:
            level = 0
            if line.startswith("  ") or line.startswith("\t"):   # an indented line is a sub-point
                level = 1
            items.append(("point", level, words))
        new_block = False

    return make_slides(items)


def read_word(source, fallback_title):
    # reads a Word file (a file path, or a file held in memory)
    document = Document(source)
    title = document.core_properties.title
    if title == "":
        title = fallback_title

    # a paragraph with the "Title" style is the title of the document
    for paragraph in document.paragraphs:
        if paragraph.style.name == "Title" and paragraph.text.strip() != "":
            title = paragraph.text.strip()
            break

    # the body is every other paragraph that is not empty
    body = []
    for paragraph in document.paragraphs:
        if paragraph.style.name != "Title" and paragraph.text.strip() != "":
            body.append(paragraph)

    has_headings = False
    for paragraph in body:
        if paragraph.style.name.startswith("Heading"):
            has_headings = True

    if not has_headings:
        # no headings, so treat it like plain notes
        all_text = ""
        for paragraph in body:
            all_text += paragraph.text + "\n"
        return title, read_text(all_text)

    items = []
    for paragraph in body:
        style = paragraph.style.name
        if style.startswith("Heading"):
            number = re.search(r"\d+", style)   # "Heading 2" -> 2
            level = 1
            if number is not None:
                level = int(number.group())
            items.append(("heading", level, paragraph.text.strip()))
        else:
            level = 0
            if style.endswith("2") or style.endswith("3"):   # styles like "List Bullet 2"
                level = 1
            items.append(("point", level, clean(paragraph.text)))

    return title, make_slides(items)


def download(link):
    # downloads any link. Google Docs links are turned into a Word download first
    match = re.search(r"/document/d/([\w-]+)", link)   # finds the document id inside a Google Docs link
    if match is not None:
        link = "https://docs.google.com/document/d/" + match.group(1) + "/export?format=docx"

    response = requests.get(link, timeout=15)
    if response.status_code != 200:
        raise ValueError("Could not download that link. If it is a Google Doc, share it as 'Anyone with the link'.")

    # a real Word file starts with "PK", anything else means Google sent a sign-in page
    if match is not None and not response.content.startswith(b"PK"):
        raise ValueError("Google asked for a sign-in. Share the doc as 'Anyone with the link' and try again.")

    return response.content


def read_source(source):
    # source is a link or a file path; returns (title, slides)
    lower = source.lower()
    name = os.path.splitext(os.path.basename(source))[0]   # file name without .docx or .txt
    name = name.replace("_", " ").title()                  # "meeting_notes" -> "Meeting Notes"

    if lower.startswith("http"):
        data = download(source)
        if data.startswith(b"PK"):                         # a real Word file
            return read_word(BytesIO(data), "Presentation")
        text = data.decode("utf-8-sig", errors="replace")
        if "<html" in text[:1000].lower():
            raise ValueError("That link is a web page. Use a Google Doc link or a link to a .docx or .txt file.")
        return "Presentation", read_text(text)

    if lower.endswith(".docx"):
        return read_word(source, name)

    if lower.endswith(".txt"):
        with open(source, encoding="utf-8-sig", errors="replace") as f:
            return name, read_text(f.read())

    raise ValueError("Unsupported file. Use a .txt or .docx file, or a Google Docs link.")


def ask_for_link():
    # asks the user to paste a link; returns the link, or None if there is no valid link
    link = input("Paste the link (Enter to quit): ").strip().strip('"')
    if link.lower().startswith("http"):
        return link
    if link != "":
        print("That doesn't look like a link. It should start with https://")
    return None


def choose_source(folder):
    # lists the files as A, B, C... with "Paste a link" as the last option
    files = []
    for name in sorted(os.listdir(folder)):
        lower = name.lower()
        if lower.endswith((".txt", ".docx")):
            if not name.startswith("~$") and lower != "requirements.txt":
                files.append(name)

    print("\nLooking in: " + folder + "\n")

    # no files found, so ask for a link straight away
    if len(files) == 0:
        print("No .docx or .txt files were found in this folder.")
        print("To use a file, save it in the folder above and run the program again.")
        print("Or paste a Google Docs link below (it must be shared as 'Anyone with the link').\n")
        return ask_for_link()

    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    files = files[:25]                  # so the letters never run out
    options = files + ["Paste a link"]

    for i in range(len(options)):
        print("  " + letters[i] + ". " + options[i])

    answer = input("\nSelect an option (Enter to quit): ").strip().upper()
    if len(answer) != 1 or answer not in letters[:len(options)]:
        return None

    if answer == letters[len(options) - 1]:     # the last option: paste a link
        return ask_for_link()

    return os.path.join(folder, files[letters.index(answer)])


# ---------- building the presentation ----------

def box(slide, left, top, width, height, colour):
    # a plain coloured rectangle with no outline
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(colour)
    shape.line.fill.background()
    return shape


def write(textbox, text, size, colour, bold=False, align=PP_ALIGN.LEFT):
    # writes one piece of text into a text box and formats it
    frame = textbox.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = remove_bad_characters(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(colour)


def add_bullet(paragraph, level, colour):
    # turns a paragraph into a real PowerPoint bullet with a coloured bullet symbol
    properties = paragraph._p.get_or_add_pPr()
    properties.set("marL", str(342900 * (level + 1)))
    properties.set("indent", "-285750")

    colour_tag = OxmlElement("a:buClr")        # the colour of the bullet symbol
    colour_value = OxmlElement("a:srgbClr")
    colour_value.set("val", colour)
    colour_tag.append(colour_value)
    properties.append(colour_tag)

    bullet = OxmlElement("a:buChar")           # the bullet symbol itself
    if level == 0:
        bullet.set("char", "•")
    else:
        bullet.set("char", "–")
    properties.append(bullet)


def add_run(paragraph, text, size, colour, bold):
    # adds a piece of formatted text to a paragraph
    run = paragraph.add_run()
    run.text = remove_bad_characters(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(colour)


def add_slide(prs, title, parent, points, section, theme):
    dark = theme[1]
    accent = theme[2]
    slide = prs.slides.add_slide(prs.slide_layouts[6])   # layout 6 = blank, so we place everything ourselves
    number = len(prs.slides)                             # the slide number

    if section:
        # a divider slide: dark background, a bar and a big title in the middle
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(dark)
        box(slide, 1, 3.0, 0.12, 1.5, accent)
        title_box = slide.shapes.add_textbox(Inches(1.4), Inches(2.8), Inches(WIDTH - 2.4), Inches(1.9))
        title_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        write(title_box, title, 44, "FFFFFF", True)
        return

    # light background and a dark header band that holds the title
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb("F7F8FA")
    box(slide, 0, 0, WIDTH, 1.5, dark)
    box(slide, 0, 1.5, WIDTH, 0.07, accent)

    title_size = 34
    if len(title) > 45:
        title_size = 26   # long titles need smaller letters

    if parent != "":
        # for a subtopic, the topic it belongs to is shown small above the title
        parent_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.15), Inches(WIDTH - 1.4), Inches(0.4))
        write(parent_box, parent[:70], 14, accent)
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.55), Inches(WIDTH - 1.4), Inches(0.85))
    else:
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.2), Inches(WIDTH - 1.4), Inches(1.1))
    title_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    write(title_box, title, title_size, "FFFFFF", True)

    # the slide number, bottom right
    number_box = slide.shapes.add_textbox(Inches(WIDTH - 1.5), Inches(6.9), Inches(0.9), Inches(0.4))
    write(number_box, str(number), 12, "808080", False, PP_ALIGN.RIGHT)

    if len(points) > 0:
        # slides with little text get bigger letters, slides with a lot get smaller letters
        total_letters = 0
        for text, level in points:
            total_letters += len(text)

        if total_letters > 400:
            size = 22
        elif total_letters > 300:
            size = 24
        elif total_letters > 150:
            size = 28
        else:
            size = 32

        box(slide, 0.7, 2.1, 0.08, 4.4, accent)   # a bar to the left of the text
        body = slide.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(WIDTH - 2.0), Inches(4.8))
        frame = body.text_frame
        frame.word_wrap = True
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE   # text sits in the middle, so there is no big empty space

        first_point = True
        for text, level in points:
            if first_point:
                paragraph = frame.paragraphs[0]   # a new text box already has one empty paragraph
                first_point = False
            else:
                paragraph = frame.add_paragraph()
            paragraph.space_after = Pt(14)
            add_bullet(paragraph, level, accent)

            letters_size = size - 4 * level   # sub-points are 4 points smaller
            position = text.find(": ")
            if 0 < position <= 40:
                # a point like "Budget: 40%" gets a bold label
                add_run(paragraph, text[:position + 1], letters_size, dark, True)
                add_run(paragraph, text[position + 1:], letters_size, "333333", False)
            else:
                add_run(paragraph, text, letters_size, "333333", False)


def split_points(points):
    # splits a long list of points into groups, one group per slide
    # a new group starts when there are too many points or too many letters
    # sub-points always stay with their main point
    groups = [[]]
    letters = 0
    for text, level in points:
        if level == 0 and len(groups[-1]) > 0:
            if len(groups[-1]) >= MAX_POINTS or letters + len(text) > MAX_LETTERS:
                groups.append([])
                letters = 0
        groups[-1].append((text, level))
        letters += len(text)
    return groups


def build(title, slides, theme, path):
    # creates the PowerPoint file and returns how many slides it has
    dark = theme[1]
    accent = theme[2]

    prs = Presentation()
    prs.slide_width = Inches(WIDTH)
    prs.slide_height = Inches(HEIGHT)

    # the title slide
    first = prs.slides.add_slide(prs.slide_layouts[6])
    first.background.fill.solid()
    first.background.fill.fore_color.rgb = rgb(dark)
    box(first, 0, 7.2, WIDTH, 0.3, accent)

    if len(title) <= 40:
        title_size = 48
    else:
        title_size = 36   # long titles need smaller letters
    title_box = first.shapes.add_textbox(Inches(1), Inches(1.9), Inches(WIDTH - 2), Inches(1.9))
    title_box.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    write(title_box, title, title_size, "FFFFFF", True, PP_ALIGN.CENTER)

    box(first, WIDTH / 2 - 0.75, 4.0, 1.5, 0.07, accent)   # a short line under the title

    today = datetime.now().strftime("%d-%m-%Y")
    date_box = first.shapes.add_textbox(Inches(1), Inches(4.3), Inches(WIDTH - 2), Inches(0.8))
    write(date_box, today, 20, "FFFFFF", False, PP_ALIGN.CENTER)

    # an agenda slide that lists the topics (only when there are between 3 and 8 topics)
    topics = []
    for s in slides:
        if s["parent"] == "" and s["title"] != "Introduction":
            topics.append((s["title"], 0))
    if len(topics) >= 3 and len(topics) <= 8:
        add_slide(prs, "Agenda", "", topics, False, theme)

    # the other slides
    for s in slides:
        if s["section"]:
            add_slide(prs, s["title"], "", [], True, theme)
            continue

        number = 0
        for group in split_points(s["points"]):
            slide_title = s["title"]
            if number > 0:
                slide_title = slide_title + " (cont.)"
            add_slide(prs, slide_title, s["parent"], group, False, theme)
            number += 1

    # the closing slide
    add_slide(prs, "Thank you", "", [], True, theme)

    prs.save(path)
    return len(prs.slides)


def yes_no(question):
    # keeps asking until the user types y or n
    while True:
        answer = input(question).strip().lower()
        if answer == "y":
            return True
        if answer == "n":
            return False
        print("Please type y or n.")


def tidy_slides(slides):
    # cuts long paragraphs into short bullet points and tidies capital letters and full stops
    # the topics and subtopics stay exactly as they were
    for s in slides:
        new_points = []
        for text, level in s["points"]:
            sentences = [text]
            if len(text) > 120:   # only long points are cut, so short ones like "Dr. Smith joins" are left alone
                sentences = re.split(r"(?<=[.!?])\s+", text)   # cut after . ! or ? followed by a space

            for sentence in sentences:
                sentence = sentence.strip().rstrip(".;,:")   # no full stop at the end of a bullet
                if sentence != "":
                    first_word = sentence.split()[0]
                    if first_word.islower():   # names like eNaira or iPhone are left alone
                        sentence = sentence[0].upper() + sentence[1:]
                    new_points.append((sentence, level))
        s["points"] = new_points


# ---------- choosing the design ----------

def detect_theme(slides):
    # put all the text of the notes into one long string
    text = ""
    for s in slides:
        text += s["title"] + " " + s["parent"] + " "
        for point_text, level in s["points"]:
            text += point_text + " "
    text = text.lower()

    # count the keywords of each design, and keep the design with the highest score
    best_theme = THEMES[0]
    best_score = 0
    for theme in THEMES:
        score = 0
        for word in KEYWORDS[theme[0]]:
            score += len(re.findall(r"\b" + word, text))   # \b means the word must start at the beginning of a word
        if score > best_score:
            best_theme = theme
            best_score = score
    return best_theme


def choose_theme(slides):
    picked = detect_theme(slides)

    print("\nDesigns:")
    number = 1
    for theme in THEMES:
        note = ""
        if theme is picked:
            note = "   <- matches your notes"
        print("  " + str(number) + ". " + theme[0] + note)
        number += 1

    while True:
        answer = input("Press Enter to use the matching design, or type a number to change it: ").strip()
        if answer == "":
            return picked
        if answer.isdigit() and 1 <= int(answer) <= len(THEMES):
            return THEMES[int(answer) - 1]
        print("Please press Enter or type one of the numbers shown.")


# ---------- running the program ----------

def estimate_slides(slides):
    # works out roughly how many slides the presentation will have
    estimate = 2   # the title slide and the closing slide
    for s in slides:
        estimate += len(split_points(s["points"]))
    return estimate


def main():
    # the program looks for files in the same folder as this script
    folder = os.path.dirname(os.path.abspath(__file__))

    source = choose_source(folder)
    if source is None:
        print("Nothing chosen.")
        return

    try:
        title, slides = read_source(source)
    except (ValueError, requests.exceptions.RequestException, OSError) as error:
        print("Could not read that: " + str(error))
        return
    except Exception:
        print("Could not read that file. Make sure it is a real .docx file (not an old .doc).")
        return

    if len(slides) == 0:
        print("No text was found to put on slides.")
        return

    print("\nI can structure the text for slides: cut long paragraphs into short bullet points")
    print("and tidy capital letters and full stops. Your topics and subtopics stay as they are.")
    print("Say n if you already have your own pattern.")
    if yes_no("Structure the text? (y/n): "):
        tidy_slides(slides)

    theme = choose_theme(slides)

    answer = input("Presentation title (Enter to use '" + title + "'): ").strip()
    if answer != "":
        title = answer

    # make a safe file name, e.g. "Meeting Notes" -> "Meeting_Notes.pptx"
    safe = re.sub(r"[^\w\- ]", "", title)   # remove symbols that are not allowed in file names
    safe = safe.strip().replace(" ", "_")
    if safe == "":
        safe = "presentation"

    path = os.path.join(folder, safe + ".pptx")
    number = 2
    while os.path.exists(path):             # never overwrite an existing file
        path = os.path.join(folder, safe + " (" + str(number) + ").pptx")
        number += 1

    estimate = estimate_slides(slides)
    if estimate > 30:
        print("\nThis will make about " + str(estimate) + " slides. A large presentation may open slowly in PowerPoint.")
        if input("Continue? (y/n): ").strip().lower() != "y":
            return

    try:
        count = build(title, slides, theme, path)
    except PermissionError:
        print("Could not save. Close the file if it is open in PowerPoint, then try again.")
        return

    print("\nDone! Created " + os.path.basename(path) + " with " + str(count) + " slides.")
    if count > 100:
        print("Note: " + str(count) + " slides is a lot. A presentation this big may open slowly in PowerPoint.")

    if hasattr(os, "startfile"):            # os.startfile only exists on Windows
        os.startfile(path)


if __name__ == "__main__":
    main()