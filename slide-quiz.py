from pptx import Presentation
from pptx.util import Inches, Pt
import openai

# Set your OpenAI API key
openai.api_key = "YOUR API KEY"

def generate_quiz_question(bullet_point):
    prompt = (
        f"Generate a short-answer quiz question based solely on the following key point from a biology lecture. "
        f"Ensure the question is specific to the given bullet point and can be answered using only the provided information. "
        f"Do not include an answer: '{bullet_point}'."
    )
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )
    return response["choices"][0]["message"]["content"].strip()

# Load the original PowerPoint file
ppt_path = "bioslides.pptx"
presentation = Presentation(ppt_path)

def insert_quiz_slide(prs, slide_index, bullet_points):
    """
    Inserts a quiz slide BEFORE the lecture slide at slide_index.
    The quiz slide includes:
      - A title: "Quiz Slide {slide_index+1}"
      - A textbox with generated questions (only if bullet_points is not empty)
    """
    # Create a new slide using a blank layout (you can adjust the layout index as needed)
    quiz_slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Add a title textbox manually (since a blank layout may not have a title placeholder)
    title_box = quiz_slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = f"Quiz Slide {slide_index + 1}"
    
    # Only add a question textbox if there are bullet points
    if bullet_points:
        left, top, width, height = Inches(1), Inches(1.5), Inches(8.5), Inches(5)
        quiz_textbox = quiz_slide.shapes.add_textbox(left, top, width, height)
        quiz_text_frame = quiz_textbox.text_frame
        quiz_text_frame.word_wrap = True  # Enable text wrapping

        for bullet_point in bullet_points:
            quiz_question = generate_quiz_question(bullet_point)
            p = quiz_text_frame.add_paragraph()
            p.text = f"• {quiz_question}"
            p.space_after = Pt(10)
    
    # Rearrange the slide order so the quiz slide comes before the lecture slide.
    # (Accessing the underlying XML list of slides.)
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    quiz_slide_element = slides[-1]  # Our newly added quiz slide is the last element
    xml_slides.remove(quiz_slide_element)
    xml_slides.insert(slide_index, quiz_slide_element)

# First, store the number of original slides before inserting any quiz slides.
original_slide_count = len(presentation.slides)

# Iterate backwards over the original slides so that inserting new slides doesn't shift the indices of earlier slides.
for i in reversed(range(original_slide_count)):
    slide = presentation.slides[i]
    # Extract bullet points from the slide
    bullet_points = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text.strip()
                if text:
                    bullet_points.append(text)
    # Insert a quiz slide immediately BEFORE this lecture slide.
    insert_quiz_slide(presentation, i, bullet_points)

# Save the updated presentation (you can overwrite the original or save to a new file)
presentation.save("bioslides_updated.pptx")
