import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO

st.set_page_config(page_title="PowerPoint Merger", layout="centered")

st.title("📊 PowerPoint Merger")

def extract_text_from_slide(slide):
    """Extract all text from a slide"""
    text_content = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                para_text = ""
                for run in paragraph.runs:
                    para_text += run.text
                if para_text.strip():
                    text_content.append(para_text.strip())
    return "\n".join(text_content)

def is_title_slide(text, slide_index):
    """Determine if a slide is a title slide (first slide or short text)"""
    # First slide is always a title
    if slide_index == 0:
        return True
    # Short text (less than 100 chars) is likely a title
    if len(text) < 100:
        return True
    return False

def create_formatted_slide(target_presentation, text, is_title):
    """Create a new slide with formatted text"""
    # Use blank layout
    blank_slide_layout = target_presentation.slide_layouts[6]
    slide = target_presentation.slides.add_slide(blank_slide_layout)
    
    # Set black background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)  # Black
    
    # Get slide dimensions
    slide_width = target_presentation.slide_width
    slide_height = target_presentation.slide_height
    
    # Create text box centered on slide
    width = Inches(10)
    height = Inches(7)
    left = (slide_width - width) / 2  # Center horizontally
    top = (slide_height - height) / 2  # Center vertically
    
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE  # Middle vertical alignment
    text_frame.margin_left = Inches(0.5)
    text_frame.margin_right = Inches(0.5)
    text_frame.margin_top = Inches(0.5)
    text_frame.margin_bottom = Inches(0.5)
    
    # Clear default paragraph
    text_frame.clear()
    
    # Add text with formatting
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER  # Center alignment
    paragraph.space_after = Pt(0)
    
    run = paragraph.add_run()
    run.text = text  # Preserve original case
    font = run.font
    font.name = 'Arial'
    font.size = Pt(72)  # Font size 72
    
    # Check if text is all uppercase (all caps)
    is_all_caps = text.isupper() and any(c.isalpha() for c in text)
    
    # Set text color: yellow for all caps, white for others
    if is_all_caps:
        font.color.rgb = RGBColor(255, 255, 0)  # Yellow
        font.bold = True  # Bold for all caps
    else:
        font.color.rgb = RGBColor(255, 255, 255)  # White
        font.bold = False
    
    return slide

uploaded_files = st.file_uploader(
    "Upload PowerPoint files",
    type=["pptx"],
    accept_multiple_files=True
)

if uploaded_files and st.button("Merge PowerPoints"):
    try:
        merged_presentation = Presentation()
        
        # Set slide dimensions to match first presentation if available
        if uploaded_files:
            first_prs = Presentation(uploaded_files[0])
            merged_presentation.slide_width = first_prs.slide_width
            merged_presentation.slide_height = first_prs.slide_height
            uploaded_files[0].seek(0)  # Reset file pointer
        
        # Remove default empty slide
        if merged_presentation.slides:
            merged_presentation.slides.remove(merged_presentation.slides[0])
        
        slide_index = 0
        
        for ppt_file in uploaded_files:
            prs = Presentation(ppt_file)
            
            for slide in prs.slides:
                # Extract text from slide
                text = extract_text_from_slide(slide)
                
                if text:  # Only create slide if there's text
                    # Determine if it's a title slide
                    is_title = is_title_slide(text, slide_index)
                    
                    # Create formatted slide
                    create_formatted_slide(merged_presentation, text, is_title)
                    slide_index += 1
        
        output = BytesIO()
        merged_presentation.save(output)
        output.seek(0)
        
        st.success("✅ PowerPoints merged successfully!")
        
        st.download_button(
            label="⬇️ Download merged PowerPoint",
            data=output,
            file_name="merged_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        st.error(f"Error merging presentations: {str(e)}")
        st.exception(e)
