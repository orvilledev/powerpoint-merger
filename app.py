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

def create_formatted_slide(target_presentation, text, is_title, title_color, verse_color):
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
    
    # Check if text is all uppercase (all caps)
    is_all_caps = text.isupper() and any(c.isalpha() for c in text)
    
    # Set font size: 72pt for titles (all caps), 65pt for verses
    if is_all_caps:
        font.size = Pt(72)  # Font size 72 for titles
        font.color.rgb = RGBColor(*title_color)  # Use selected title color
        font.bold = True  # Bold for all caps
    else:
        font.size = Pt(65)  # Font size 65 for verses
        font.color.rgb = RGBColor(*verse_color)  # Use selected verse color
        font.bold = False
    
    return slide

# Initialize session state for file ordering
if 'file_order' not in st.session_state:
    st.session_state.file_order = []
if 'uploaded_files_dict' not in st.session_state:
    st.session_state.uploaded_files_dict = {}
if 'title_color' not in st.session_state:
    st.session_state.title_color = [255, 255, 0]  # Default yellow
if 'verse_color' not in st.session_state:
    st.session_state.verse_color = [255, 255, 255]  # Default white

# Color selection
st.subheader("Color Settings")
col1, col2 = st.columns(2)

with col1:
    # Convert RGB to hex for color picker
    title_hex = f"#{st.session_state.title_color[0]:02x}{st.session_state.title_color[1]:02x}{st.session_state.title_color[2]:02x}".upper()
    title_color = st.color_picker("Title Color (All Caps)", title_hex, key="title_color_picker")
    # Convert hex to RGB
    title_color_rgb = tuple(int(title_color[i:i+2], 16) for i in (1, 3, 5))
    st.session_state.title_color = list(title_color_rgb)

with col2:
    # Convert RGB to hex for color picker
    verse_hex = f"#{st.session_state.verse_color[0]:02x}{st.session_state.verse_color[1]:02x}{st.session_state.verse_color[2]:02x}".upper()
    verse_color = st.color_picker("Verse Color", verse_hex, key="verse_color_picker")
    # Convert hex to RGB
    verse_color_rgb = tuple(int(verse_color[i:i+2], 16) for i in (1, 3, 5))
    st.session_state.verse_color = list(verse_color_rgb)

uploaded_files = st.file_uploader(
    "Upload PowerPoint files",
    type=["pptx"],
    accept_multiple_files=True
)

# Update session state when new files are uploaded
if uploaded_files:
    # Create a dictionary to track files by name
    new_files_dict = {file.name: file for file in uploaded_files}
    
    # Add new files to the order (if not already present)
    for file in uploaded_files:
        if file.name not in st.session_state.file_order:
            st.session_state.file_order.append(file.name)
    
    # Remove files that are no longer in the upload
    st.session_state.file_order = [name for name in st.session_state.file_order if name in new_files_dict]
    
    # Update the files dictionary
    st.session_state.uploaded_files_dict = new_files_dict

# Display file reordering interface
if st.session_state.file_order:
    st.subheader("Arrange Files (use buttons to reorder)")
    
    # Create a list to store new order
    new_order = st.session_state.file_order.copy()
    
    # Display files with reordering controls
    for i, file_name in enumerate(st.session_state.file_order):
        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
        
        with col1:
            st.write(f"**{i+1}.** {file_name}")
        
        with col2:
            if st.button("⬆️ Up", key=f"up_{i}", disabled=(i == 0)):
                if i > 0:
                    new_order[i], new_order[i-1] = new_order[i-1], new_order[i]
                    st.session_state.file_order = new_order
                    st.rerun()
        
        with col3:
            if st.button("⬇️ Down", key=f"down_{i}", disabled=(i == len(st.session_state.file_order) - 1)):
                if i < len(new_order) - 1:
                    new_order[i], new_order[i+1] = new_order[i+1], new_order[i]
                    st.session_state.file_order = new_order
                    st.rerun()
        
        with col4:
            if st.button("🗑️ Remove", key=f"remove_{i}"):
                new_order.remove(file_name)
                st.session_state.file_order = new_order
                if file_name in st.session_state.uploaded_files_dict:
                    del st.session_state.uploaded_files_dict[file_name]
                st.rerun()

# Get ordered list of files
ordered_files = [st.session_state.uploaded_files_dict[name] for name in st.session_state.file_order if name in st.session_state.uploaded_files_dict]

if ordered_files and st.button("Merge PowerPoints"):
    try:
        merged_presentation = Presentation()
        
        # Set slide dimensions to match first presentation if available
        if ordered_files:
            first_prs = Presentation(ordered_files[0])
            merged_presentation.slide_width = first_prs.slide_width
            merged_presentation.slide_height = first_prs.slide_height
            ordered_files[0].seek(0)  # Reset file pointer
        
        # Remove default empty slide
        if merged_presentation.slides:
            merged_presentation.slides.remove(merged_presentation.slides[0])
        
        slide_index = 0
        
        for ppt_file in ordered_files:
            prs = Presentation(ppt_file)
            
            for slide in prs.slides:
                # Extract text from slide
                text = extract_text_from_slide(slide)
                
                if text:  # Only create slide if there's text
                    # Determine if it's a title slide
                    is_title = is_title_slide(text, slide_index)
                    
                    # Create formatted slide
                    create_formatted_slide(merged_presentation, text, is_title, st.session_state.title_color, st.session_state.verse_color)
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
