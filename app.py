import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO
from PIL import Image

st.set_page_config(page_title="PowerPoint Merger", layout="centered")

st.title("📊 PowerPoint Merger")

def resize_image_to_1920x1080(image_bytes):
    """Resize image to 1920x1080 pixels"""
    try:
        # Open image from bytes
        img = Image.open(BytesIO(image_bytes))
        # Resize to 1920x1080 (use LANCZOS resampling for quality)
        try:
            # Try newer API first
            resized_img = img.resize((1920, 1080), Image.Resampling.LANCZOS)
        except AttributeError:
            # Fallback for older Pillow versions
            resized_img = img.resize((1920, 1080), Image.LANCZOS)
        # Convert to RGB if necessary (for JPEG compatibility)
        if resized_img.mode != 'RGB':
            resized_img = resized_img.convert('RGB')
        # Save to bytes
        output = BytesIO()
        resized_img.save(output, format='PNG')
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        # If resizing fails, return original
        return image_bytes

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

def is_all_caps(text):
    """Check if text is all uppercase (all caps)"""
    return text.isupper() and any(c.isalpha() for c in text)

def create_formatted_slide(target_presentation, text, is_title, title_color, verse_color, title_font_size, verse_font_size, title_font, verse_font, background_image=None):
    """Create a new slide with formatted text"""
    # Use blank layout
    blank_slide_layout = target_presentation.slide_layouts[6]
    slide = target_presentation.slides.add_slide(blank_slide_layout)
    
    # Get slide dimensions
    slide_width = target_presentation.slide_width
    slide_height = target_presentation.slide_height
    
    # Add background image if provided (add it first so text appears on top)
    if background_image:
        try:
            # Add image to cover entire slide
            slide.shapes.add_picture(
                BytesIO(background_image),
                0,  # left
                0,  # top
                slide_width,  # width
                slide_height  # height
            )
        except Exception:
            pass
    
    # Set black background (if no image, or as fallback)
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)  # Black
    
    # Create text box stretching from left to right edge
    width = slide_width  # Full width of slide
    height = Inches(7)  # Keep reasonable height
    left = 0  # Start from leftmost edge
    top = (slide_height - height) / 2  # Center vertically
    
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE  # Middle vertical alignment
    text_frame.margin_left = Inches(0.5)  # Small padding from left edge
    text_frame.margin_right = Inches(0.5)  # Small padding from right edge
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
    
    # Check if text is all uppercase (all caps)
    is_all_caps = text.isupper() and any(c.isalpha() for c in text)
    
    # Set font properties: use title settings if it's marked as title OR if it's all caps
    if is_title or is_all_caps:
        font.name = title_font  # Use selected title font
        font.size = Pt(title_font_size)  # Use selected title font size
        font.color.rgb = RGBColor(*title_color)  # Use selected title color
        font.bold = True  # Bold for titles
    else:
        font.name = verse_font  # Use selected verse font
        font.size = Pt(verse_font_size)  # Use selected verse font size
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
if 'title_font_size' not in st.session_state:
    st.session_state.title_font_size = 72  # Default 72pt
if 'verse_font_size' not in st.session_state:
    st.session_state.verse_font_size = 65  # Default 65pt
if 'title_font' not in st.session_state:
    st.session_state.title_font = 'Arial'  # Default Arial
if 'verse_font' not in st.session_state:
    st.session_state.verse_font = 'Arial'  # Default Arial
if 'background_image' not in st.session_state:
    st.session_state.background_image = None

# Common fonts compatible across all systems
COMMON_FONTS = [
    'Arial',
    'Times New Roman',
    'Calibri',
    'Verdana',
    'Georgia',
    'Tahoma',
    'Trebuchet MS',
    'Courier New',
    'Comic Sans MS',
    'Impact'
]

# Color and Font Settings
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

st.subheader("Font Size Settings")
font_col1, font_col2 = st.columns(2)

with font_col1:
    title_font_size = st.number_input("Title Font Size (pt)", min_value=10, max_value=200, value=st.session_state.title_font_size, step=1, key="title_font_size_input")
    st.session_state.title_font_size = int(title_font_size)

with font_col2:
    verse_font_size = st.number_input("Verse Font Size (pt)", min_value=10, max_value=200, value=st.session_state.verse_font_size, step=1, key="verse_font_size_input")
    st.session_state.verse_font_size = int(verse_font_size)

st.subheader("Font Family Settings")
font_family_col1, font_family_col2 = st.columns(2)

with font_family_col1:
    title_font = st.selectbox("Title Font", COMMON_FONTS, index=COMMON_FONTS.index(st.session_state.title_font) if st.session_state.title_font in COMMON_FONTS else 0, key="title_font_select")
    st.session_state.title_font = title_font

with font_family_col2:
    verse_font = st.selectbox("Verse Font", COMMON_FONTS, index=COMMON_FONTS.index(st.session_state.verse_font) if st.session_state.verse_font in COMMON_FONTS else 0, key="verse_font_select")
    st.session_state.verse_font = verse_font

st.markdown("## Upload PowerPoint Files")
uploaded_files = st.file_uploader(
    "Upload PowerPoint files",
    type=["pptx"],
    accept_multiple_files=True
)

st.subheader("Background Image (Optional)")
background_image_file = st.file_uploader(
    "Upload background image (will be applied to all slides)",
    type=["png", "jpg", "jpeg", "gif", "bmp"],
    key="background_image_uploader"
)

if background_image_file:
    # Read image data
    image_data = background_image_file.read()
    # Resize to 1920x1080
    resized_image_data = resize_image_to_1920x1080(image_data)
    st.session_state.background_image = resized_image_data
    st.success(f"✅ Background image loaded and resized to 1920x1080: {background_image_file.name}")
    # Display image preview
    background_image_file.seek(0)  # Reset file pointer
    st.image(background_image_file, width=300)
    if st.button("Remove Background Image", key="remove_bg_image"):
        st.session_state.background_image = None
        st.rerun()
else:
    st.session_state.background_image = None

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
            first_slide_found = False  # Track if we've found the first slide with text in this file
            first_all_caps_found = False  # Track if we've found the first all-caps slide in this file
            
            for slide in prs.slides:
                # Extract text from slide
                text = extract_text_from_slide(slide)
                
                if text:  # Only create slide if there's text
                    # Check if text is all caps
                    is_all_caps_text = is_all_caps(text)
                    
                    # Determine if it's a title slide:
                    # - First slide of each PowerPoint file (regardless of caps)
                    # - OR first all-caps slide in each PowerPoint file
                    is_title = False
                    if not first_slide_found:
                        # First slide with text in this file is always a title
                        is_title = True
                        first_slide_found = True
                    elif is_all_caps_text and not first_all_caps_found:
                        # First all-caps slide in this file is also a title
                        is_title = True
                        first_all_caps_found = True
                    
                    # Create formatted slide
                    create_formatted_slide(merged_presentation, text, is_title, st.session_state.title_color, st.session_state.verse_color, st.session_state.title_font_size, st.session_state.verse_font_size, st.session_state.title_font, st.session_state.verse_font, st.session_state.background_image)
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
