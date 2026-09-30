import streamlit as st
from google import genai
from google.genai import types
import json
import pymupdf4llm
import pypandoc
import os
import shutil
import re
import base64
import mimetypes
import PyPDF2
from docx import Document
from docx.shared import Mm, Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- 3. HELPER FUNCTIONS ---
def extract_text_from_pdf(pdf_file):
    # Clean up old temporary folders from PREVIOUS uploads before extracting new ones
    if os.path.exists("media"): shutil.rmtree("media")
    if os.path.exists("word"): shutil.rmtree("word")
    
    os.makedirs("media", exist_ok=True)
    with open("temp.pdf", "wb") as f:
        f.write(pdf_file.read())
        
    text = pymupdf4llm.to_markdown("temp.pdf", write_images=True, image_path="media")
    os.remove("temp.pdf")
    return text

def extract_text_from_docx(docx_file):
    # Clean up old temporary folders from PREVIOUS uploads before extracting new ones
    if os.path.exists("media"): shutil.rmtree("media")
    if os.path.exists("word"): shutil.rmtree("word")
    
    with open("temp.docx", "wb") as f:
        f.write(docx_file.read())
        
    text = pypandoc.convert_file("temp.docx", "markdown", extra_args=["--extract-media=."])
    os.remove("temp.docx")
    return text

def clean_math_backslashes(obj):
    """Recursively cleans up over-escaped double backslashes back to standard LaTeX single backslashes."""
    if isinstance(obj, str):
        return obj.replace('\\\\', '\\')
    elif isinstance(obj, list):
        return [clean_math_backslashes(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: clean_math_backslashes(v) for k, v in obj.items()}
    return obj

def format_option_labels(parsed_data):
    """Ensures options follow 'A. Title' format, and forces 'correct_option' to be strictly a single letter."""
    import re
    if isinstance(parsed_data, list):
        for mcq in parsed_data:
            if isinstance(mcq, dict):
                # 1. Clean and lock the 'correct_option' to a single letter
                if 'correct_option' in mcq and isinstance(mcq['correct_option'], str):
                    co_val = mcq['correct_option'].strip()
                    
                    # Try to extract A, B, C, or D from common AI outputs (e.g., "A", "A.", "Option B", "C) ")
                    match = re.match(r'^(?:Option\s*)?([A-D])(?:[\.\)\-:\s]|$)', co_val, re.IGNORECASE)
                    if match:
                        mcq['correct_option'] = match.group(1).upper()
                    else:
                        # Fallback: if the AI output the full text answer instead, find which option it matches
                        for opt in ['A', 'B', 'C', 'D']:
                            if opt in mcq and isinstance(mcq[opt], str):
                                # Strip prefixes to compare raw text
                                raw_opt_val = re.sub(rf'^{opt}[\.\)\-]\s*', '', mcq[opt], flags=re.IGNORECASE).strip()
                                if raw_opt_val == co_val or mcq[opt].strip() == co_val:
                                    mcq['correct_option'] = opt
                                    break
                                    
                # 2. Format the actual A, B, C, D option columns
                for opt in ['A', 'B', 'C', 'D']:
                    if opt in mcq and isinstance(mcq[opt], str):
                        clean_val = re.sub(rf'^{opt}[\.\)\-]\s*', '', mcq[opt], flags=re.IGNORECASE).strip()
                        mcq[opt] = f"{opt}. {clean_val}"
                        
    return parsed_data

# 1. Add api_key to the function arguments
def generate_mcqs(context_text, user_topics, custom_instructions, num_easy, num_medium, num_hard, selected_model, api_key):
    # 2. Initialize the new Client
    client = genai.Client(api_key=api_key)
    total_questions = num_easy + num_medium + num_hard
    
    if context_text.strip():
        source_instruction = "Based on the provided Context Text and user instructions, generate"
        context_block = f"Context Text:\n{context_text}"
    else:
        source_instruction = "Based on your expert general knowledge and user instructions, generate"
        context_block = "No context text provided. Generate purely based on the requested topics and instructions."
    
    prompt = f"""
    Act as an expert educator. {source_instruction} exactly {total_questions} board-standard multiple-choice questions.
    
    SPECIAL USER GENERATION INSTRUCTIONS (CRITICAL):
    {custom_instructions if custom_instructions else "None provided. Follow standard board-level question generation."}

    CRITICAL INSTRUCTION FOR MATH AND JSON ESCAPING:
    1. If the text contains mathematical equations or physics/chemistry formulas, you MUST use standard LaTeX (e.g., $E=mc^2$).
    2. You MUST double-escape all LaTeX backslashes for valid JSON (e.g., use \\frac{1}{2} instead of \frac{1}{2}, and \\sum instead of \sum).
    3. Output EXACTLY ONE continuous JSON array. Do not split the output into multiple arrays. Do not add any conversational text before or after the JSON.
    
    CRITICAL INSTRUCTION FOR IMAGES:
    If the source text contains markdown image links (e.g., ![image](media/img.png) or ![alt](word/media/image1.jpeg)), you MUST preserve them exactly as they appear. Place them in the relevant JSON field (usually "question_title" or "solution_body"). Never modify, translate, or delete the image file paths.

    CRITICAL INSTRUCTION FOR DATA TABLES:
    If the source text contains data tables (formatted in Markdown like |---|---|), you MUST convert them into basic HTML tables (e.g., <table border='1'><tr><td>...</td></tr></table>) inside the JSON string. 
    - Do NOT use Markdown tables in your output.
    - Write the ENTIRE HTML table on a SINGLE LINE (do not use \n characters inside the HTML table) so it does not break the JSON string or downstream formatting.
    
    CRITICAL INSTRUCTIONS FOR DIFFICULTY LEVEL:
    You must generate EXACTLY:
    - {num_easy} questions where "difficulty_level" is "Easy"
    - {num_medium} questions where "difficulty_level" is "Medium"
    - {num_hard} questions where "difficulty_level" is "Hard"
    Do not use any other words for difficulty level.

    CRITICAL INSTRUCTION FOR THE 'topics' COLUMN FIELD:
    Here is the list of allowed topic names: {user_topics}
    For the "topics" field in each question object, you MUST select EXACTLY ONE topic from the list above that best fits the generated question. Do not invent any new topics.
    DO NOT include sequence identifiers like '1.1', '2.1', etc. in the topic field.

    CRITICAL INSTRUCTION FOR 'solution_body':
    The "solution_body" field MUST follow this exact 3-line format:
    Line 1: The exact text of the correct option (DO NOT include sequence identifiers like 'A.', 'B.', 'Option A', '১.', etc.). Place an Enter at the end of the line.
    Line 2: Enter just a space (paragraph mark). DO NOT put any texts here.
    Line 3: A clear, detailed explanation of why this answer is correct. Always start with the line 'ব্যাখ্যা:'

    Example format for solution_body:
    "ইনপুট, প্রসেসিং, আউটপুট, মেমোরি ও কন্ট্রোল ইউনিট
    
    ব্যাখ্যা: কম্পিউটারের কাজ করার মূল পদ্ধতি হলো তথ্য গ্রহণ, প্রসেসিং, প্রদর্শন ও সংরক্ষণ করা।"

    Output the result STRICTLY as a JSON array of objects.
    Each object must have the following exact keys:
    "sl_no", "question_title", "A", "B", "C", "D", 
    "solution_body", "correct_option", "subject", "chapter", "topics", 
    "question_category", "difficulty_level".
    
    {context_block}
    """
    
    # 3. Use the new generation syntax
    response = client.models.generate_content(
        model=selected_model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        )
    )
    
    raw_output = response.text.strip()
    
    # 1. Strip rogue markdown formatting if Gemini ignored the mime_type
    if raw_output.startswith("```json"):
        raw_output = raw_output[7:-3].strip()
    elif raw_output.startswith("```"):
        raw_output = raw_output[3:-3].strip()
        
    # --- 2. The LaTeX Backslash Sanitizer ---
    # Temporarily hide valid JSON newlines and quotes
    raw_output = raw_output.replace('\\n', '<<NEWLINE>>')
    raw_output = raw_output.replace('\\"', '<<QUOTE>>')
    # Forcefully double-escape all remaining rogue LaTeX backslashes (e.g., \frac becomes \\frac)
    raw_output = raw_output.replace('\\', '\\\\')
    # Restore the valid JSON formatting
    raw_output = raw_output.replace('<<NEWLINE>>', '\\n')
    raw_output = raw_output.replace('<<QUOTE>>', '\\"')
        
    try:
        parsed_data = json.loads(raw_output)
        parsed_data = clean_math_backslashes(parsed_data)
        return format_option_labels(parsed_data) # <-- Added formatter here
    except json.JSONDecodeError as e:
        import re
        match = re.search(r'\[.*?\](?=\s*$|\s*```)', raw_output, re.DOTALL)
        if match:
            try:
                parsed_data = json.loads(match.group(0))
                parsed_data = clean_math_backslashes(parsed_data)
                return format_option_labels(parsed_data) # <-- Added formatter here
            except:
                pass
        raise ValueError(f"JSON Error: {str(e)} \n\nRAW AI OUTPUT (Debug this):\n{raw_output}")
        
# 1. Add api_key to the function arguments
def parse_existing_mcqs(raw_mcq_text, user_topics, special_instructions, selected_model, api_key):
    # 2. Initialize the new Client
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    You are an expert educational content parser and converter.
    Your task is to take the provided raw, unformatted, or existing MCQs and map them strictly into a standard JSON structure.
    
    CRITICAL PARSING RULES:
    1. Parse ALL MCQs found in the input text into a JSON array of objects.
    2. Number the sl_no sequentially starting from 1.
    3. Extract the Question Title, Option A, Option B, Option C, Option D, Correct Option, and Solution Body (if solution body isn't provided, create a brief accurate explanation. If the user says no explanation then DO NOT generate any explanation.).
    4. Infer appropriate "subject", "chapter", "question_category" (e.g., Board, Model Test), and "difficulty_level" (Easy, Medium, Hard) for each question.

    CRITICAL INSTRUCTION FOR MATH AND EQUATIONS:
    If the source text contains mathematical equations, chemical formulas, or physics expressions, you MUST preserve them using standard LaTeX format. Use $ for inline math (e.g., $E=mc^2$) and $$ for display math. Do NOT use plain text approximations.

    CRITICAL INSTRUCTION FOR IMAGES:
    If the source text contains markdown image links (e.g., ![image](media/img.png) or ![alt](word/media/image1.jpeg)), you MUST preserve them exactly as they appear. Place them in the relevant JSON field (usually "question_title" or "solution_body"). Never modify, translate, or delete the image file paths.
    
    CRITICAL INSTRUCTION FOR DATA TABLES:
    If the source text contains data tables (formatted in Markdown like |---|---|), you MUST convert them into basic HTML tables (e.g., <table border='1'><tr><td>...</td></tr></table>) inside the JSON string. 
    - Do NOT use Markdown tables in your output.
    - Write the ENTIRE HTML table on a SINGLE LINE (do not use \n characters inside the HTML table) so it does not break the JSON string or downstream formatting.
    
    CRITICAL INSTRUCTION FOR THE 'topics' FIELD:
    Here is a list of allowed topics: {user_topics}
    For the "topics" field in each question, you MUST select EXACTLY ONE topic from the list above that best fits the question. Do not invent any new topics.
    DO NOT include numbering/sequence identifiers like '1.1', '2.1',.... etc.

    CRITICAL INSTRUCTION FOR 'solution_body':
    The "solution_body" field MUST follow this exact 3-line format:
    Line 1: The exact text of the correct option (DO NOT include sequence identifiers like 'A.', 'B.', 'Option A', '১.', etc.). Place an Enter at the end of the line.
    Line 2: Enter just a space (paragraph mark). DO NOT put any texts here.
    Line 3: A clear, detailed explanation of why this answer is correct. Always start with the line 'ব্যাখ্যা:'

    Example format for solution_body:
    "ইনপুট, প্রসেসিং, আউটপুট, মেমোরি ও কন্ট্রোল ইউনিট
    
    ব্যাখ্যা: কম্পিউটারের কাজ করার মূল পদ্ধতি হলো তথ্য গ্রহণ, প্রসেসিং, প্রদর্শন ও সংরক্ষণ করা।"
    
    SPECIAL USER INSTRUCTIONS (CRITICAL):
    {special_instructions if special_instructions else "None provided. Follow standard parsing."}

    Output the result STRICTLY as a JSON array of objects. Do not include markdown formatting like ```json.
    Each object must have the following exact keys:
    "sl_no", "question_title", "A", "B", "C", "D", 
    "solution_body", "correct_option", "subject", "chapter", "topics", 
    "question_category", "difficulty_level".

    Raw MCQs to parse:
    {raw_mcq_text}
    """
    
    # 3. Use the new generation syntax
    response = client.models.generate_content(
        model=selected_model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        )
    )
    
    raw_output = response.text.strip()
    
    # 1. Strip rogue markdown formatting if Gemini ignored the mime_type
    if raw_output.startswith("```json"):
        raw_output = raw_output[7:-3].strip()
    elif raw_output.startswith("```"):
        raw_output = raw_output[3:-3].strip()
        
    try:
        parsed_data = json.loads(raw_output)
        parsed_data = clean_math_backslashes(parsed_data)
        return format_option_labels(parsed_data) # <-- Added formatter here
    except json.JSONDecodeError as e:
        import re
        match = re.search(r'\[.*?\](?=\s*$|\s*```)', raw_output, re.DOTALL)
        if match:
            try:
                parsed_data = json.loads(match.group(0))
                parsed_data = clean_math_backslashes(parsed_data)
                return format_option_labels(parsed_data) # <-- Added formatter here
            except:
                pass
        raise ValueError(f"JSON Error: {str(e)} \n\nRAW AI OUTPUT (Debug this):\n{raw_output}")
        
def process_html_images(text):
    """Converts Markdown images to HTML, locks width to 2 inches, 
    and aggressively annihilates multi-line Pandoc artifacts."""
    text_str = str(text)
    
    def replacer(match):
        img_path = match.group(1).strip()
        
        # --- THE FIX: Self-healing memory cache ---
        # If the app forgot the cache, rebuild it instantly before it crashes.
        if "image_cache" not in st.session_state:
            st.session_state["image_cache"] = {}
            
        mime_type, _ = mimetypes.guess_type(img_path)
        if not mime_type:
            mime_type = "image/png"
            
        if img_path in st.session_state["image_cache"]:
            b64 = st.session_state["image_cache"][img_path]
            return f'<img src="data:{mime_type};base64,{b64}" width="2in" />'
            
        if os.path.exists(img_path):
            with open(img_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode('utf-8')
            st.session_state["image_cache"][img_path] = b64
            return f'<img src="data:{mime_type};base64,{b64}" width="2in" />'
            
        return f'<img src="{img_path}" width="2in" />'

    # 1. Match images and their immediate dimension blocks
    cleaned = re.sub(r'!\[[^\]]*\]\(([^)]+)\)(?:\s*\{[^{}]*\})?', replacer, text_str, flags=re.DOTALL)
    
    # 2. The Multi-Line Assassin: Hunts down ANY leftover curly brace block 
    cleaned = re.sub(r'\{[^{}]*(?:width|height)[^{}]*\}', '', cleaned, flags=re.IGNORECASE | re.DOTALL)
    
    return cleaned

def safe_newline_to_br(text):
    """Replaces \n with <br> EXCEPT when the \n is inside an HTML table."""
    text_str = str(text)
    
    # If there's no table, safely replace all newlines
    if "<table" not in text_str.lower():
        return text_str.replace('\n', '<br>')
        
    # If there is a table, split by table tags and only replace \n outside of them
    parts = re.split(r'(<table.*?</table>)', text_str, flags=re.IGNORECASE | re.DOTALL)
    for i, part in enumerate(parts):
        if not part.lower().startswith("<table"):
            parts[i] = part.replace('\n', '<br>')
            
    return "".join(parts)

def format_docx_layout(docx_filename):
    """Forces A4 Landscape, Narrow Margins, and injects raw XML to guarantee Table Borders for both outer and nested tables."""
    doc = Document(docx_filename)
    
    # 1. Force Page Layout: A4, Landscape, Narrow Margins (0.5 inches)
    for section in doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Mm(297)
        section.page_height = Mm(210)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        
    # Helper function to apply raw XML borders and autofit to any given table object
    def apply_xml_borders(tbl_obj):
        tbl = tbl_obj._tbl
        tblPr = tbl.tblPr
        
        # --- A. Force Borders ---
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)
        else:
            tblBorders.clear() 
        
        # Draw solid black lines on all sides and internal grids
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4') 
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000') 
            tblBorders.append(border)
            
        # --- B. Force Autofit to Window ---
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        tblW.set(qn('w:type'), 'pct') 
        tblW.set(qn('w:w'), '5000')   

    # 2. Iterate through all tables (both Main and Nested)
    for main_table in doc.tables:
        # Apply borders to the master outer table
        apply_xml_borders(main_table)
        
        # Dig into the rows and cells to find any inner tables
        for row in main_table.rows:
            for cell in row.cells:
                for nested_table in cell.tables:
                    # Apply the exact same borders to the inner table
                    apply_xml_borders(nested_table)
        
    doc.save(docx_filename)

def create_mcq_docx(mcq_data, output_filename="MCQs.docx"):
    html = "<h1>Generated MCQs</h1>\n<table border='1'>\n"
    html += "<tr><th>Sl no.</th><th>Question Title</th><th>Option A</th><th>Option B</th><th>Option C</th><th>Option D</th><th>Solution Body</th><th>Correct Option</th><th>Subject</th><th>Chapter</th><th>Topics</th><th>Question Category</th><th>Difficulty Level</th></tr>\n"
    
    for mcq in mcq_data:
        html += "<tr>"
        html += f"<td>{mcq.get('sl_no', '')}</td>"
        
        # Process images and tables in the title
        q_title = safe_newline_to_br(mcq.get('question_title', ''))
        q_title = process_html_images(q_title)
        html += f"<td>{q_title}</td>"
        
        # Process images in options
        html += f"<td>{process_html_images(mcq.get('A', ''))}</td>"
        html += f"<td>{process_html_images(mcq.get('B', ''))}</td>"
        html += f"<td>{process_html_images(mcq.get('C', ''))}</td>"
        html += f"<td>{process_html_images(mcq.get('D', ''))}</td>"
        
        # Process line breaks, tables, AND images in solution body safely
        sol_body = safe_newline_to_br(mcq.get('solution_body', ''))
        sol_body = process_html_images(sol_body)
        html += f"<td>{sol_body}</td>"
        
        html += f"<td>{mcq.get('correct_option', '')}</td>"
        html += f"<td>{mcq.get('subject', '')}</td>"
        html += f"<td>{mcq.get('chapter', '')}</td>"
        html += f"<td>{mcq.get('topics', '')}</td>"
        html += f"<td>{mcq.get('question_category', '')}</td>"
        html += f"<td>{mcq.get('difficulty_level', '')}</td>"
        html += "</tr>\n"
    
    html += "</table>"
    
    # Convert HTML + LaTeX to Word document
    pypandoc.convert_text(html, 'docx', format='html+tex_math_dollars', outputfile=output_filename)
    
    # --- NEW: Apply Page Layout and Table Borders ---
    format_docx_layout(output_filename)
    
    # NOTE: Do NOT delete media folders here! 
    # Keeping them on disk allows infinite re-downloads and interactive cell editing.

def run_mcq_interface(model_choice, api_key_input):
    # Initialize session state for API Key and Table Data
    if "user_api_key" not in st.session_state:
        st.session_state["user_api_key"] = ""
    if "mcq_data_t1" not in st.session_state:
        st.session_state["mcq_data_t1"] = None
    if "mcq_data_t2" not in st.session_state:
        st.session_state["mcq_data_t2"] = None
    
    # Create two clean tabs
    tab1, tab2 = st.tabs(["✨ Generate New MCQs", "📋 Format Existing MCQs"])
    
    # ================= TAB 1: GENERATE NEW MCQS =================
    with tab1:
        with st.container(border=True):
            st.markdown("#### 📄 1. Source Material (Optional)")
            st.caption("Leave blank to generate questions based purely on the topics provided below.")
            
            col_file, col_text = st.columns(2)
            with col_file:
                uploaded_file = st.file_uploader("Upload Raw MCQs One DOCX File", type=["pdf", "docx"])
            with col_text:
                raw_text = st.text_area("Or paste raw text here", height=100)
    
        with st.container(border=True):
            st.markdown("#### 🎯 2. Question Parameters")
            topics_input = st.text_input(
                "Allowed Topics for Table Mapping (comma-separated)", 
                placeholder="e.g., Hardware, Memory, Super Computers"
            )
            custom_instructions = st.text_area(
                "Custom Generation Instructions (Optional)", 
                placeholder="e.g., 'Focus heavily on numerical problems', 'Generate questions in Bengali'",
                height=100,
                key="t1_custom_instructions"
            )
    
        with st.container(border=True):
            st.markdown("#### 📊 3. Difficulty Breakdown")
            col1, col2, col3 = st.columns(3)
            with col1:
                num_easy = st.number_input("🟢 Easy", min_value=0, max_value=50, value=10)
            with col2:
                num_medium = st.number_input("🟡 Medium", min_value=0, max_value=50, value=15)
            with col3:
                num_hard = st.number_input("🔴 Hard", min_value=0, max_value=50, value=5)
    
        total_q = num_easy + num_medium + num_hard
    
        st.markdown("<br>", unsafe_allow_html=True) # Spacer
        
        col_btn, col_msg = st.columns([1, 2])
        with col_btn:
            generate_btn = st.button("🚀 Generate MCQs", use_container_width=True, type="primary")
        with col_msg:
            st.info(f"Target Generation: **{total_q}** Questions")
    
        if generate_btn:
            if not api_key_input:
                st.error("Please enter your API Key in the sidebar first!")
            elif total_q == 0:
                st.error("Please specify at least one question to generate.")
            elif not topics_input:
                st.error("Please provide at least one topic for table mapping.")
            else:
                with st.spinner(f"Engine running... Generating {total_q} questions using {model_choice}"):
                    text_to_process = ""
                    if uploaded_file:
                        if uploaded_file.name.endswith('.pdf'):
                            text_to_process = extract_text_from_pdf(uploaded_file)
                        elif uploaded_file.name.endswith('.docx'):
                            text_to_process = extract_text_from_docx(uploaded_file)
                    elif raw_text:
                        text_to_process = raw_text
                        
                    try:
                        # Fetch and save data into Session State so it doesn't vanish when edited
                        st.session_state["mcq_data_t1"] = generate_mcqs(
                            text_to_process, topics_input, custom_instructions, 
                            num_easy, num_medium, num_hard, model_choice, api_key_input
                        )
                        st.success("✨ Generation Complete!")
                    except Exception as e:
                        st.error(f"An error occurred: {e}")
    
        # Display the editable table and download button IF data exists in session state
        if st.session_state["mcq_data_t1"]:
            with st.container(border=True):
                st.markdown("#### 👀 Preview & Edit Generated Table")
                st.caption("Double-click any cell to edit its text. You can also add or delete rows using the tools on the right. Changes instantly apply to your download.")
                
                # Interactive Data Editor
                edited_data_t1 = st.data_editor(
                    st.session_state["mcq_data_t1"], 
                    use_container_width=True, 
                    num_rows="dynamic", 
                    key="editor_t1"
                )
            
            # Build the Word Document using the EDITED data
            create_mcq_docx(edited_data_t1, "Generated_MCQs.docx")
            
            with open("Generated_MCQs.docx", "rb") as file:
                st.download_button(
                    label="📥 Download Edited Word Document (.docx)",
                    data=file,
                    file_name="Generated_MCQs.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    
    # ================= TAB 2: FORMAT EXISTING MCQS =================
    with tab2:
        with st.container(border=True):
            st.markdown("#### 📝 1. Raw Input")
            st.caption("Upload or paste your unformatted questions.")
            col_file_t2, col_text_t2 = st.columns(2)
            with col_file_t2:
                uploaded_file_t2 = st.file_uploader("Upload Raw MCQs One DOCX File)", type=["pdf", "docx"], key="t2_pdf")
            with col_text_t2:
                raw_text_t2 = st.text_area("Or paste raw pre-written MCQs here", height=150, placeholder="1. What is CPU?\nA. Brain\nB. Memory\nC. Output\nD. Storage\nAnswer: A", key="t2_text")
        
        with st.container(border=True):
            st.markdown("#### 🎯 2. Formatting Parameters")
            topics_input_t2 = st.text_input("Target Topics (Optional)", placeholder="e.g., Computer Basics, Hardware", key="t2_topics")
            special_instructions = st.text_area(
                "Special AI Instructions (Optional)", 
                placeholder="e.g., 'Set Subject to ICT', 'Fix any Bengali spelling mistakes', 'Automatically fill in missing solution explanations'",
                height=100,
                key="t2_instructions"
            )
    
        st.markdown("<br>", unsafe_allow_html=True) # Spacer
    
        if st.button("🛠️ Format Existing MCQs", use_container_width=True, type="primary", key="btn_t2"):
            if not api_key_input:
                st.error("Please configure your API Key in the sidebar.")
            elif not uploaded_file_t2 and not raw_text_t2:
                st.error("Please provide your raw MCQs in the text box or upload a document.")
            else:
                with st.spinner(f"Parsing and reformatting questions using {model_choice}..."):
                    text_to_process = ""
                    if uploaded_file_t2:
                        if uploaded_file_t2.name.endswith('.pdf'):
                            text_to_process = extract_text_from_pdf(uploaded_file_t2)
                        elif uploaded_file_t2.name.endswith('.docx'):
                            text_to_process = extract_text_from_docx(uploaded_file_t2)
                    elif raw_text_t2:
                        text_to_process = raw_text_t2
                        
                    try:
                        # Save parsed data to Session State
                        st.session_state["mcq_data_t2"] = parse_existing_mcqs(
                            text_to_process, topics_input_t2, special_instructions, model_choice, api_key_input
                        )
                        st.success("✨ Successfully reformatted into table format!")
                    except Exception as e:
                        st.error(f"Error parsing MCQs: {e}")
    
        # Display the editable table and download button IF data exists in session state
        if st.session_state["mcq_data_t2"]:
            with st.container(border=True):
                st.markdown("#### 👀 Preview & Edit Formatted Table")
                st.caption("Double-click any cell to edit its text. You can also add or delete rows. Changes instantly apply to your download.")
                
                # Interactive Data Editor
                edited_data_t2 = st.data_editor(
                    st.session_state["mcq_data_t2"], 
                    use_container_width=True, 
                    num_rows="dynamic", 
                    key="editor_t2"
                )
            
            # Build the Word Document using the EDITED data
            create_mcq_docx(edited_data_t2, "Formatted_Ready_MCQs.docx")
            
            with open("Formatted_Ready_MCQs.docx", "rb") as f:
                st.download_button(
                    label="📥 Download Edited Word Document (.docx)",
                    data=f,
                    file_name="Formatted_Ready_MCQs.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_t2"
                )
