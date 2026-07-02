import pdfplumber
import pandas as pd
import re

# --- 1. CONFIGURATION ---
CHECKPOINT_MAP = {
    "issue for consideration": "khannanabagidamak puthorakpa waapham",
    "headnotes": "machang warol",
    "held": "waroishanda",
    "case law cited": "sijinnakhiba case law",
    "list of acts": "act shinggi maming",
    "list of keywords": "maru oiba wahei shing",
    "case arising from": "waathokki hourakpham",
    "appearances for parties":"waakatloyshinggi maikeida leplibashing",
    "judgment / order": "mapung yathang/waayel yathang",
}

JUNK_HEADERS = [
    "digital high court of manipur reports",
    "hcmr",
    "="
]

# --- 2. CLEANERS & SPLITTERS ---
def clean_for_excel(text):
    if not text:
        return ""
    clean_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    if len(clean_text) > 32000:
        clean_text = clean_text[:32000] + "\n\n... [TEXT TRUNCATED DUE TO EXCEL LIMIT]"
    return clean_text

def is_junk_header(line):
    clean = line.strip().lower()
    if clean.isdigit(): 
        return True 
    for junk in JUNK_HEADERS:
        if junk in clean:
            return True
    return False

def split_into_sentences(text):
    if not text:
        return []
        
    protected = text
    
    # 1. Protect single-letter initials (e.g., "S.", "H.", "M.V.")
    protected = re.sub(r'(?<=\b[A-Za-z])\.', '<DOT>', protected)
    
    # 2. Protect specific legal & common abbreviations (case-insensitive)
    abbrs = [
        'adv', 'sr', 'mr', 'mrs', 'dr', 'j', 'v', 'vs', 'no', 
        'addl', 'pp', 'ref', 'cril', 'rs', 'pw', 'p.w', 'viz', 'i.e', 'e.g'
    ]
    for a in abbrs:
        # \1 captures the exact original casing of the abbreviation
        pattern = r'(?i)\b(' + a.replace('.', r'\.') + r')\.'
        protected = re.sub(pattern, r'\1<DOT>', protected)
        
    # 3. Protect specific complex ones like Cr.P.C.
    protected = re.sub(r'(?i)(cr\.p\.c)\.', r'\1<DOT>', protected)
    
    # Now split safely at real periods, question marks, and exclamation points
    sentences = re.split(r'(?<=[.!?])\s+', protected.strip())
    
    # Restore the periods and clean up
    return [s.replace('<DOT>', '.').strip() for s in sentences if s.strip()]

# --- 3. EXTRACTION LOGIC ---
def parse_pdf_into_blocks(pdf_path, is_manipuri):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    except Exception as e:
        print(f"Error opening {pdf_path}: {e}")
        return {}
    
    lines = full_text.split('\n')
    sections = {}
    current_marker = "PREAMBLE"
    current_text = []
    
    skipping_disclaimer = False
    in_judgment_section = False  
    
    para_pattern = re.compile(r'^\s*([\[\(]\d+[\]\)]|\d+[\.\)])(?:\s+|$)')
    
    for line in lines:
        clean_line = line.strip()
        if not clean_line: continue
            
        if is_junk_header(clean_line): continue
            
        if "DISCLAIMER:" in clean_line:
            skipping_disclaimer = True
        if skipping_disclaimer:
            if "===" in clean_line or "YUMNAM" in clean_line.upper():
                skipping_disclaimer = False
            else:
                continue
                
        found_marker = None
        is_numbered_para = False
        
        # Normalize extra spaces to catch messy PDF spacing
        norm_line = re.sub(r'\s+', ' ', clean_line.lower())
        
        for eng_key, man_key in CHECKPOINT_MAP.items():
            if not is_manipuri and eng_key.lower() in norm_line:
                found_marker = eng_key.upper()
                break
            elif is_manipuri and man_key.lower() in norm_line:
                found_marker = eng_key.upper()
                break
                
        if found_marker == "JUDGMENT / ORDER":
            in_judgment_section = True
                
        match = None
        if not found_marker and in_judgment_section:
            match = para_pattern.match(line) 
            if match:
                raw_match = match.group(1) 
                para_num = re.search(r'\d+', raw_match).group()
                found_marker = f"PARAGRAPH {para_num}"
                is_numbered_para = True
                
        if found_marker:
            if current_marker in sections:
                sections[current_marker] += "\n" + "\n".join(current_text)
            else:
                sections[current_marker] = "\n".join(current_text)
            
            current_marker = found_marker
            
            if is_numbered_para:
                text_after_num = line[match.end():].strip()
                current_text = [text_after_num] if text_after_num else []
            else:
                current_text = [] 
        else:
            current_text.append(clean_line)
            
    if current_text:
        if current_marker in sections:
            sections[current_marker] += "\n" + "\n".join(current_text)
        else:
            sections[current_marker] = "\n".join(current_text)
        
    return sections

# --- 4. EXECUTION ---
def main():
    print("Reading PDFs, aligning, and applying strict abbreviation rules...")
    eng_blocks = parse_pdf_into_blocks('english3.pdf', is_manipuri=False)
    man_blocks = parse_pdf_into_blocks('manipuri2.pdf', is_manipuri=True)

    aligned_rows = []
    
    for marker_key in eng_blocks.keys():
        eng_text = eng_blocks.get(marker_key, "").strip()
        man_text = man_blocks.get(marker_key, "").strip()
        
        man_display_marker = marker_key
        for eng, man in CHECKPOINT_MAP.items():
            if eng.upper() == marker_key:
                man_display_marker = man.upper()
                break
        
        # 1. Insert the Checkpoint Header Row
        aligned_rows.append({
            'English': f"--- {marker_key} ---", 
            'Manipuri (Latin)': f"--- {man_display_marker} ---"
        })
        
        # 2. Protect the Preamble from the sentence splitter
        if marker_key == "PREAMBLE":
            eng_sentences = [eng_text] if eng_text else []
            man_sentences = [man_text] if man_text else []
        else:
            eng_sentences = split_into_sentences(eng_text)
            man_sentences = split_into_sentences(man_text)
        
        # 3. Pair them up
        max_sentences = max(len(eng_sentences), len(man_sentences))
        for i in range(max_sentences):
            e_sent = eng_sentences[i] if i < len(eng_sentences) else ""
            m_sent = man_sentences[i] if i < len(man_sentences) else ""
            
            aligned_rows.append({
                'English': clean_for_excel(e_sent), 
                'Manipuri (Latin)': clean_for_excel(m_sent)
            })

    df = pd.DataFrame(aligned_rows)
    output_filename = 'sentence_aligned_judgment_FINAL.xlsx'
    
    try:
        df.to_excel(output_filename, index=False)
        print(f"✅ Success! Sentences aligned and saved to '{output_filename}'")
    except Exception as e:
        print(f"❌ Error saving to Excel: {e}")

if __name__ == "__main__":
    main()