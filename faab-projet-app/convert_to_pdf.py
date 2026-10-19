import os
import subprocess

# Paths
KIT_DIR = "/Users/koffiefabienmichael/Downloads/FAAB APPLI/FAAB PROJET/KIT CHEF PROJET"
PUBLIC_DOCS_DIR = "/Users/koffiefabienmichael/Downloads/FAAB APPLI/FAAB PROJET/faab-projet-app/public/pdfs"
os.makedirs(PUBLIC_DOCS_DIR, exist_ok=True)

def convert_word_to_pdf(input_path, output_path):
    print(f"Converting WORD: {input_path}")
    applescript = f'''
    tell application "Microsoft Word"
        set input_file to POSIX file "{input_path}"
        set output_file to "{output_path}"
        set theDoc to open (input_file as alias)
        save as theDoc file name output_file file format format PDF
        close theDoc saving no
    end tell
    '''
    try:
        subprocess.run(["osascript", "-e", applescript], check=True, capture_output=True)
        return True
    except Exception as e:
        print(f"Error converting {input_path}: {e}")
        return False

def convert_ppt_to_pdf(input_path, output_path):
    print(f"Converting PPT: {input_path}")
    applescript = f'''
    tell application "Microsoft PowerPoint"
        set input_file to POSIX file "{input_path}"
        set output_file to "{output_path}"
        set thePres to open (input_file as alias)
        save thePres in output_file as save as PDF
        close thePres
    end tell
    '''
    try:
        subprocess.run(["osascript", "-e", applescript], check=True, capture_output=True)
        return True
    except Exception as e:
        print(f"Error converting {input_path}: {e}")
        return False

files_processed = 0
for root, _, files in os.walk(KIT_DIR):
    for f in files:
        if f.startswith('~'): continue
        f_path = os.path.join(root, f)
        
        # Unique safe name for internal mapping
        # Let's use the exact label-matching logic here too.
        safe_name = f.lower().replace(" ", "_").replace("'", "").replace("’", "").replace("é", "e").replace("è", "e").replace("ê", "e").replace("ç", "c").replace("à", "a")
        output_f = f"{safe_name}.pdf"
        output_path = os.path.join(PUBLIC_DOCS_DIR, output_f)

        if f.endswith(".docx"):
            convert_word_to_pdf(f_path, output_path)
            files_processed += 1
        elif f.endswith(".pptx"):
            convert_ppt_to_pdf(f_path, output_path)
            files_processed += 1

print(f"Conversion complete! {files_processed} files processed.")
