import os
import glob
from oletools.olevba import VBA_Parser
import msoffcrypto
import tempfile

PASSWORD = "BGDP"
SOURCE_DIR = "/Users/koffiefabienmichael/Downloads/FAAB APPLI/FAAB PROJET/KIT CHEF PROJET"
DEST_DIR = "/Users/koffiefabienmichael/Downloads/FAAB APPLI/FAAB PROJET/VBA_EXTRACTS"

if not os.path.exists(DEST_DIR):
    os.makedirs(DEST_DIR)

def try_decrypt_and_parse(file_path):
    print(f"Processing: {file_path}")
    
    # Check if the entire file is encrypted
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        if office_file.is_encrypted():
            print(f"  -> File is encrypted. Attempting decryption with password...")
            office_file.load_key(password=PASSWORD)
            fd, temp_path = tempfile.mkstemp(suffix=".xlsm")
            with open(temp_path, "wb") as f_decrypted:
                office_file.decrypt(f_decrypted)
            vba = VBA_Parser(temp_path)
            content = parse_vba(vba, file_path)
            os.close(fd)
            os.remove(temp_path)
            return content
        else:
            print(f"  -> File is not encrypted at document level.")
            vba = VBA_Parser(file_path)
            # oletools can usually read locked VBA projects anyway in modern Excel files
            # because the code is compressed, not always strongly encrypted.
            return parse_vba(vba, file_path)

def parse_vba(vba, original_path):
    if vba.detect_vba_macros():
        extracted_code = ""
        for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
            extracted_code += f"' --- Module: {vba_filename} ---\n"
            extracted_code += vba_code
            extracted_code += "\n\n"
        return extracted_code
    else:
        print("  -> No VBA macros detected.")
        return None

found_macros = 0
for root, dirs, files in os.walk(SOURCE_DIR):
    for file in files:
        if file.endswith((".xlsm", ".xlsb")):
            full_path = os.path.join(root, file)
            try:
                code = try_decrypt_and_parse(full_path)
                if code:
                    found_macros += 1
                    rel_path = os.path.relpath(full_path, SOURCE_DIR)
                    safe_name = rel_path.replace("/", "_").replace("\\", "_") + ".vba"
                    out_path = os.path.join(DEST_DIR, safe_name)
                    with open(out_path, "w", encoding="utf-8") as out_f:
                        out_f.write(code)
                    print(f"  => Saved extracted code to {safe_name}")
            except Exception as e:
                print(f"  -> Error processing {file}: {e}")

print(f"\nExtraction completed. Found macros in {found_macros} files.")
