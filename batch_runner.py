import os
from word_extractor import extract_sections_from_docx

INPUT_DIR = "input_files"
OUTPUT_DIR = "output_files"

def main():
    print("=== Batch mode: Word → CSV ===")

    # Upewnij się że foldery istnieją (dla bezpieczeństwa)
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".docx")]

    if not files:
        print("Brak plików DOCX w folderze input_files")
        return

    for filename in files:
        docx_path = os.path.join(INPUT_DIR, filename)
        csv_path = os.path.join(OUTPUT_DIR, filename.replace(".docx", ".csv"))

        print(f"Przetwarzanie: {filename}")
        try:
            extract_sections_from_docx(
                docx_path,
                csv_path,
                detect_language_flag=False,
                progress_callback=None
            )
            print(f" -> Zapisano CSV: {csv_path}")
        except Exception as e:
            print(f" !! BŁĄD dla {filename}: {e}")

    print("=== GOTOWE ===")

if __name__ == "__main__":
    main()
