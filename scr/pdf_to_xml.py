import os
import logging
import time
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from itertools import islice
from pdfprocessor import PDFProcessor  # Ensure this module handles env vars internally

# 1. Configure Logging
logging.basicConfig(filename="pdf_processing.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")


# 2. Single PDF Worker (With Retry Logic)
def process_single_pdf(pdf_path, retry_count=3):
    input_path = Path(pdf_path)
    output_path = input_path.with_suffix('.xml')

    for attempt in range(retry_count):
        try:
            processor = PDFProcessor(str(input_path), str(output_path))
            processor.run()
            extracted_text = processor.extracted_text

            if extracted_text:
                document_name = input_path.stem
                xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<document>
    <name>{document_name}</name>
    <pages>
        <page number="1">
            <text>{extracted_text}</text>
        </page>
    </pages>
</document>"""

                with open(output_path, 'w', encoding='utf-8') as xml_file:
                    xml_file.write(xml_content)

                return True
        except Exception as e:
            logging.error(f"Error processing {pdf_path} (Attempt {attempt + 1}): {e}")
            time.sleep(1)

    return False


# 3. Batch Processing Function
def process_pdf_batch(pdf_batch):
    return [(pdf, process_single_pdf(pdf)) for pdf in pdf_batch]


# 4. Helper: Chunking PDFs for Batching
def chunked_iterable(iterable, size):
    it = iter(iterable)
    while chunk := list(islice(it, size)):
        yield chunk


# 5. Main Function
def main():
    start_time = time.time()

    # Use environment variable or fallback default
    pdf_folder_path = os.environ.get("PDF_INPUT_FOLDER", "data/pdfs")

    all_pdfs = [os.path.join(root, f)
                for root, _, files in os.walk(pdf_folder_path)
                for f in files if f.lower().endswith(".pdf")]

    total_pdfs = len(all_pdfs)
    if total_pdfs == 0:
        print("No PDFs found in the specified folder.")
        return

    print(f"Processing folder with {total_pdfs} PDFs...")
    logging.info(f"Processing folder with {total_pdfs} PDFs...")

    num_workers = min(30, multiprocessing.cpu_count() - 1)
    batch_size = 5

    success_count = 0
    failed_files = []

    with ProcessPoolExecutor(max_workers=num_workers) as executor:
        future_map = {executor.submit(process_pdf_batch, batch): batch
                      for batch in chunked_iterable(all_pdfs, batch_size)}

        for future in as_completed(future_map):
            batch = future_map[future]
            try:
                results = future.result()
                for pdf, success in results:
                    if success:
                        success_count += 1
                    else:
                        failed_files.append(pdf)
            except Exception as e:
                logging.error(f"Batch processing error: {e}")

    # Compare expected and actual output
    expected_xmls = {Path(pdf).with_suffix('.xml').name for pdf in all_pdfs}
    actual_xmls = {f for f in os.listdir(pdf_folder_path) if f.lower().endswith(".xml")}
    successfully_created_count = len(expected_xmls & actual_xmls)

    # Log final report
    elapsed = time.time() - start_time
    minutes, seconds = divmod(elapsed, 60)
    msg = f"Processing complete. {successfully_created_count} of {total_pdfs} documents processed. Time: {int(minutes)} min {seconds:.2f} sec."
    print(msg)
    logging.info(msg)

    if failed_files:
        logging.error(f"{len(failed_files)} PDFs failed:")
        for pdf in failed_files:
            logging.error(pdf)


if __name__ == "__main__":
    main()
