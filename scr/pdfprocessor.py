from SampleUtils import *
from omnipage import *

class PDFProcessor:
    def __init__(self, input_pdf, output_file, settings_file="settings_new.sts"):
        self.input_pdf = input_pdf
        self.output_file = output_file
        self.settings_file = settings_file
        self.company = YOUR_COMPANY  # Replace with your company name
        self.product = YOUR_PRODUCT  # Replace with your product name
        self.sdk_path = r"c:\Program Files\OmniPage\CSDK22\Bin"
        self._extracted_text = ""  # Initialize extracted text

    def initialize_sdk(self):
        """Initialize the OmniPage SDK."""
        rc = kRecInit(self.company, self.product, self.sdk_path)
        if rc not in [REC_OK, API_INIT_WARN, API_LICENSEVALIDATION_WARN]:
            ErrMsg("kRecInit error code = {}\n", rc)
            return False

        rc = RecInitPlus(self.company, self.product, self.sdk_path)
        if rc not in [REC_OK, API_INIT_WARN, API_LICENSEVALIDATION_WARN]:
            ErrMsg("RecInitPlus error code = {}\n", rc)
            kRecQuit()
            return False

        return True

    def load_settings(self):
        """Load settings for the OCR process."""
        rc = kRecLoadSettings(SID, self.settings_file)
        if rc != REC_OK:
            ErrMsg("kRecLoadSettings error code = {}\n", rc)
            return False
        return True

    def set_output_format(self):
        """Set the output format for the OCR process."""
        rc = RecSetOutputFormat(SID, "Converters.Text.XML")
        if rc != REC_OK:
            ErrMsg("RecSetOutputFormat error code = {}\n", rc)
            return False
        return True

    def process_pdf(self):
        """Process the input PDF using OmniPage OCR."""
        pdf_files = [self.input_pdf]
        rc = RecProcessPagesEx(SID, self.output_file, pdf_files, None)
        if rc != REC_OK:
            ErrMsg("RecProcessPagesEx error code = {}\n", rc)
            rc, rpp_errors = RecGetRPPErrorList()
            for error in rpp_errors:
                ErrMsg("- {}", error.rc)
            return False
        return True

    def extract_text_from_xml(self):
        """Extract text from the generated XML file."""
        try:
            import xml.etree.ElementTree as ET
            tree = ET.parse(self.output_file)
            root = tree.getroot()
            text_content = ""
            for text_elem in root.findall(".//text"):
                if text_elem.text:
                    text_content += text_elem.text + " "
            self._extracted_text = text_content.strip()
        except Exception as e:
            ErrMsg("Error extracting text from XML: {}\n", e)
            return False
        return True

    def cleanup(self):
        """Clean up and release OmniPage resources."""
        RecQuitPlus()
        kRecQuit()

    @property
    def extracted_text(self):
        """Return the extracted text."""
        return self._extracted_text

    def run(self):
        """Main function to execute the OCR process."""
        if not self.initialize_sdk():
            return

        if not self.load_settings():
            self.cleanup()
            return

        if not self.set_output_format():
            self.cleanup()
            return

        if not self.process_pdf():
            self.cleanup()
            return

        if not self.extract_text_from_xml():
            self.cleanup()
            return

        self.cleanup()
        print(f"OCR process completed successfully. Output saved to: {self.output_file}")