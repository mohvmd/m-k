import re
import os
import platform

class PhotoProcessor:
    def __init__(self):
        self.ocr_engine = None
        self.engine_type = None
        self.init_ocr()

    def init_ocr(self):
        """Initializes the OCR engine, preferring native macOS Vision OCR if available on Mac, otherwise RapidOCR."""
        # On macOS, try to check if we can use native Apple Vision framework
        if platform.system() == "Darwin":
            try:
                # Try to import native macOS modules (via PyObjC or ocrmac if available)
                from ocrmac import OCR
                self.ocr_engine = OCR
                self.engine_type = "mac_vision"
                print("Prod by Mohamed Ayman Samir")
                return
            except ImportError:
                pass
            
        # Fallback/Primary for all platforms: RapidOCR
        try:
            from rapidocr_onnxruntime import RapidOCR
            self.ocr_engine = RapidOCR()
            self.engine_type = "rapidocr"
            print("تم تحميل محرك RapidOCR بنجاح.")
        except ImportError:
            self.ocr_engine = None
            self.engine_type = None
            print("تحذير: لم يتم العثور على مكتبات الـ OCR. يرجى تثبيت المتطلبات.")

    def extract_text(self, image_path):
        """Extracts all text from the image using the loaded OCR engine."""
        if not self.ocr_engine:
            raise Exception("لم يتم تهيئة محرك OCR. يرجى التأكد من تثبيت المكتبات المطلوبة.")

        if not os.path.exists(image_path):
            raise FileNotFoundError(f"الملف غير موجود: {image_path}")

        extracted_texts = []

        if self.engine_type == "mac_vision":
            # Using ocrmac wrapper for macOS Vision
            try:
                results = self.ocr_engine(image_path).recognize()
                # ocrmac returns list of tuples: (text, confidence, bounding_box)
                for item in results:
                    if item and len(item) > 0:
                        extracted_texts.append(item[0])
            except Exception as e:
                print(f"خطأ أثناء تشغيل macOS Vision OCR: {e}. محاولة استخدام RapidOCR كبديل...")
                # If native fails, try to fallback to RapidOCR if imported
                self.engine_type = "rapidocr"
                try:
                    from rapidocr_onnxruntime import RapidOCR
                    self.ocr_engine = RapidOCR()
                except ImportError:
                    raise Exception(f"فشل macOS OCR ولم يتم العثور على RapidOCR كبديل: {e}")

        if self.engine_type == "rapidocr":
            # Using RapidOCR
            try:
                result, elapse = self.ocr_engine(image_path)
                if result:
                    for line in result:
                        if len(line) >= 2:
                            extracted_texts.append(line[1])
            except Exception as e:
                raise Exception(f"خطأ أثناء معالجة الصورة باستخدام OCR: {str(e)}")

        return extracted_texts

    def find_site_codes(self, image_path):
        """
        Processes the image, extracts text, and filters out Site Codes using regex.
        Returns a list of unique Site Codes found in the image.
        """
        texts = self.extract_text(image_path)
        combined_text = "\n".join(texts)

        # Regex explanation:
        # Match pattern: optional U/u followed by DEL/del (or similar variants) followed by 3 to 5 digits.
        # Example: DEL4996, UDEL4654, del1039, Udel4888, etc.
        pattern = r'[Uu]?[Dd][Ee][Ll]\d{3,5}'
        
        matches = re.findall(pattern, combined_text)
        
        # Clean and uppercase matches, remove duplicates while preserving order
        unique_codes = []
        for match in matches:
            cleaned_code = match.strip().upper()
            if cleaned_code not in unique_codes:
                unique_codes.append(cleaned_code)

        return unique_codes
