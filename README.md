[INFO] [1/4] Opening document...
[INFO] [2/4] Analyzing document...
[WARNING] Words count: 0. It might be a scanned pdf, which is not supported yet.
[INFO] [3/4] Parsing pages...
[INFO] [4/4] Creating pages...
Traceback (most recent call last):
  File "\\\Testpdf.py", line 78, in <module>
    split_pdf_to_docx_per_page(pdf_path, docx_output_dir, page_count)
  File "\\\Testpdf.py", line 27, in split_pdf_to_docx_per_page
    cv.convert(page_docx, start=i, end=i)
  File "\pdf2docx\converter.py", line 349, in convert
    self.parse(start, end, pages, **settings).make_docx(docx_filename, **settings)
  File "s\pdf2docx\converter.py", line 207, in make_docx
    raise ConversionException('No parsed pages. Please parse page first.')
pdf2docx.converter.ConversionException: No parsed pages. Please parse page first.
