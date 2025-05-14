file_path = os.path.join(directory, file)  # already defined in your loop
wb = openpyxl.load_workbook(file_path)
process_sachaufwand_links(wb, file_path)
wb.save(file_path)
