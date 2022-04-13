import zipfile


path_to_zip_file = r"Y:\RU\Archive\Scraped Files\Russia\6\2018-01.zip"
directory_to_extract_to = r"Y:\RU\Archive\Scraped Files\Russia\6"

with zipfile.ZipFile(path_to_zip_file, 'r') as zip_ref:
    zip_ref.extractall(directory_to_extract_to)
