Submitter name: Abhishek Kumar

Roll No.: 2019csb1062

Course: CS305 Software Engineering

=================================

1. What does this program do?

   This is a python program, which is used for extracting metadata from book cover Pages such as Title of Book, Author names, ISBN Number and Publisher Name.
   In this program we used Tesseract OCR(Optical Image Recognization) for extracting text from images.
   From the extracted text to find out which entity belongs to which category we used NER(Named Entity Recognization) under Spacy.
   Extracted metadata is stored in an XLSX output file where 4 columns are there, 1. Title of the book, 2. Names of the authors, 3. Publishers, 4. ISBN numbers.

2. A description of how this program works (i.e. its logic)

   This python program used Tessearct OCR to recognize text in images of book cover pages. In this program we used NER for identifying which word belogs to which entity
   like if the identified word is "Robert" then NER categorize it as "PERSON" entity, similarly for ISBN number and other stuff. In this program we made primarily two main
   functions namely: "single_image_ocr" and "multiple_image_ocr". Apart from these six secondary fucntions are there which are like child of these two primary functions.
   After all the data is extracted and identified the data is feed into the output XLSX file. This program takes two imputs from user, one is whether user wants to process
   a single image file or a directory of images, second is the path of corresponding item. These details user can provide through CLI arguments, first arguments is flag value
   which determine whether user wants to process a single image (flag value = 0) file or a directory of images (flag value != 0). Second argument is the path of image in
   "single_image_ocr" case and path of directory in case of "multiple_image_ocr". Input images can be in .jpg/.jpeg/.png format.

   Testing of this program is done using the pytest.
   For testing following functions are implemented:
   test_single_image_ocr: This will take its two arguments as mentioned above and make assertion on the basis of number of rows updated in XLSX file.
   test_multiple_image_ocr: This is also works similar to above test but its is for multiple images.

3. How to compile and run this program

   To run the main.py file: python main.py [flag value] [Path]
   Flag value: If '0' then "single_image_ocr" will work else "multiple_image_ocr" will work.
   Path: Image path in case of "single_image_ocr" and directory path in case of "multiple_image_ocr".

   To run the test_main.py file: coverage run -m pytest -q -s test_main.py
   To see the coverage Report: coverage report -m / coverage report -m -i