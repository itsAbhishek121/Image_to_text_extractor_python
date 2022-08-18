import main

def test_single_image_ocr():
    r = main.single_image_ocr("C:\\Users\\Asus\\Desktop\\305\\data\\x.png")
    assert r == 2

def test_multiple_image_ocr():
    r = main.multiple_image_ocr("C:\\Users\\Asus\\Desktop\\305\\data\\abc")
    assert r == 3
