from powerpoint import PPT


PPT_PATH = r'D:\Project\commerce\parser_powerpoint\init_pptx\Fashion.pptx'  # исходник
PPT_CHANGED_PATH = r'D:\Project\commerce\parser_powerpoint\init_pptx\new.pptx' # изменённый файл


def duplicate_slide():
    """Пример копирование 3-го слайда и вставки его в 5-ую позицию"""
    # не работает
    ppt = PPT(PPT_PATH)
    ppt.duplicate_slide(2, 4)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_text_example():
    """пример изменения пераого текстового блока в 3 слайде"""
    ppt = PPT(PPT_PATH)
    ppt.slides[2].texts[0].change_text('ИЗМЕНЁННЫЙ ТЕКСТ')
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


def change_image_example():
    """пример изменения первого изображения блока в 2 слайде"""
    ppt = PPT(PPT_PATH)
    image_path = r'image.png'
    ppt.slides[1].images[0].change_image(image_path)
    ppt.save_as(PPT_CHANGED_PATH)
    ppt.close()


if __name__ == '__main__':
    change_image_example()