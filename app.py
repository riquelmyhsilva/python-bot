import pyautogui as pygui
import openpyxl as opxl
import pyperclip as clip
from time import sleep

workbook = opxl.load_workbook("products.xlsx")
sheet_products = workbook["Produtos"]

for row in sheet_products.iter_rows(min_row=2):
    product_name = row[0].value
    clip.copy(product_name)
    pygui.click(435, 326, duration=1)
    pygui.hotkey("ctrl", "v")

    description = row[1].value
    clip.copy(description)
    pygui.click(448, 434, duration=1)
    pygui.hotkey("ctrl", "v")

    category = row[2].value
    clip.copy(category)
    pygui.click(441, 545, duration=1)
    pygui.hotkey("ctrl", "v")

    code = row[3].value
    clip.copy(code)
    pygui.click(449, 632, duration=1)
    pygui.hotkey("ctrl", "v")

    kg_weight = row[4].value
    clip.copy(kg_weight)
    pygui.click(444, 720, duration=1)
    pygui.hotkey("ctrl", "v")

    dimensions_lap = row[5].value  # lap = L x A x P (Length x Height x Depth)
    clip.copy(dimensions_lap)
    pygui.click(448, 806, duration=1)
    pygui.hotkey("ctrl", "v")

    pygui.click(447, 865, duration=1)
    sleep(3)

    price = row[6].value
    clip.copy(price)
    pygui.click(443, 350, duration=1)
    pygui.hotkey("ctrl", "v")

    quantity = row[7].value
    clip.copy(quantity)
    pygui.click(445, 437, duration=1)
    pygui.hotkey("ctrl", "v")

    valid_date = row[8].value
    clip.copy(valid_date)
    pygui.click(447, 519, duration=1)
    pygui.hotkey("ctrl", "v")

    color = row[9].value
    clip.copy(color)
    pygui.click(443, 605, duration=1)
    pygui.hotkey("ctrl", "v")

    size = row[10].value
    pygui.click(446, 692, duration=1)
    if size == "Pequeno":
        pygui.click(446, 735, duration=1)
    elif size == "Médio":
        pygui.click(446, 754, duration=1)
    else:
        pygui.click(446, 787, duration=1)

    material = row[11].value
    clip.copy(material)
    pygui.click(440, 781, duration=1)
    pygui.hotkey("ctrl", "v")

    pygui.click(444, 839, duration=1)
    sleep(3)

    producer = row[12].value
    clip.copy(producer)
    pygui.click(427, 368, duration=1)
    pygui.hotkey("ctrl", "v")

    origin_country = row[13].value
    clip.copy(origin_country)
    pygui.click(433, 455, duration=1)
    pygui.hotkey("ctrl", "v")

    observations = row[14].value
    clip.copy(observations)
    pygui.click(439, 549, duration=1)
    pygui.hotkey("ctrl", "v")

    bars_code = row[15].value
    clip.copy(bars_code)
    pygui.click(441, 672, duration=1)
    pygui.hotkey("ctrl", "v")

    warehouse_location = row[16].value
    clip.copy(warehouse_location)
    pygui.click(446, 762, duration=1)
    pygui.hotkey("ctrl", "v")

    pygui.click(451, 821, duration=1)
    pygui.click(853, 592, duration=1)
    pygui.click(856, 591, duration=1)
    pygui.click(1047, 596, duration=1)

    quit()
