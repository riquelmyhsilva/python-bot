import keyboard
import pyautogui as pygui
import openpyxl as opxl
import pyperclip as clip
from time import sleep
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Starting the script...")


def copy_and_paste(value, x, y, duration=1):
    """Helper function to copy a value to clipboard and paste it at specified coordinates."""
    try:
        if value is None:
            logging.warning(f"Missing value for coordinates ({x}, {y})")
            return
        clip.copy(value)
        pygui.click(x, y, duration=duration)
        pygui.hotkey("ctrl", "v")
    except Exception as e:
        logging.error(
            f"Error while copying and pasting value '{value}' at ({x}, {y}): {e}"
        )


def click_with_logging(x, y, duration=1):
    """Helper function to click at specified coordinates with logging."""
    try:
        pygui.click(x, y, duration=duration)
        logging.info(f"Clicked at coordinates ({x}, {y})")
    except Exception as e:
        logging.error(f"Error while clicking at ({x}, {y}): {e}")


# Load the Excel workbook
try:
    workbook = opxl.load_workbook("products.xlsx")
    sheet_products = workbook["Produtos"]
except Exception as e:
    logging.error(f"Error loading workbook or sheet: {e}")
    quit()

# Process each row in the Excel sheet
try:
    for row in sheet_products.iter_rows(min_row=2):
        # Check if the Esc key is pressed
        if keyboard.is_pressed("esc"):
            logging.info("Esc key pressed. Exiting script...")
            break

        try:
            product_name = row[0].value
            logging.info(f"Processing product: {product_name}")
            copy_and_paste(product_name, 435, 326)

            description = row[1].value
            copy_and_paste(description, 448, 434)

            category = row[2].value
            copy_and_paste(category, 441, 545)

            code = row[3].value
            copy_and_paste(code, 449, 632)

            kg_weight = row[4].value
            copy_and_paste(kg_weight, 444, 720)

            dimensions_lap = row[5].value  # lap = L x A x P (Length x Height x Depth)
            copy_and_paste(dimensions_lap, 448, 806)

            click_with_logging(447, 865)
            sleep(3)

            price = row[6].value
            copy_and_paste(price, 443, 350)

            quantity = row[7].value
            copy_and_paste(quantity, 445, 437)

            valid_date = row[8].value
            copy_and_paste(valid_date, 447, 519)

            color = row[9].value
            copy_and_paste(color, 443, 605)

            size = row[10].value
            click_with_logging(446, 692)
            if size == "Pequeno":
                click_with_logging(446, 735)
            elif size == "Médio":
                click_with_logging(446, 754)
            elif size == "Grande":
                click_with_logging(446, 787)
            else:
                logging.warning(f"Unexpected size value: {size}")

            material = row[11].value
            copy_and_paste(material, 440, 781)

            click_with_logging(444, 839)
            sleep(3)

            producer = row[12].value
            copy_and_paste(producer, 427, 368)

            origin_country = row[13].value
            copy_and_paste(origin_country, 433, 455)

            observations = row[14].value
            copy_and_paste(observations, 439, 549)

            bars_code = row[15].value
            copy_and_paste(bars_code, 441, 672)

            warehouse_location = row[16].value
            copy_and_paste(warehouse_location, 446, 762)

            click_with_logging(451, 821)
            click_with_logging(853, 592)
            click_with_logging(856, 591)
            click_with_logging(1047, 596)

        except Exception as e:
            logging.error(f"Error processing row: {e}")

except KeyboardInterrupt:
    logging.info("Script interrupted by user. Exiting...")

logging.info("Script finished.")
