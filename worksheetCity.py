from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re

def convert_to_mm(value):
    print(f"Original value: {value}")  # Debug print

    # Try to convert the value based on different measurement types
    match_cm = re.search(r'(\d+(\.\d+)?)\s*cm', value, re.IGNORECASE)
    match_mm = re.search(r'(\d+(\.\d+)?)\s*mm', value, re.IGNORECASE)
    match_inch = re.search(r'(\d+(\.\d+)?)\s*inch', value, re.IGNORECASE)
    
    if match_cm:
        numeric_value = float(match_cm.group(1))
        converted_value = int(numeric_value * 10)  # Convert cm to mm
        print(f"Converted from cm to mm: {converted_value}")  # Debug print
        return converted_value
    elif match_mm:
        numeric_value = float(match_mm.group(1))
        converted_value = int(numeric_value)  # Value is already in mm
        print(f"Value in mm: {converted_value}")  # Debug print
        return converted_value
    elif match_inch:
        numeric_value = float(match_inch.group(1))
        converted_value = int(numeric_value * 25.4)  # Convert inches to mm
        print(f"Converted from inches to mm: {converted_value}")  # Debug print
        return converted_value
    else:
        try:
            numeric_value = float(value)
            print(f"Value in mm: {int(numeric_value)}")  # Debug print
            return int(numeric_value)
        except ValueError:
            print("Conversion error: Invalid numeric value")  # Debug print
            return None




def calculate_sizes_vertical_blinds(product_data):
    # Filter only for Vertical Blinds
    if product_data.get('Product Type') == 'Vertical Blind':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)

        if width is None:
            product_data['Finished Size'] = 'Invalid Width'
            product_data['Cut Rail Size'] = 'Invalid Width'
            product_data['Qty Louvers'] = ''
            return product_data

        operation_type = product_data.get('Operation Types')
        measurement_type = product_data.get('Measurement Type')

        finished_size = cut_rail_size = 0

        if measurement_type == 'Recess':
            finished_size = width - 10
        elif measurement_type == 'Exact':
            finished_size = width
            # Increase louver drop size by 10mm for 'Exact' measurement type
            louver_drop_size = convert_to_mm(product_data.get('Length', '0'))
            if louver_drop_size is not None:
                louver_drop_size += 10

        if operation_type == 'Manual Operation':
            cut_rail_size = finished_size - 20
        elif operation_type == 'Wand Operation':
            cut_rail_size = finished_size - 10

        product_data['Finished Size'] = f'{finished_size}mm'
        product_data['Cut Rail Size'] = f'{cut_rail_size}mm'
        product_data['Qty Louvers'] = ''
        
        # Check if louver drop size is not None and add it to product_data
        if louver_drop_size is not None:
            product_data['Louver Drop Size'] = f'{louver_drop_size}mm'
        else:
            product_data['Louver Drop Size'] = 'Invalid Length'

        product_data.pop('Length', None)
        product_data.pop('Measurement Protection', None)


    return product_data


def extract_vertical_blind_data(driver, customer_name):
    if "[Sample]" in driver.page_source:
        return None

    table = driver.find_element(By.ID, 'data-table')
    rows = table.find_elements(By.TAG_NAME, 'tr')

    all_product_data = []

    for row in rows[1:]:
        product_data = {'Customer Name': customer_name}

        fabric_name_elements = row.find_elements(By.XPATH, './/td/a')
        fabric_name = fabric_name_elements[1].text.strip() if len(fabric_name_elements) > 1 else 'Unknown Fabric'
        product_data['Fabric Name'] = fabric_name

        product_details = row.find_elements(By.XPATH, './/div[@class="basket_custom_option"]')
        for detail in product_details:
            label = detail.find_element(By.CLASS_NAME, 'basket_custom_option_label').text.strip(':')
            value = detail.find_element(By.XPATH, './/following-sibling::div').text
            product_data[label] = value

        # Call the appropriate calculate_sizes function based on Product Type
        if product_data.get('Product Type') == 'Vertical Blind':
            product_data = calculate_sizes_vertical_blinds(product_data)
        elif product_data.get('Product Type') == 'Roller Blind' and product_data.get('Roller Type') == 'Standard Roller':
            product_data = calculate_sizes_standard_roller_blinds(product_data)
        elif product_data.get('Product Type') == 'Allusion Blind':
            product_data = calculate_sizes_allusion_blinds(product_data)
        elif product_data.get('Product Type') == 'Roller Blind' and product_data.get('Roller Type') == 'Cassette Roller':
            product_data = calculate_sizes_cassette_roller_blind(product_data)
        elif product_data.get('Product Type') == 'Grip Fit Roller Blind':
            product_data = calculate_sizes_grip_fit_roller_blinds(product_data)
        elif product_data.get('Product Type') == 'Day & Night Blind':
            product_data = calculate_sizes_day_and_night_blinds(product_data)
        elif product_data.get('Product Type') == 'Perfect Fit Pleated':
            product_data = calculate_sizes_perfect_fit_pleated(product_data)
        

        all_product_data.append(product_data)

    return all_product_data


def calculate_sizes_allusion_blinds(product_data):
    # Filter only for Allusion Blinds
    if product_data.get('Product Type') == 'Allusion Blind':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)

        if width is None:
            product_data['Finished Size'] = 'Invalid Width'
            product_data['Cut Rail Size'] = 'Invalid Width'
            product_data['Qty Louvers'] = ''
            return product_data

        operation_type = product_data.get('Operation Types')
        measurement_type = product_data.get('Measurement Type')

        finished_size = cut_rail_size = 0
        louver_drop_size = None  # Initialize louver_drop_size with None

        if measurement_type == 'Recess':
            finished_size = width - 10
        elif measurement_type == 'Exact':
            finished_size = width
            # Increase louver drop size by 10mm for 'Exact' measurement type
            louver_drop_size = convert_to_mm(product_data.get('Length', '0'))
            if louver_drop_size is not None:
                louver_drop_size += 10

        # For Allusion Blinds, subtract 12mm from the length
        length_str = product_data.get('Length', '0')
        length = convert_to_mm(length_str)
        if length is not None:
            adjusted_length = length - 12
            louver_drop_size = adjusted_length  # Store the adjusted louver drop size in a variable
            product_data['Length'] = f'{adjusted_length}mm'  # Update the Length value
        else:
            louver_drop_size = None
            product_data['Louver Drop Size'] = 'Invalid Length'

        # Check if louver drop size is not None and add it to product_data
        if louver_drop_size is not None:
            product_data['Louver Drop Size'] = f'{louver_drop_size}mm'

        if operation_type == 'Manual Operation':
            cut_rail_size = finished_size - 20
        elif operation_type == 'Want Operation':
            cut_rail_size = finished_size - 10

        product_data['Finished Size'] = f'{finished_size}mm'
        product_data['Cut Rail Size'] = f'{cut_rail_size}mm'
        product_data['Qty Louvers'] = ''
        
        # Check if louver drop size is not None and add it to product_data
        if louver_drop_size is not None:
            product_data['Louver Drop Size'] = f'{louver_drop_size}mm'
        else:
            product_data['Louver Drop Size'] = 'Invalid Length'

        # Update the Length value for Allusion Blinds
        product_data['Length'] = f'{length}mm'

    return product_data



def calculate_sizes_standard_roller_blinds(product_data):
    # Filter only for Standard Roller Blinds
    if product_data.get('Product Type') == 'Roller Blind' and product_data.get('Roller Type') == 'Standard Roller':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)

        if width is None:
            product_data['Fabric Width'] = 'Invalid Width'
            product_data['Rail Width'] = 'Invalid Width'
            product_data['Fabric Drop'] = 'Invalid Length'
            return product_data

        # Calculate Fabric Width and Rail Width
        fabric_width = rail_width = width - 35

        # Calculate Fabric Drop
        length_str = product_data.get('Length', '0')
        length = convert_to_mm(length_str)
        fabric_drop = length + 300

        # Debug prints
        print(f'Width: {width}')
        print(f'Fabric Width: {fabric_width}')
        print(f'Rail Width: {rail_width}')
        print(f'Length: {length}')
        print(f'Fabric Drop: {fabric_drop}')

        # Add the calculated values to the product_data dictionary
        product_data['Fabric Width'] = f'{fabric_width}mm'
        product_data['Rail Width'] = f'{rail_width}mm'
        product_data['Fabric Drop'] = f'{fabric_drop}mm'

        # Remove unwanted keys
        product_data.pop('Qty Louvers', None)
        product_data.pop('Measurement Protection', None)

    return product_data

def calculate_sizes_cassette_roller_blind(product_data):
    # Filter only for Cassette Roller Blinds
    if product_data.get('Product Type') == 'Roller Blind' and product_data.get('Roller Type') == 'Cassette Roller':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)
        length_str = product_data.get('Length', '0')
        length = convert_to_mm(length_str)

        if width is None or length is None:
            product_data['Cassette Size'] = 'Invalid Width'
            product_data['Tube and Fabric Width'] = 'Invalid Width'
            product_data['Fabric Drop'] = 'Invalid Length'
            return product_data

        measurement_type = product_data.get('Measurement Type')

        if measurement_type == 'Exact':
            cassette_size = width - 4
            tube_and_fabric_width = width - 39
        elif measurement_type == 'Recess':
            cassette_size = width - 14
            tube_and_fabric_width = width - 49

        fabric_drop = length + 300

        product_data['Cassette Size'] = f'{cassette_size}mm'
        product_data['Tube and Fabric Width'] = f'{tube_and_fabric_width}mm'
        product_data['Fabric Drop'] = f'{fabric_drop}mm'

    return product_data


def calculate_sizes_grip_fit_roller_blinds(product_data):
    # Filter only for Grip Fit Roller Blinds
    if product_data.get('Product Type') == 'Grip Fit Roller Blind':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)

        if width is None:
            product_data['Cassette Width'] = 'Invalid Width'
            product_data['Tube and Fabric Width'] = 'Invalid Width'
            product_data['Fabric Drop'] = 'Invalid Length'
            return product_data

        # Calculate Cassette Width, Tube and Fabric Width
        cassette_width = width - 20  # Subtract 20mm from the width
        tube_and_fabric_width = width - 42  # Subtract 42mm from the width

        # Calculate Fabric Drop
        length_str = product_data.get('Length', '0')
        length = convert_to_mm(length_str)
        fabric_drop = length + 300  # Add 300mm to the length

        # Add the calculated values to the product_data dictionary
        product_data['Cassette Width'] = f'{cassette_width}mm'
        product_data['Tube and Fabric Width'] = f'{tube_and_fabric_width}mm'
        product_data['Fabric Drop'] = f'{fabric_drop}mm'

        # Remove unwanted keys
        product_data.pop('Qty Louvers', None)
        product_data.pop('Measurement Protection', None)

    return product_data

def calculate_sizes_day_and_night_blinds(product_data):
    # Filter only for Day & Night Blinds
    if product_data.get('Product Type') == 'Day & Night Blind':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)

        if width is None:
            product_data['Cassette Width'] = 'Invalid Width'
            product_data['Tube and Fabric Width'] = 'Invalid Width'
            product_data['In Bottom'] = 'Invalid Width'
            product_data['Outer Bottom'] = 'Invalid Width'
            return product_data

        measurement_type = product_data.get('Measurement Type')

        if measurement_type == 'Recess':
            cassette_width = width - 14
            tube_and_fabric_width = cassette_width - 35
            in_bottom = cassette_width - 39
            outer_bottom = cassette_width - 27
        elif measurement_type == 'Exact':
            cassette_width = width - 4
            tube_and_fabric_width = cassette_width - 35
            in_bottom = cassette_width - 39
            outer_bottom = cassette_width - 27

        # The drop is the same as the length
        length_str = product_data.get('Length', '0')
        drop = convert_to_mm(length_str)

        # Update product_data with the calculated values
        product_data['Cassette Width'] = f'{cassette_width}mm'
        product_data['Tube and Fabric Width'] = f'{tube_and_fabric_width}mm'
        product_data['In Bottom'] = f'{in_bottom}mm'
        product_data['Outer Bottom'] = f'{outer_bottom}mm'
        product_data['Drop'] = f'{drop}mm'

        # Remove unwanted keys
        product_data.pop('Measurement Protection', None)

    return product_data

def calculate_sizes_perfect_fit_pleated(product_data):
    # Filter only for Perfect Fit Pleated Blinds
    if product_data.get('Product Type') == 'Perfect Fit Pleated':
        width_str = product_data.get('Width', '0')
        width = convert_to_mm(width_str)
        length_str = product_data.get('Length', '0')
        length = convert_to_mm(length_str)

        if width is None or length is None:
            product_data['Fabric Width'] = 'Invalid Width'
            product_data['Fabric Rail'] = 'Invalid Width'
            product_data['Top and Bottom Frame'] = 'Invalid Width'
            product_data['Side Frame'] = 'Invalid Length'
            product_data['Cells'] = ''
            return product_data

        # Calculate Fabric Width and Fabric Rail
        fabric_width = fabric_rail = width - 16

        # Calculate Top and Bottom Frame
        top_bottom_frame = width - 28

        # Calculate Side Frame
        side_frame = length - 28

        # Add the calculated values to the product_data dictionary
        product_data['Fabric Width'] = f'{fabric_width}mm'
        product_data['Fabric Rail'] = f'{fabric_rail}mm'
        product_data['Top and Bottom Frame'] = f'{top_bottom_frame}mm'
        product_data['Side Frame'] = f'{side_frame}mm'
        product_data['Cells'] = ''  # Placeholder for manual entry

        # Remove unwanted keys
        product_data.pop('Measurement Protection', None)

    return product_data










def main():
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, 10)

    driver.get("https://www.emeraldblindsandcurtains.co.uk/z-admin/login/")
    username_field = driver.find_element(By.CLASS_NAME, "form-control[name='email']")
    password_field = driver.find_element(By.CLASS_NAME, "form-control[name='password']")
    username_field.send_keys("shaun_mcgrath451@btinternet.com")
    password_field.send_keys("zBURS0MzzJ@gwTyiLzGIHgObkChm")
    password_field.send_keys(Keys.RETURN)



    specific_order_url = "https://www.emeraldblindsandcurtains.co.uk/z-admin/orders/view/4105/"
    driver.get(specific_order_url)

    wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@class="customer-description"]/div[@class="name mb10"]')))
    customer_name = driver.find_element(By.XPATH, '//div[@class="customer-description"]/div[@class="name mb10"]').text.strip()
    order_data = extract_vertical_blind_data(driver, customer_name)

    if order_data is not None:
        df = pd.DataFrame(order_data).drop(columns=['Width', 'Length'], errors='ignore')
        safe_customer_name = customer_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        excel_file_path = f'{safe_customer_name}_extracted_products.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        print("No valid order data extracted.")

    driver.quit()

if __name__ == "__main__":
    main()
