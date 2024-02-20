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
