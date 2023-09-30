import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image
import pandas as pd
import sys
import os
from pdf2image import convert_from_path


# Function to update materials from an Excel file
def update_data_from_excel(file_name, sheet_name):
    # Determine if we're running as a script or a frozen exe
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    file_path = os.path.join(application_path, file_name)

    if not os.path.exists(file_path):
        raise Exception(f"The file {file_path} does not exist")

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        raise Exception(f"Failed to read the Excel file: {e}")

    data = {}
    for index, row in df.iterrows():
        data[row['name']] = {
            'mini': row['mini'],
            'from 2000': row['from 2000'],
            'from 5000': row['from 5000'],
            'from 10000': row['from 10000'],
            'mega': row['mega']
        }

    return data


# Define prices for materials and print quality
materials = {
    'solvent': update_data_from_excel('materials.xlsx', 'solvent'),
    'ecosolvent': update_data_from_excel('materials.xlsx', 'ecosolvent'),
    'uv': update_data_from_excel('materials.xlsx', 'uv')
}

# Define fixed services
fixed_services = ['Поклейка зображення на основу']

# Define area services
area_services = ['Ламінація: китайська плівка прозора, біла мат/гл', 'Ламінація: EU прозора гл/мат']

# Define amount services
amount_services = ['Набивання люверсів', 'Зварювання банерної тканини']

# Define other services
other_services = ['порізка', 'макет', 'інше']

# Define additional services
additional_services = update_data_from_excel('materials.xlsx', 'extra')

# Add other services to the additional services
for service in other_services:
    additional_services[service] = {}


# Calculate area function
def calculate_area_manually(height_m, width_m):
    # Calculate area
    area_m2 = height_m * width_m

    return area_m2, f"Height: {height_m:.2f} m, Width: {width_m:.2f} m, Area: {area_m2:.2f} m²"


def calculate_area_from_file(file_path):
    if file_path.endswith('.pdf'):
        total_area_m2 = 0
        area_str = ''
        images = convert_from_path(file_path)

        for i in range(len(images)):
            # Get dimensions in pixels
            image = images[i]
            width, height = image.size

            # Extract DPI and convert to meters (default to 300 DPI if not available)
            dpi = image.info.get('dpi', (300, 300))[0]
            width_m = width / dpi * 0.0254
            height_m = height / dpi * 0.0254

            # Calculate area and update the area string
            area_m2 = width_m * height_m
            total_area_m2 += area_m2
            area_str += f"File: {os.path.basename(file_path)}, Height: {height_m:.2f} m, Width: {width_m:.2f} m, " \
                        f"Area: {area_m2:.2f} m²\n"

        area_str += f"Total area of all images is {total_area_m2:.2f} m²."

        return total_area_m2, area_str
    else:
        # Open image and get dimensions in pixels
        image = Image.open(file_path)
        width, height = image.size

        # Extract DPI and convert to meters (default to 300 DPI if not available)
        dpi = image.info.get('dpi', (300, 300))[0]
        width_m = width / dpi * 0.0254
        height_m = height / dpi * 0.0254

        # Calculate area
        area_m2 = width_m * height_m

        return area_m2, f"File: {os.path.basename(file_path)}, Height: {height_m:.2f} m, Width: {width_m:.2f} m, " \
                        f"Area: {area_m2:.2f} m²"


def calculate_area_from_directory(directory_path):
    # Initialize total area
    total_area_m2 = 0

    # For storing areas of individual files
    area_strs = []

    # List all files in directory
    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)

        # If it is an image file, calculate its area
        if os.path.isfile(file_path) and file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.tif')):
            area_m2, area_str = calculate_area_from_file(file_path)
            total_area_m2 += area_m2
            area_strs.append(area_str)

    # Combine all area strings
    all_area_str = "\n".join(area_strs)

    total_area_str = f"\nDirectory: {os.path.basename(directory_path)}, Total Area: {total_area_m2:.2f} m²"

    return total_area_m2, all_area_str + total_area_str


def calculate_area():
    file_path = file_entry.get()

    # Check if manual entry is selected
    if manual_entry_var.get():
        # Get height and width values from entry fields
        height_m = float(height_entry.get())
        width_m = float(width_entry.get())

        # Calculate area using manual height and width
        area_m2, area_str = calculate_area_manually(height_m, width_m)
    elif os.path.isfile(file_path):
        # Calculate area from file
        area_m2, area_str = calculate_area_from_file(file_path)
    elif os.path.isdir(file_path):
        # Calculate area from directory
        area_m2, area_str = calculate_area_from_directory(file_path)
    else:
        messagebox.showerror("Error",
                             f"Please either select a file, a directory or enter height and width manually")
        return

    area_label.config(text=area_str)

    return area_m2, area_str


def calculate_price():
    # Retrieve input values
    file_path = file_entry.get()
    material = material_var.get()
    print_quality = quality_var.get()
    services = {service: var.get() for service, var in service_vars.items()}
    price_type = price_type_var.get()

    # Check if file exists
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File or directory {file_path} does not exist")
        return
    try:
        # Calculate height, width, and area
        area_m2, area_str = calculate_area()

        # Calculate print cost
        price_per_m2 = materials[print_quality][material][price_type]
        print_cost = area_m2 * price_per_m2

        # Calculate cost of additional services
        service_cost = 0
        service_formula = ""
        for service, checked in services.items():
            if checked:
                if service in other_services:
                    quantity = float(service_entries[service].get())
                    service_cost += quantity
                    service_formula += f" + {quantity:.2f}"
                elif service in amount_services:
                    quantity = float(service_entries[service].get())
                    service_cost += quantity * additional_services[service][price_type]
                    service_formula += f" + {quantity} * {additional_services[service][price_type]:.2f}"
                elif service in area_services:
                    service_cost += area_m2 * additional_services[service][price_type]
                    service_formula += f" + {area_m2:.2f} * {additional_services[service][price_type]:.2f}"
                elif service in fixed_services:
                    service_cost += additional_services[service][price_type]
                    service_formula += f" + {additional_services[service][price_type]:.2f}"

        # Calculate total cost
        total_cost = print_cost + service_cost

        # Display total cost and calculation formula
        formula_label.config(
            text=f"Calculation formula: ({area_m2:.2f} m² * {price_per_m2:.2f} UAH/m²)"
                 f"{service_formula} =")
        cost_label.config(text=f"Total cost: {total_cost:.2f} UAH")
    except IOError as e:
        print(e)
        messagebox.showerror("Error", f"Could not open file {file_path}")
    except KeyError as e:
        print(e)
        messagebox.showerror("Error", f"Unrecognized service in {services}")
    except ValueError as e:
        print(e)
        messagebox.showerror("Error", f"Something is wrong: {str(e)[:100]}...")


def browse_file():
    file_path = filedialog.askopenfilename()
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)


def browse_directory():
    dir_path = filedialog.askdirectory()
    file_entry.delete(0, tk.END)
    file_entry.insert(0, dir_path)


def enable_manual_entry():
    height_entry.config(state=tk.NORMAL)
    width_entry.config(state=tk.NORMAL)


# Function to update material options based on selected print quality
def update_material_options(*args):
    # Clear the current options
    material_option['menu'].delete(0, 'end')

    # Clear the current selection
    material_var.set('Select material')

    # Get the new options
    new_choices = materials[quality_var.get()].keys()

    # Add the new options to the option menu
    for choice in new_choices:
        material_option['menu'].add_command(label=choice, command=tk._setit(material_var, choice))


def clear_inputs():
    # Clear entry fields
    file_entry.delete(0, 'end')
    height_entry.delete(0, 'end')
    width_entry.delete(0, 'end')

    # Reset dropdown selections
    quality_var.set('Select print quality')
    material_var.set('Select material')
    price_type_var.set('mini')

    # Uncheck manual entry option
    manual_entry_var.set(0)
    height_entry.config(state='disabled')
    width_entry.config(state='disabled')

    # Reset checkboxes and their entries
    for service, var in service_vars.items():
        var.set(0)
        if service in service_entries:
            service_entries[service].delete(0, 'end')
            service_entries[service].config(state='disabled')

    # Clear labels
    formula_label.config(text="")
    cost_label.config(text="")
    area_label.config(text="")


if __name__ == '__main__':

    # Create main window
    root = tk.Tk()

    # Create input fields
    tk.Label(root, text="File/Folder:").grid(row=0, column=0)
    file_entry = tk.Entry(root, width=50)
    file_entry.grid(row=0, column=1)
    tk.Button(root, text="Browse File", command=browse_file).grid(row=0, column=2)
    tk.Button(root, text="Browse Folder", command=browse_directory).grid(row=0, column=3)

    # Create the label and option menu for print quality
    tk.Label(root, text="Print quality:").grid(row=1, column=0)
    quality_var = tk.StringVar()
    quality_option = tk.OptionMenu(root, quality_var, *materials.keys())
    quality_option.grid(row=1, column=1)

    # Create the label and option menu for material
    tk.Label(root, text="Material:").grid(row=2, column=0)
    material_var = tk.StringVar()
    material_option = tk.OptionMenu(root, material_var, *materials['solvent'].keys())  # initial values
    material_option.grid(row=2, column=1)

    # Call update_material_options function whenever quality_var changes
    quality_var.trace('w', update_material_options)

    # Price type selection
    tk.Label(root, text="Price type:").grid(row=3, column=0)
    price_type_var = tk.StringVar(value='mini')
    price_types = ['mini', 'from 2000', 'from 5000', 'from 10000', 'mega']
    price_type_option = tk.OptionMenu(root, price_type_var, *price_types)
    price_type_option.grid(row=3, column=1)

    # Additional services selection
    tk.Label(root, text="Services:").grid(row=4, column=0)
    service_vars = {service: tk.BooleanVar() for service in additional_services.keys()}
    service_entries = {}


    def update_service_entries():
        for service, var in service_vars.items():
            if var.get() and service in service_entries:
                service_entries[service].config(state='normal')
            elif not var.get() and service in service_entries:
                service_entries[service].config(state='disabled')


    # Additional services selection
    for i, (service, var) in enumerate(service_vars.items()):
        cb = tk.Checkbutton(root, text=service, variable=var, command=update_service_entries)
        cb.grid(row=4 + i, column=1)
        if service in amount_services or service in other_services:
            service_entries[service] = tk.Entry(root, state='disabled')
            service_entries[service].grid(row=4 + i, column=2)

    # Create calculate button
    calc_button = tk.Button(root, text="Calculate", command=calculate_price)
    calc_button.grid(row=5 + len(service_vars), column=0, columnspan=3)

    # Create label to display calculation formula
    formula_label = tk.Label(root, text="")
    formula_label.grid(row=6 + len(service_vars), column=0, columnspan=3)

    # Create label to display total cost
    cost_label = tk.Label(root, text="")
    cost_label.grid(row=7 + len(service_vars), column=0, columnspan=3)

    # Create manual entry option
    manual_entry_var = tk.BooleanVar()
    manual_entry_checkbox = tk.Checkbutton(root, text="Enter height and width manually", variable=manual_entry_var,
                                           command=enable_manual_entry)
    manual_entry_checkbox.grid(row=8 + len(service_vars), column=0, columnspan=3)

    # Create height and width entry fields
    tk.Label(root, text="Height (m):").grid(row=9 + len(service_vars), column=0)
    height_entry = tk.Entry(root, width=10, state=tk.DISABLED)
    height_entry.grid(row=9 + len(service_vars), column=1)

    tk.Label(root, text="Width (m):").grid(row=9 + len(service_vars), column=2)
    width_entry = tk.Entry(root, width=10, state=tk.DISABLED)
    width_entry.grid(row=9 + len(service_vars), column=3)

    # Create calculate area button
    calc_area_button = tk.Button(root, text="Calculate Area", command=calculate_area)
    calc_area_button.grid(row=10 + len(service_vars), column=0, columnspan=3)

    # Create label to display height, width, and total area
    area_label = tk.Label(root, text="")
    area_label.grid(row=11 + len(service_vars), column=0, columnspan=3)

    # Create clear button
    clear_button = tk.Button(root, text="Clear", command=clear_inputs)
    clear_button.grid(row=12 + len(service_vars), column=0, columnspan=3)

    # Run main event loop
    root.mainloop()
