import PySimpleGUI as sg
import pandas as pd
import json
import os

class Product:
    product_counter = 1

    def __init__(self, model, serial_number, hard_drive_details, cosmetic_condition, is_desktop, is_hard_drive_wiped):
        self.id = f"Asset-{Product.product_counter}"
        Product.product_counter += 1
        self.model = model
        self.serial_number = serial_number
        self.hard_drive_details = hard_drive_details
        self.cosmetic_condition = cosmetic_condition
        self.is_desktop = is_desktop
        self.is_hard_drive_wiped = is_hard_drive_wiped

class Job:
    job_counter = 1

    def __init__(self):
        self.job_id = f"JOB-{Job.job_counter}"
        Job.job_counter += 1
        self.products = []

    def add_product(self, product):
        self.products.append(product)

    def close_job(self):
        return self.generate_report()

    def generate_report(self):
        total_count = len(self.products)
        wiped_drives_count = sum(p.is_hard_drive_wiped for p in self.products)
        total_cost = sum(self.calculate_cost(p) for p in self.products)
        
        report = {
            'job_id': self.job_id,
            'total_count': total_count,
            'wiped_drives_count': wiped_drives_count,
            'total_cost': total_cost,
            'products': self.products
        }
        return report

    def calculate_cost(self, product):
        cost = 0
        if product.is_desktop:
            cost += 5
        if product.is_hard_drive_wiped:
            cost += 7.5
        return cost

def export_to_excel(inventory, job_id):
    data = []
    for job in inventory:
        for product in job['products']:
            data.append({
                'Job ID': job['job_id'],
                'Product ID': product.id,
                'Model': product.model,
                'Serial Number': product.serial_number,
                'Hard Drive Details': product.hard_drive_details,
                'Cosmetic Condition': product.cosmetic_condition,
                'Is Desktop': product.is_desktop,
                'Is Hard Drive Wiped': product.is_hard_drive_wiped,
                'Cost': job['total_cost']
            })
    df = pd.DataFrame(data)
    
    # Determine the export number
    export_number = 1
    while os.path.exists(f'{job_id}_export_{export_number}.xlsx'):
        export_number += 1
    
    # Save the file with the appropriate name
    filename = f'{job_id}_export_{export_number}.xlsx'
    df.to_excel(filename, index=False)
    print(f"Exported to {filename}")

def load_details():
    if os.path.exists('details.json'):
        with open('details.json', 'r') as file:
            return json.load(file)
    return {
        'model': [],
        'serial_number': [],
        'hard_drive_details': [],
        'cosmetic_condition': []
    }

def save_details(details):
    with open('details.json', 'w') as file:
        json.dump(details, file)

def load_job_history():
    if os.path.exists('job_history.json'):
        with open('job_history.json', 'r') as file:
            return json.load(file)
    return []

def save_job_history(job_history):
    with open('job_history.json', 'w') as file:
        json.dump(job_history, file)

def load_asset_history():
    if os.path.exists('asset_history.json'):
        with open('asset_history.json', 'r') as file:
            return json.load(file)
    return []

def save_asset_history(asset_history):
    with open('asset_history.json', 'w') as file:
        json.dump(asset_history, file)

def main():
    details = load_details()
    job_history = load_job_history()
    asset_history = load_asset_history()

    # Ensure unique job and asset counters
    if job_history:
        Job.job_counter = max(int(job['job_id'].split('-')[1]) for job in job_history) + 1
    if asset_history:
        Product.product_counter = max(int(asset['id'].split('-')[1]) for asset in asset_history) + 1

    # Define the layout for each tab
    contracts_layout = [
        [sg.Text('Pricing Information')],
        [sg.Text('Desktop Charge'), sg.InputText('5', key='desktop_charge')],
        [sg.Text('Hard Drive Wipe Charge'), sg.InputText('7.5', key='hdd_wipe_charge')],
    ]

    audit_layout = [
        [sg.Text('Model'), sg.InputCombo(details['model'], size=(20, 1), key='model'), sg.Button('+', key='add_model')],
        [sg.Text('Serial Number'), sg.InputCombo(details['serial_number'], size=(20, 1), key='serial_number'), sg.Button('+', key='add_serial_number')],
        [sg.Text('Hard Drive Details'), sg.InputCombo(details['hard_drive_details'], size=(20, 1), key='hard_drive_details'), sg.Button('+', key='add_hard_drive_details')],
        [sg.Text('Cosmetic Condition'), sg.InputCombo(details['cosmetic_condition'], size=(20, 1), key='cosmetic_condition'), sg.Button('+', key='add_cosmetic_condition')],
        [sg.Checkbox('Is Desktop', key='is_desktop')],
        [sg.Checkbox('Is Hard Drive Wiped', key='is_hard_drive_wiped')],
        [sg.Button('Add Product')],
        [sg.Output(size=(50, 10))]
    ]

    inventory_layout = [
        [sg.Text('Previous Audits')],
        [sg.Listbox(values=[], size=(50, 10), key='audit_list')],
        [sg.Text('Filter by Job ID'), sg.InputText(key='filter_job_id')],
        [sg.Text('Filter by Serial Number'), sg.InputText(key='filter_serial_number')],
        [sg.Button('Apply Filter')],
        [sg.Button('Export to Excel')]
    ]

    # Combine the layouts into tabs
    layout = [
        [sg.TabGroup([
            [sg.Tab('Contracts', contracts_layout),
             sg.Tab('Audit', audit_layout),
             sg.Tab('Inventory', inventory_layout)]
        ])],
        [sg.Button('Close Job')]
    ]

    window = sg.Window('Product Entry', layout)
    job = Job()
    inventory = []

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == 'Add Product':
            product = Product(
                model=values['model'],
                serial_number=values['serial_number'],
                hard_drive_details=values['hard_drive_details'],
                cosmetic_condition=values['cosmetic_condition'],
                is_desktop=values['is_desktop'],
                is_hard_drive_wiped=values['is_hard_drive_wiped']
            )
            job.add_product(product)
            asset_history.append({
                'id': product.id,
                'model': product.model,
                'serial_number': product.serial_number,
                'hard_drive_details': product.hard_drive_details,
                'cosmetic_condition': product.cosmetic_condition,
                'is_desktop': product.is_desktop,
                'is_hard_drive_wiped': product.is_hard_drive_wiped
            })
            save_asset_history(asset_history)
            print(f"Added: {product.model}, {product.serial_number}, UID: {product.id}")

        if event == 'Close Job':
            report = job.close_job()
            inventory.append(report)
            job_history.append(report)
            save_job_history(job_history)
            window['audit_list'].update(values=[f"Job ID: {r['job_id']}, Total Products: {r['total_count']}" for r in inventory])
            print(f"Job {report['job_id']} closed with {report['total_count']} products.")

        if event == 'Export to Excel':
            if inventory:
                export_to_excel(inventory, job.job_id)
            else:
                print("No jobs to export.")

        if event == 'add_model':
            new_model = sg.popup_get_text('Enter new model')
            if new_model:
                details['model'].append(new_model)
                window['model'].update(values=details['model'])
                save_details(details)

        if event == 'add_serial_number':
            new_serial_number = sg.popup_get_text('Enter new serial number')
            if new_serial_number:
                details['serial_number'].append(new_serial_number)
                window['serial_number'].update(values=details['serial_number'])
                save_details(details)

        if event == 'add_hard_drive_details':
            new_hard_drive_details = sg.popup_get_text('Enter new hard drive details')
            if new_hard_drive_details:
                details['hard_drive_details'].append(new_hard_drive_details)
                window['hard_drive_details'].update(values=details['hard_drive_details'])
                save_details(details)

        if event == 'add_cosmetic_condition':
            new_cosmetic_condition = sg.popup_get_text('Enter new cosmetic condition')
            if new_cosmetic_condition:
                details['cosmetic_condition'].append(new_cosmetic_condition)
                window['cosmetic_condition'].update(values=details['cosmetic_condition'])
                save_details(details)

    window.close()

if __name__ == "__main__":
    main()