from flask import Flask, render_template, request
from pandas import read_excel
from datetime import datetime

import os
import pdfkit   # pdfkit requires that wkhtmltopdf be installed in order to work
import platform
import shutil
import socket
import tempfile


NAME_DICTIONARY = {
    'Fresh Food': 'fresh-food',
    'Freezer Meats': 'freezer-meats',
    'Freezer Bonus': 'freezer-bonus',
    'Fridge': 'fridge',
    'Canned Vegetables': 'canned-veg',
    'Broth': 'broth',
    'Canned Soup': 'canned-soup',
    'Canned Meat': 'canned-meat',
    'Beans & Lentils': 'beans',
    'Juice': 'juice',
    'Shelf-stable Milk': 'up-milk',
    'Snacks': 'snacks',
    'Pantry': 'pantry',
    'Rice': 'rice',
    'Canned Fruit': 'canned-fruit',
    'Pantry 2': 'pantry-2',
    'Breakfast': 'breakfast',
    'Peanut Butter & Jelly': 'pbj',
    'Canned Tomatoes': 'canned-tom',
    'Bonus Items': 'bonus',
    'Bonus Items 2': 'bonus-2',
    'Personal Hygiene Items': 'hygiene',
    'Paper Goods': 'paper',
    'Snack Bags for Kids': 'snack_bags',
    'Diapers & Pull-ups': 'diapers',
    'Formula': 'formula',
    'Baby Food': 'baby-food',
    'Coffee/Tea/Cocoa': 'coffee'
}


if not os.path.isfile(os.path.expanduser('~/Desktop/DriveThruGroceryList.xlsx')):
    if __name__ == '__main__':
        shutil.copy(os.path.dirname(__file__) + '/static/DriveThruGroceryList.xlsx',
                    os.path.expanduser('~/Desktop/DriveThruGroceryList.xlsx'))
    else:
        try:
            import importlib.resources as pkg_resources
        except ImportError:
            # Try backported to PY<37 `importlib_resources`.
            import importlib_resources as pkg_resources
        with pkg_resources.path(__package__, 'server.py') as server_py:
            spreadsheet = os.path.join(os.path.dirname(server_py), 'static', 'DriveThruGroceryList.xlsx')
            shutil.copy(spreadsheet,
                    os.path.expanduser('~/Desktop/DriveThruGroceryList.xlsx'))

DriveThruGroceryList = read_excel(os.path.expanduser('~/Desktop/DriveThruGroceryList.xlsx'), engine='openpyxl')

app = Flask(__name__, static_url_path='/static')


def print_html(html):
    """ Convert HTML markup to PDF and then send the PDF to the default printer.
        Windows is unique in that it has no support on it's own for printing PDF
        files, so users of Windows must install PDFtoPrinter from this URL:
        http://www.columbia.edu/~em36/PDFtoPrinter.exe
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, 'packing_list.pdf')
        pdfkit.from_string(html, pdf_path, options={'page-size': 'Letter',
                                                    'zoom': '1.22',
                                                    'margin-bottom': '0',
                                                    'margin-left': '5',
                                                    'margin-right': '2'})
        operating_system = platform.system()

        if operating_system in ['Darwin', 'Linux']:  # send ot printer on Mac
            os.system('lp "{}"'.format(pdf_path))
            # os.system('cp "{}" ~/Desktop/packing_list.pdf && open ~/Desktop/packing_list.pdf'.format(pdf_path))
        elif operating_system == 'Windows':
            os.system('PDFtoPrinter.exe "{}"'.format(pdf_path))


def my_ip_address():
    """ Discover the current IP address (other than localhost) of this machine. """
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't even have to be reachable
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip


@app.template_filter('shortname')
def shortname(section):
    """ This is a Jinja2 filter that returns a short name for each grocery
        list section (the section names must not change). """
    return NAME_DICTIONARY[section]


@app.template_filter('simplify')
def simplify(stringlist):
    """ This is a Jinja2 filter that takes a list of strings and concatenates them. """
    return ', '.join(stringlist)


@app.route('/')
def form():
    """ Get the grocery list form. """
    return render_template('order_form.html', grocery_options=DriveThruGroceryList)


@app.route('/print', methods=['POST'])
def print_form():
    """ Receive the grocery list, prepare a packing list, and print it. """
    fam_color = {'1: Yellow': '#ffff00', '2-4: Blue': '#6464ff', '5+: Pink': '#ff69b4'}[request.form['family_size']]
    grocery_list = request.form.to_dict(flat=False)
    packing_list = render_template('packing_list.html',
                                   grocery_list=grocery_list,
                                   timestamp=datetime.now().strftime('%Y-%m-%dT%H:%M:%S.%f%z'),
                                   fam_color=fam_color)
    print_html(packing_list)

    return "Success"


@app.route('/reprint', methods=['GET', 'POST'])
def reprint_form():
    """ (Unimplemented) Reprint a previous list. """
    return 'This feature is currently not implemented.'


if __name__ == '__main__':
    app.run(host=my_ip_address())
