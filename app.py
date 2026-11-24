from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import json
import re
import unicodedata

ROOT = Path(__file__).parent
PORT = 8765
CACHE_VERSION = 20
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def cell_value(cell, shared):
    if cell.attrib.get('t') == 'inlineStr':
        return ''.join(node.text or '' for node in cell.iterfind('.//m:t', NS)).strip()
    value = cell.find('m:v', NS)
    raw = '' if value is None else (value.text or '')
    return shared[int(raw)] if cell.attrib.get('t') == 's' and raw else raw

def excel_rows(path):
    with ZipFile(path) as book:
        shared = []
        if 'xl/sharedStrings.xml' in book.namelist():
            root = ET.fromstring(book.read('xl/sharedStrings.xml'))
            shared = [''.join(t.text or '' for t in item.iterfind('.//m:t', NS))
                      for item in root.findall('m:si', NS)]
        workbook = ET.fromstring(book.read('xl/workbook.xml'))
        rels_root = ET.fromstring(book.read('xl/_rels/workbook.xml.rels'))
        rels = {item.attrib['Id']: item.attrib['Target'].lstrip('/') for item in rels_root}
        sheets = []
        for sheet in workbook.find('m:sheets', NS):
            key = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
            target = rels[sheet.attrib[key]]
            target = target if target.startswith('xl/') else 'xl/' + target
            xml = ET.fromstring(book.read(target))
            rows = []
            for row in xml.findall('.//m:sheetData/m:row', NS):
                values = {}
                for cell in row.findall('m:c', NS):
                    match = re.match(r'([A-Z]+)', cell.attrib.get('r', ''))
                    if match:
                        values[match.group(1)] = cell_value(cell, shared)
                if values:
                    rows.append(values)
            sheets.append((sheet.attrib['name'], rows))
        return sheets

def accounting_summary():
    paths, account_export = accounting_export_paths()
    by_year = {}
    monthly_net_profit = {}
    covered_periods = set()
    for path in paths:
        sheets = excel_rows(path)
        file_periods = {date_label(row.get('A')) for _, rows in sheets for row in rows[1:]}
        allowed_periods = file_periods - covered_periods
        if path != account_export and not allowed_periods:
            continue
        for _, rows in sheets:
            for row in rows[1:]:
                period = date_label(row.get('A'))
                if path != account_export and period not in allowed_periods:
                    continue
                account = str(row.get('E', '')).strip()
                if not account or not account[0] in '56':
                    continue
                amount = number(row.get('L'))
                prefix = account[:2]
                year = period[:4]
                totals = by_year.setdefault(year, {'revenue': 0, 'expense': 0,
                                                    'depreciation': 0, 'interest': 0,
                                                    'subsidies': 0})
                if prefix in ('50', '51', '52', '53', '54', '55', '56', '57', '58', '59'):
                    totals['expense'] += amount
                elif prefix in ('60', '61', '62', '63', '64', '65', '66', '67', '68', '69'):
                    totals['revenue'] -= amount
                if prefix == '55':
                    totals['depreciation'] += amount
                elif prefix == '56':
                    totals['interest'] += amount
                elif prefix == '64':
                    totals['subsidies'] += amount
                if period.startswith('2026-'):
                    monthly_net_profit[period] = monthly_net_profit.get(period, 0) - amount
        covered_periods.update(file_periods)
    current_year = max(by_year, default=str(datetime.now().year))
    totals = by_year.get(current_year, {'revenue': 0, 'expense': 0, 'depreciation': 0,
                                        'interest': 0, 'subsidies': 0})
    totals['operating_profit'] = (totals['revenue'] - totals['expense']
                                  + totals['depreciation'] + totals['interest']
                                  + totals['subsidies'])
    totals['net_profit_after_adjustments'] = totals['revenue'] - totals['expense']
    totals['monthly_net_profit'] = monthly_net_profit
    return totals

def report_summary():
    path = next((item for item in ROOT.glob('*.xlsx')
                 if unicodedata.normalize('NFC', item.name).startswith('Október')
                 and 'rozbor' in unicodedata.normalize('NFC', item.name).lower()), None)
    if not path:
        return {}
    for sheet_name, rows in excel_rows(path):
        if sheet_name != 'FINZÁKLAD':
            continue
        header = next((row for row in rows
                       if any(date_label(value).startswith('2026-') for value in row.values())), None)
        if not header:
            return {'file': path.name}
        month_columns = {date_label(value): column
                         for column, value in header.items()
                         if date_label(value).startswith('2026-')
                         and date_label(value) >= '2026-01'
                         and date_label(value) <= '2026-10'}
        ytd = {}
        for row in rows:
            label = row.get('B', '').strip()
            if label in ('Prev zisk/strata', 'Zisk/strata', 'Odpisy', 'Úroky', 'Dotácie'):
                ytd[label] = number(row.get('R'))
        months = {}
        for period, month_column in month_columns.items():
            values = {}
            for row in rows:
                label = row.get('B', '').strip()
                if label in ('Predaj', 'Tržby', 'Priame náklady', 'Zisk/strata',
                             'Doprava', 'Enviro náklady', 'Spolu rež. náklady',
                             'Prev zisk/strata', 'Odpisy', 'Úroky', 'Dotácie'):
                    values.setdefault(label, number(row.get(month_column)))
            months[period] = {
                'revenue': values.get('Tržby', 0),
                'cost': abs(values.get('Priame náklady', 0)),
                'quantity': abs(values.get('Predaj', 0)),
                'net_profit': values.get('Zisk/strata', 0),
                'operating_profit': values.get('Prev zisk/strata', 0),
                'direct_cost': abs(values.get('Priame náklady', 0)),
                'transport': abs(values.get('Doprava', 0)),
                'environment': abs(values.get('Enviro náklady', 0)),
                'overhead': abs(values.get('Spolu rež. náklady', 0)),
                'depreciation': abs(values.get('Odpisy', 0)),
                'interest': values.get('Úroky', 0),
                'subsidies': values.get('Dotácie', 0),
            }
        october = months.get('2026-10', {})
        return {
            'file': path.name,
            'months': months,
            'ytd_operating_profit': ytd.get('Prev zisk/strata', 0),
            'ytd_net_profit': ytd.get('Zisk/strata', 0),
            'ytd_depreciation': abs(ytd.get('Odpisy', 0)),
            'ytd_interest': ytd.get('Úroky', 0),
            'ytd_subsidies': ytd.get('Dotácie', 0),
            **october,
        }
    return {'file': path.name}

def number(value):
    try:
        return float(str(value).replace(',', '.').replace(' ', ''))
    except (ValueError, TypeError):
        return 0.0

def date_label(value):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime('%Y-%m')
    except (ValueError, TypeError):
        return 'Neznámy mesiac'

def accounting_export_paths():
    account_export = next((path for path in ROOT.glob('*.xlsx')
                           if 'podľa účtov' in unicodedata.normalize('NFC', path.name)), None)
    if account_export:
        supplements = sorted(
            (path for path in ROOT.glob('*.xlsx')
             if path != account_export
             and unicodedata.normalize('NFC', path.name).startswith('Vecné položky')
             and path.stat().st_mtime_ns > account_export.stat().st_mtime_ns),
            key=lambda path: path.stat().st_mtime_ns,
            reverse=True)
        return [account_export, *supplements], account_export
    paths = []
    for path in ROOT.glob('*.xlsx'):
        name = unicodedata.normalize('NFC', path.name)
        if name.startswith('Vecné položky 2021') or name.startswith('Vecné položky 07'):
            paths.append(path)
    return paths, None

def dimensions():
    ranges = []
    path = next(ROOT.glob('Dimenzie *.xlsx'), None)
    if not path:
        return ranges
    for sheet_name, rows in excel_rows(path):
        if sheet_name.lower() == 'cover':
            continue
        category = next((row.get('B', '') for row in rows[:5] if row.get('B')), sheet_name)
        for row in rows[7:]:
            code = row.get('B', '').strip()
            if code.isdigit():
                ranges.append({'code': code, 'name': row.get('C', '').strip() or code,
                               'category': category})
    return ranges

def category_for(account, ranges):
    code = ''.join(ch for ch in str(account) if ch.isdigit())
    for item in ranges:
        if code == item['code']:
            return item['category']
    prefixes = {'11': 'Opravy a údržba', '22': 'Consumer marketing', '33': 'Trade marketing',
                '44': 'Trade marketing - letákové akcie', '55': 'Logistika a doprava',
                '66': 'Réžia', '77': 'Materiál a výroba'}
    return next((name for prefix, name in prefixes.items() if code.startswith(prefix)), 'Nezaradené')

def customer_directory():
    directory = {}
    path = next((item for item in ROOT.glob('*.xlsx')
                 if unicodedata.normalize('NFC', item.name).startswith('Zoznam zákazníkov')), None)
    if not path:
        return directory
    for _, rows in excel_rows(path):
        for row in rows[1:]:
            code = row.get('A', '').strip()
            name = row.get('B', '').strip()
            if code and name:
                directory[code] = name
    return directory

def _load_data():
    ranges, categories, months, suppliers, customers, files = dimensions(), {}, {}, {}, {}, []
    rozbor = accounting_summary()
    report = report_summary()
    for period, report_month in report.get('months', {}).items():
        rozbor.setdefault('monthly_net_profit', {})[period] = report_month['net_profit']
    if report.get('ytd_net_profit') is not None:
        rozbor['net_profit_after_adjustments'] = report['ytd_net_profit']
    if report.get('ytd_operating_profit') is not None:
        rozbor['operating_profit'] = report['ytd_operating_profit']
    if report.get('ytd_depreciation') is not None:
        rozbor['depreciation'] = report['ytd_depreciation']
    if report.get('ytd_interest') is not None:
        rozbor['interest'] = report['ytd_interest']
    if report.get('ytd_subsidies') is not None:
        rozbor['subsidies'] = report['ytd_subsidies']
    account_paths, account_export = accounting_export_paths()
    account_path_order = {path: index for index, path in enumerate(account_paths)}
    periods = {}
    categories_by_year = {}
    customer_profit = {}
    customer_revenue = {}
    customer_cost = {}
    customer_quantity = {}
    customer_period_sales = {}
    product_period_sales = {}
    product_names = {}
    delivery_items = {}
    delivery_costs = {}
    customer_names = customer_directory()
    covered_account_periods = set()
    paths = sorted(ROOT.glob('*.xlsx'), key=lambda path: account_path_order.get(path, len(account_path_order)))
    for path in paths:
        normalized_name = unicodedata.normalize('NFC', path.name)
        if normalized_name.startswith('VZOR') or normalized_name.startswith('Dimenzie'):
            continue
        is_account_source = normalized_name.startswith('Vecné položky')
        if account_export and is_account_source and path not in account_paths:
            continue
        try:
            sheets = excel_rows(path)
        except Exception:
            continue
        if not sheets:
            continue
        file_periods = {date_label(row.get('A')) for _, rows in sheets for row in rows[1:]}
        allowed_account_periods = file_periods - covered_account_periods
        if is_account_source and account_export and path != account_export and not allowed_account_periods:
            continue
        files.append({'name': path.name, 'rows': max(0, len(sheets[0][1]) - 1),
                      'kind': path.stem.split(' 07')[0]})
        for _, rows in sheets:
            if not rows:
                continue
            for row in rows[1:]:
                if normalized_name.startswith('Vecné'):
                    if (account_export and path != account_export
                            and date_label(row.get('A')) not in allowed_account_periods):
                        continue
                    account = str(row.get('E', '')).strip()
                    amount = number(row.get('J'))
                    if not amount or not account.startswith('5'):
                        continue
                    account_labels = {'50': 'Spotreba materiálu a tovaru', '51': 'Služby',
                                      '52': 'Osobné náklady', '53': 'Dane a poplatky',
                                      '54': 'Iné prevádzkové náklady', '55': 'Odpisy',
                                      '56': 'Finančné náklady'}
                    label = account_labels.get(account[:2], f'Účty {account[:2]}x')
                    categories[label] = categories.get(label, 0) + amount
                    month = date_label(row.get('A'))
                    months[month] = months.get(month, 0) + amount
                    year_categories = categories_by_year.setdefault(month[:4], {})
                    year_categories[label] = year_categories.get(label, 0) + amount
                elif normalized_name.startswith('Položky dodávateľa'):
                    amount = abs(number(row.get('J')))
                    label = row.get('F', '').strip() or 'Bez popisu'
                    suppliers[label] = suppliers.get(label, 0) + amount
                elif normalized_name.startswith('Položky zákazníka'):
                    label = row.get('D', '').strip() or 'Bez čísla zákazníka'
                    if row.get('B', '').strip() == 'Platba':
                        description = row.get('F', '').strip()
                        description = re.sub(r'^(preplatok|preplatok zákazníka)-', '', description, flags=re.IGNORECASE)
                        description = re.sub(r',\s*\d{7,}$', '', description).strip()
                        if description and label not in customer_names and not description.lower().startswith(('zápočet', 'platba')):
                            customer_names[label] = description
                        continue
                    if row.get('B', '').strip() not in ('Faktúra', 'Dobropis'):
                        continue
                    amount = number(row.get('H'))
                    customers[label] = customers.get(label, 0) + amount
                elif normalized_name.startswith('Položky tovaru'):
                    if row.get('B', '').strip() != 'Predaj':
                        continue
                    raw_customer = row.get('G', '').strip()
                    customer = raw_customer or 'Bez priradeného zákazníka'
                    delivery_id = row.get('D', '').strip()
                    customer_code = raw_customer[:5]
                    revenue = number(row.get('L'))
                    cost = number(row.get('M'))
                    period = date_label(row.get('A'))
                    product_code = row.get('E', '').strip()
                    product_name = row.get('F', '').strip()
                    product_id = product_code or product_name or 'Bez priradeného produktu'
                    if product_name:
                        product_names.setdefault(product_id, product_name)
                    product_entry = product_period_sales.setdefault(period, {}).setdefault(
                        product_id, {'revenue': 0, 'cost': 0, 'quantity': 0, 'transport': 0})
                    product_entry['revenue'] += revenue
                    product_entry['cost'] += cost
                    product_entry['quantity'] += abs(number(row.get('J')))
                    entry = periods.setdefault(period, {'revenue': 0, 'cost': 0, 'quantity': 0})
                    entry['revenue'] += revenue
                    entry['cost'] += cost
                    entry['quantity'] += abs(number(row.get('J')))
                    customer_revenue[customer] = customer_revenue.get(customer, 0) + revenue
                    customer_cost[customer] = customer_cost.get(customer, 0) + cost
                    customer_profit[customer] = customer_profit.get(customer, 0) + revenue + cost
                    customer_quantity[customer] = customer_quantity.get(customer, 0) + abs(number(row.get('J')))
                    period_customer = customer_period_sales.setdefault(period, {})
                    period_entry = period_customer.setdefault(
                        customer_code or 'Bez priradeného zákazníka',
                        {'revenue': 0, 'cost': 0, 'quantity': 0, 'transport': 0})
                    period_entry['revenue'] += revenue
                    period_entry['cost'] += cost
                    period_entry['quantity'] += abs(number(row.get('J')))
                    if delivery_id:
                        delivery = delivery_items.setdefault(
                            delivery_id, {'customer_codes': set(), 'revenue': 0,
                                          'cost': 0, 'quantity': 0, 'periods': {},
                                          'products': {}})
                        if customer_code:
                            delivery['customer_codes'].add(customer_code)
                        delivery['revenue'] += revenue
                        delivery['cost'] += cost
                        delivery['quantity'] += abs(number(row.get('J')))
                        delivery_period = delivery['periods'].setdefault(
                            period, {'revenue': 0, 'cost': 0, 'quantity': 0})
                        delivery_period['revenue'] += revenue
                        delivery_period['cost'] += cost
                        delivery_period['quantity'] += abs(number(row.get('J')))
                        delivery_product = delivery['products'].setdefault(
                            period, {}).setdefault(product_id, {'revenue': 0})
                        delivery_product['revenue'] += revenue
                    if customer_code in customer_names:
                        customer_names[customer] = customer_names[customer_code]
                elif normalized_name.startswith('Účtované predajné dodávky'):
                    delivery_id = row.get('A', '').strip()
                    if not delivery_id:
                        continue
                    delivery = delivery_costs.setdefault(
                        delivery_id, {'customer': row.get('F', '').strip(),
                                       'name': row.get('I', '').strip(),
                                       'transport': 0})
                    delivery['transport'] += number(row.get('Z'))
        if is_account_source and account_export:
            covered_account_periods.update(file_periods)
    for customer in customer_profit:
        customer_names.setdefault(customer, customer_names.get(customer[:5], f'Zákazník {customer[:5]}'))
    deliveries = []
    delivery_transport_by_customer = {}
    unmatched_delivery_count = 0
    unmatched_delivery_transport = 0
    mismatched_delivery_count = 0
    mismatched_delivery_transport = 0
    for delivery_id, delivery_cost in delivery_costs.items():
        items = delivery_items.get(delivery_id)
        if not items:
            unmatched_delivery_count += 1
            unmatched_delivery_transport += delivery_cost['transport']
            continue
        for period, products in items['products'].items():
            period_revenue = sum(abs(value['revenue']) for value in products.values())
            for product_id, value in products.items():
                product_period_sales[period][product_id]['transport'] += (
                    delivery_cost['transport'] * abs(value['revenue']) / period_revenue
                    if period_revenue else 0)
        shipment_customer = delivery_cost['customer']
        item_customers = items['customer_codes']
        customer_matches = (len(item_customers) == 1
                            and shipment_customer in item_customers)
        if customer_matches:
            delivery_transport_by_customer[shipment_customer] = (
                delivery_transport_by_customer.get(shipment_customer, 0)
                + delivery_cost['transport'])
            delivery_revenue_by_period = {
                period: abs(values['revenue'])
                for period, values in items['periods'].items()
            }
            total_period_revenue = sum(delivery_revenue_by_period.values())
            for period, period_revenue in delivery_revenue_by_period.items():
                period_transport = (
                    delivery_cost['transport'] * period_revenue / total_period_revenue
                    if total_period_revenue else 0)
                customer_period_sales.setdefault(period, {}).setdefault(
                    shipment_customer,
                    {'revenue': 0, 'cost': 0, 'quantity': 0, 'transport': 0}
                )['transport'] += period_transport
            if delivery_cost['name']:
                customer_names.setdefault(shipment_customer, delivery_cost['name'])
        else:
            mismatched_delivery_count += 1
            mismatched_delivery_transport += delivery_cost['transport']
        item_customer = next(iter(item_customers)) if len(item_customers) == 1 else ''
        deliveries.append({
            'number': delivery_id,
            'customer': shipment_customer,
            'customer_name': delivery_cost['name'] or customer_names.get(
                shipment_customer, f'Zákazník {shipment_customer}'),
            'item_customer': item_customer,
            'customer_matches': customer_matches,
            'revenue': items['revenue'],
            'direct_cost': abs(items['cost']),
            'gross_profit': items['revenue'] + items['cost'],
            'transport': delivery_cost['transport'],
            'profit_after_transport': (
                items['revenue'] + items['cost'] - delivery_cost['transport']
                if customer_matches else None),
        })
    deliveries.sort(key=lambda item: item['number'])
    for period, report_month in report.get('months', {}).items():
        entry = periods.setdefault(period, {'revenue': 0, 'cost': 0, 'quantity': 0})
        entry['revenue'] = report_month['revenue']
        entry['cost'] = report_month['cost']
        entry['quantity'] = report_month['quantity']
    customer_period_data = {}
    for period, customer_values in customer_period_sales.items():
        revenue_total = sum(value['revenue'] for value in customer_values.values())
        transport_total = sum(value['transport'] for value in customer_values.values())
        gross_after_transport = sum(
            value['revenue'] + value['cost'] - value['transport']
            for value in customer_values.values())
        if period in report.get('months', {}):
            shared_cost = gross_after_transport - report['months'][period]['net_profit']
        elif period in months:
            direct_cost_total = abs(sum(value['cost'] for value in customer_values.values()))
            shared_cost = months[period] - direct_cost_total - transport_total
        else:
            continue
        customer_period_data[period] = {}
        for customer, value in customer_values.items():
            allocated_shared_cost = (
                shared_cost * value['revenue'] / revenue_total if revenue_total else 0)
            customer_period_data[period][customer] = {
                'revenue': value['revenue'],
                'cost': abs(value['cost']),
                'profit': value['revenue'] + value['cost'],
                'transport': value['transport'],
                'shared': allocated_shared_cost,
                'net': (value['revenue'] + value['cost'] - value['transport']
                        - allocated_shared_cost),
                'quantity': value['quantity'],
            }
    product_period_data = {}
    for period, product_values in product_period_sales.items():
        revenue_total = sum(value['revenue'] for value in product_values.values())
        transport_total = sum(value['transport'] for value in product_values.values())
        gross_after_transport = sum(
            value['revenue'] + value['cost'] - value['transport']
            for value in product_values.values())
        if period in report.get('months', {}):
            shared_cost = gross_after_transport - report['months'][period]['net_profit']
        elif period in months:
            direct_cost_total = abs(sum(value['cost'] for value in product_values.values()))
            shared_cost = months[period] - direct_cost_total - transport_total
        else:
            continue
        product_period_data[period] = {}
        for product_id, value in product_values.items():
            allocated_shared_cost = (
                shared_cost * value['revenue'] / revenue_total if revenue_total else 0)
            product_period_data[period][product_id] = {
                'revenue': value['revenue'],
                'cost': abs(value['cost']),
                'profit': value['revenue'] + value['cost'],
                'transport': value['transport'],
                'shared': allocated_shared_cost,
                'net': (value['revenue'] + value['cost'] - value['transport']
                        - allocated_shared_cost),
                'quantity': value['quantity'],
            }
    for period, entry in periods.items():
        entry['cost'] = abs(entry['cost'])
        entry['profit'] = entry['revenue'] - entry['cost']
        entry['unit_price'] = entry['revenue'] / entry['quantity'] if entry['quantity'] else 0
    annual = {}
    for period, entry in periods.items():
        year = period[:4]
        row = annual.setdefault(year, {'revenue': 0, 'cost': 0, 'quantity': 0})
        for key in ('revenue', 'cost', 'quantity'):
            row[key] += entry[key]
    for entry in annual.values():
        entry['profit'] = entry['revenue'] - entry['cost']
        entry['unit_price'] = entry['revenue'] / entry['quantity'] if entry['quantity'] else 0
    current_year = max((period[:4] for period in periods), default=str(datetime.now().year))
    current_year_cost = sum(value for period, value in months.items() if period.startswith(current_year))
    current_year_data = annual.get(current_year, {'revenue': 0, 'cost': 0, 'quantity': 0, 'profit': 0, 'unit_price': 0})
    total_cost = sum(categories.values())
    total_revenue = sum(customer_revenue.values())
    direct_cost = abs(sum(customer_cost.values()))
    allocated_delivery_transport = sum(delivery_transport_by_customer.values())
    shared_cost = max(total_cost - direct_cost, 0) - allocated_delivery_transport
    revenue_by_customer_code = {}
    for customer, revenue in customer_revenue.items():
        customer_code = customer[:5]
        revenue_by_customer_code[customer_code] = (
            revenue_by_customer_code.get(customer_code, 0) + revenue)
    customer_shared_cost = {}
    customer_transport_cost = {}
    customer_net_profit = {}
    for customer, revenue in customer_revenue.items():
        allocated = shared_cost * revenue / total_revenue if total_revenue else 0
        customer_code = customer[:5]
        code_revenue = revenue_by_customer_code.get(customer_code, 0)
        transport = (delivery_transport_by_customer.get(customer_code, 0) * revenue
                     / code_revenue if code_revenue else 0)
        customer_shared_cost[customer] = allocated
        customer_transport_cost[customer] = transport
        customer_net_profit[customer] = customer_profit[customer] - allocated - transport
    return {'files': files, 'total_cost': total_cost, 'total_revenue': total_revenue,
            'current_year': current_year, 'current_year_cost': current_year_cost,
            'current_year_revenue': current_year_data['revenue'],
            'current_year_profit': current_year_data['revenue'] - current_year_cost,
            'current_year_operating_profit': rozbor.get('operating_profit', 0),
            'current_year_depreciation': rozbor.get('depreciation', 0),
            'current_year_interest': rozbor.get('interest', 0),
            'current_year_subsidies': rozbor.get('subsidies', 0),
            'current_year_net_profit': rozbor.get('net_profit_after_adjustments', 0),
            'monthly_net_profit': rozbor.get('monthly_net_profit', {}),
            'report_months': report.get('months', {}),
            'report_ytd': {
                'operating_profit': report.get('ytd_operating_profit', 0),
                'net_profit': report.get('ytd_net_profit', 0),
                'depreciation': report.get('ytd_depreciation', 0),
                'interest': report.get('ytd_interest', 0),
                'subsidies': report.get('ytd_subsidies', 0),
            },
            'current_categories': categories_by_year.get(current_year, {}),
            'net_profit': total_revenue - total_cost, 'direct_cost': direct_cost,
            'shared_cost': shared_cost, 'customer_shared_cost': customer_shared_cost,
            'customer_transport_cost': customer_transport_cost,
            'customer_net_profit': customer_net_profit, 'categories': categories,
            'months': months, 'periods': periods, 'annual': annual,
            'suppliers': suppliers, 'customers': customers,
            'customer_profit': customer_profit, 'customer_revenue': customer_revenue,
            'customer_cost': customer_cost, 'customer_quantity': customer_quantity,
            'customer_names': customer_names,
            'customer_period_data': customer_period_data,
            'product_names': product_names,
            'product_period_data': product_period_data,
            'deliveries': deliveries,
            'delivery_summary': {
                'has_export': bool(delivery_costs),
                'export_count': len(delivery_costs),
                'matched_count': len(deliveries) - mismatched_delivery_count,
                'matched_transport': allocated_delivery_transport,
                'unmatched_count': unmatched_delivery_count,
                'unmatched_transport': unmatched_delivery_transport,
                'mismatched_count': mismatched_delivery_count,
                'mismatched_transport': mismatched_delivery_transport,
            },
            'dimension_count': len(ranges), 'dimensions_in_export': False,
            'updated': datetime.now().strftime('%d.%m.%Y %H:%M')}

_data_cache = None
_data_signature = None
_cache_file = ROOT / '.dashboard-cache.json'
_load_lock = Lock()

def load_data():
    global _data_cache, _data_signature
    signature = tuple(sorted((path.name, path.stat().st_mtime_ns, path.stat().st_size)
                             for path in ROOT.glob('*.xlsx') if not path.name.startswith('~$')))
    signature = (CACHE_VERSION,) + signature
    if _data_cache is not None and signature == _data_signature:
        return _data_cache
    with _load_lock:
        signature = tuple(sorted((path.name, path.stat().st_mtime_ns, path.stat().st_size)
                                 for path in ROOT.glob('*.xlsx') if not path.name.startswith('~$')))
        signature = (CACHE_VERSION,) + signature
        if _data_cache is not None and signature == _data_signature:
            return _data_cache
        if _cache_file.exists():
            try:
                cached = json.loads(_cache_file.read_text(encoding='utf-8'))
                cached_signature = tuple(item if isinstance(item, int) else tuple(item)
                                        for item in cached.get('_signature', []))
                if cached_signature == signature:
                    _data_cache = cached['data']
                    _data_signature = signature
                    return _data_cache
            except (OSError, ValueError, KeyError):
                pass
        _data_cache = _load_data()
        _data_signature = signature
        try:
            _cache_file.write_text(json.dumps({'_signature': signature, 'data': _data_cache},
                                               ensure_ascii=False), encoding='utf-8')
        except OSError:
            pass
        return _data_cache

class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload):
        body = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == '/api/data':
            self.send_json(load_data())
            return
        target = ROOT / ('index.html' if self.path in ('/', '') else self.path.lstrip('/'))
        if target.exists() and target.suffix == '.html':
            body = target.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_error(404)

    def do_POST(self):
        if self.path != '/api/upload':
            self.send_error(404)
            return
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)
        boundary = self.headers.get('Content-Type', '').split('boundary=')[-1].encode()
        for part in body.split(b'--' + boundary):
            match = re.search(rb'filename="([^"]+\.xlsx)"', part)
            if match:
                name = Path(match.group(1).decode('utf-8', 'ignore')).name
                content = part.split(b'\r\n\r\n', 1)[-1].rsplit(b'\r\n', 1)[0]
                (ROOT / name).write_bytes(content)
        self.send_json({'ok': True})

if __name__ == '__main__':
    print(f'Prehľad hospodárenia: http://localhost:{PORT}')
    ThreadingHTTPServer(('127.0.0.1', PORT), Handler).serve_forever()