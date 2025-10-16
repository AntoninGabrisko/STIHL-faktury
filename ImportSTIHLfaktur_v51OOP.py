import pdfplumber
import pandas as pd
import os
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import requests
import base64
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import subprocess
import time
import tkinter as tk
from tkinter import ttk


class MathHelper:
    """Pomocná třída pro matematické operace"""
    
    @staticmethod
    def zaokrouhli(hodnota, desetinna_mista=2):
        """Matematické zaokrouhlování s variabilním počtem desetinných míst"""
        zaokrouleni_format = '0.' + '0' * desetinna_mista
        return Decimal(hodnota).quantize(Decimal(zaokrouleni_format), rounding=ROUND_HALF_UP)
    
    @staticmethod
    def preved_na_desetinne_cislo(value):
        """Převod řetězce na desetinné číslo (bez mezer, desetinný oddělovač je tečka)"""
        formatted_value = value.replace(" ", "").replace(",", ".")
        return MathHelper.zaokrouhli(formatted_value)


class FileHelper:
    """Pomocná třída pro práci se soubory"""
    
    @staticmethod
    def zapis_do_souboru(cesta_k_souboru, data, system_kodovani):
        """Zapíše data do souboru s ošetřením chyb"""
        try:
            with open(cesta_k_souboru, 'w', encoding=system_kodovani) as file:
                file.write(data)
        except PermissionError:
            print(f"Chyba: Nemám oprávnění zapisovat do souboru {cesta_k_souboru}.")
        except IOError as e:
            print(f"Chyba I/O: {e}")
        except Exception as e:
            print(f"Neočekávaná chyba: {e}")


class MServerStatus:
    """Třída pro kontrolu stavu mServeru"""
    
    @staticmethod
    def zjisti_stav_serveru():
        """Zjistí základní stav POHODA mServeru"""
        try:
            response = requests.get("http://localhost:444/status")
            response.raise_for_status()
            root = ET.fromstring(response.content)
            return {'processing': int(root.find('processing').text)}
        except requests.exceptions.RequestException as e:
            raise Exception(f"Chyba při komunikaci se serverem: {str(e)}")
        except (AttributeError, ValueError) as e:
            raise Exception(f"Neočekávaný formát odpovědi: {str(e)}")


class MServerInitializer:
    """Třída pro inicializaci mServeru"""
    
    @staticmethod
    def inicializace_mServeru():
        """Zobrazí dialogové okno a inicializuje mServer"""
        root = tk.Tk()
        root.title("Inicializace")
        
        window_width = 300
        window_height = 100
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)
        root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
        
        label = ttk.Label(root, text="POHODA mServer není spuštěn !\n...spouštím mServer vzdáleným voláním...\nČekejte prosím !", anchor="center")
        label.pack(expand=True, padx=20, pady=20)

        try:
            result = subprocess.Popen(r'\\POHODA\Pohoda\Pohoda.exe /http start "Firma"', shell=True)
            root.update()
            time.sleep(25)
            root.destroy()
            print("Aplikace POHODA mServer byla spuštěna jako proces na pozadí...")
        except FileNotFoundError:
            print("Cesta k aplikaci nebyla nalezena ! Zjisti, zda je síťová cesta k Pohodě na serveru správná a že k ní máš přístup.")
        except Exception as e:
            print(f"Chyba při spuštění aplikace: {e}")


class PDFParser:
    """Třída pro zpracování PDF faktur"""
    
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.text = self._preved_pdf_na_text()
    
    def _preved_pdf_na_text(self):
        """Převede PDF fakturu na textový soubor"""
        with pdfplumber.open(self.pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() if page.extract_text() else ""
        
        FileHelper.zapis_do_souboru(self.pdf_path + '.txt', text, 'utf-8')
        return text
    
    def extrahuj_produktove_informace(self):
        """Extrahuje detaily produktů z faktury"""
        regex_product_details = r"(\d{3,5})\s+(\d{1,3}[\dA-Za-z-]{13})\s+([\d\s]{1,8},\d{2})\s+(\d+\s+\w+)\s+(-?\d{2},\d)?\s*([\d\s]{1,8},\d{2})\s+([\d\s]{1,8},\d{2})\s+(\w{2})"
        regex_shipping_details = r"(Dopravné)\s+([\d\s]{1,8},\d{2})\s+(\w{2})"
        
        lines = self.text.split("\n")
        products = []
        
        for i in range(len(lines)):
            match_product = re.search(regex_product_details, lines[i])
            match_shipping = re.search(regex_shipping_details, lines[i])
            
            if match_product:
                match_product = match_product.groups()
                product_name = lines[i + 1]
                
                quantity_product_number = match_product[1]
                quantity = int(quantity_product_number[:-13])
                product_number = quantity_product_number[-13:].replace("-", "")
                
                discount = match_product[4] if match_product[4] else "0,0"
                list_price = MathHelper.preved_na_desetinne_cislo(match_product[2])
                total_amount = MathHelper.preved_na_desetinne_cislo(match_product[6])
                unit_price_after_discount = total_amount / quantity
                
                product = {
                    "Číslo řádku": match_product[0],
                    "Množství": quantity,
                    "Číslo produktu": product_number,
                    "Ceníková cena": str(list_price),
                    "Jednotka": match_product[3],
                    "Sleva": discount,
                    "Jednotková cena po slevě": str(unit_price_after_discount),
                    "Částka celkem": str(total_amount),
                    "Kód DPH": match_product[7],
                    "Název produktu": product_name
                }
                products.append(product)
            
            elif match_shipping:
                match_shipping = match_shipping.groups()
                shipping_price = MathHelper.preved_na_desetinne_cislo(match_shipping[1])
                
                shipping = {
                    "Číslo řádku": None,
                    "Množství": '1.00',
                    "Číslo produktu": None,
                    "Ceníková cena": str(shipping_price),
                    "Jednotka": None,
                    "Sleva": '0.00',
                    "Jednotková cena po slevě": None,
                    "Částka celkem": str(shipping_price),
                    "Kód DPH": 'A1',
                    "Název produktu": 'Vedlejší náklady'
                }
                products.append(shipping)
        
        return pd.DataFrame(products)
    
    def rozdel_fakturu_podle_objednavek(self):
        """Rozdělí fakturu podle jednotlivých objednávek"""
        objednavka = {}
        cislo_objednavky = None
        
        lines = self.text.split('\n')
        for line in lines:
            if "Číslo zák. obj.:" in line:
                cislo_objednavky = line.split("Číslo zák. obj.:")[1].strip()
                if cislo_objednavky not in objednavka:
                    objednavka[cislo_objednavky] = []
            if cislo_objednavky:
                objednavka[cislo_objednavky].append(line)
        
        return objednavka


class OrderProcessor:
    """Třída pro zpracování objednávek"""
    
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.username = "Monzar"
        self.password = "Tonda"
        self.url = "http://localhost:444/xml"
        self.objednavka_kompletne_zpracovana = None
    
    def _get_headers(self):
        """Vytvoří autentizační hlavičku"""
        auth_string = f"{self.username}:{self.password}"
        auth_bytes = auth_string.encode('ascii')
        base64_bytes = base64.b64encode(auth_bytes)
        base64_string = base64_bytes.decode('ascii')
        return {"STW-Authorization": f"Basic {base64_string}"}
    
    def zpracuj_objednavku(self, cislo_objednavky, order_products_df):
        """Zpracuje objednávku a rozšíří DataFrame o informace z POHODA"""
        self.objednavka_kompletne_zpracovana = True
        
        xml_data = f"""<dat:dataPack id="ExportujObj" ico="75973502" application="ImportSTIHLfaktur_v41" version="2.0" note="Export vydané objednávky" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd" xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd" xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd" xmlns:lst="http://www.stormware.cz/schema/version_2/list.xsd">
    <dat:dataPackItem id="ExportujObj" version="2.0" note="Export ID položek z Vydané objednávky">
    <lst:listOrderRequest version="2.0" orderType="issuedOrder" orderVersion="2.0">
    <lst:requestOrder><ftr:filter><ftr:selectedNumbers><ftr:number><typ:numberRequested>{cislo_objednavky}</typ:numberRequested></ftr:number></ftr:selectedNumbers></ftr:filter></lst:requestOrder></lst:listOrderRequest>
    </dat:dataPackItem></dat:dataPack>"""
        
        headers = self._get_headers()
        headers.update({
            'User-Agent': 'MujPythonKlient/1.0',
            'Content-Type': 'text/xml',
            'Accept-Encoding': 'gzip, deflate'
        })
        
        try:
            response = requests.post(self.url, headers=headers, data=xml_data.encode('utf-8'))
            if response.ok:
                response_file_path = os.path.join(self.output_dir, f"{cislo_objednavky}.xml")
                print(f"\nOdpověď dotazu mServeru na Objednávku {cislo_objednavky} ukládám do souboru \"{response_file_path}\"")
                
                response_text = response.content.decode('windows-1250', errors='replace')
                FileHelper.zapis_do_souboru(response_file_path, response_text, 'windows-1250')
                
                state_pattern = r'<rsp:responsePack[^>]+state="([^"]*)"'
                matches = re.findall(state_pattern, response_text)
                
                if matches:
                    status = matches[0]
                    if status == 'ok':
                        print(f"\"state\" odpovědi na objednávku číslo {cislo_objednavky} je \"{status}\"")
                    else:
                        print(f"Objednávka {cislo_objednavky} neexistuje, neboť \"state\" odpovědi je \"{status}\"")
                        return False
                
                xml_root = ET.fromstring(response.content.decode('windows-1250', errors='replace'))
                
                order_products_df['ObjednacíID'] = None
                order_products_df['Objednané množství'] = None
                order_products_df['Dodané množství'] = None
                
                get_is_executed = xml_root.find('.//ord:isExecuted', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text
                get_is_delivered = xml_root.find('.//ord:isDelivered', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text
                print(f"\nStav objednávky {cislo_objednavky} je:\n     isExecuted:{get_is_executed}\n     isDelivered:{get_is_delivered}")
                
                if get_is_delivered == 'true':
                    print(f"\nObjednávka {cislo_objednavky} je označena jako kompletně přenesená !!!")
                    print("Zjisti, proč faktura odkazuje právě na tuto (kompletně přenesenou) objednávku.\n")
                    input("Stiskněte Enter pro ukončení...")
                    sys.exit("Import faktury nemohl být z důvodu chyby dokončen.")
                
                if get_is_executed == 'true':
                    print(f"\nObjednávka {cislo_objednavky} je označena jako vyřízená !!!")
                    print("Zruš příznak \"Vyřízeno\" u této objednávky !\n")
                    input("Stiskněte Enter pro ukončení...")
                    sys.exit("Import faktury nemohl být z důvodu chyby dokončen.")
                
                text_objednavky = xml_root.find('.//ord:orderHeader/ord:text', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'})
                if text_objednavky is not None:
                    text_objednavky = text_objednavky.text
                    text_objednavky_bez_data = re.sub(r'^\d{4}-\d{2}-\d{2}\s+', '', text_objednavky)
                    text_objednavky = text_objednavky_bez_data
                    print(f"Text objednávky {cislo_objednavky} je \"{text_objednavky}\"")
                else:
                    text_objednavky = ""
                    print(f"Text objednávky {cislo_objednavky} nebyl nalezen")
                
                for get_order_item in xml_root.findall('.//ord:orderItem', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}):
                    get_item_id = get_order_item.find('.//ord:id', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text
                    get_ids = get_order_item.find('.//ord:code', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text
                    order_quantity = MathHelper.zaokrouhli(get_order_item.find('.//ord:quantity', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text)
                    delivered_quantity = MathHelper.zaokrouhli(get_order_item.find('.//ord:delivered', namespaces={'ord': 'http://www.stormware.cz/schema/version_2/order.xsd'}).text)
                    
                    print(f"\nPoložka objednávky s ID {get_item_id} má kód {get_ids} s objednaným množstvím {order_quantity}, dodaná v počtu {delivered_quantity}.")
                    
                    mask = order_products_df['Číslo produktu'] == get_ids
                    order_products_df.loc[mask, 'ObjednacíID'] = get_item_id
                    order_products_df.loc[mask, 'Objednané množství'] = order_quantity
                    order_products_df.loc[mask, 'Dodané množství'] = delivered_quantity
                    
                    if not order_products_df['Číslo produktu'].isin([get_ids]).any():
                        print(f"Objednací číslo '{get_ids}' nebylo nalezeno, proto je DataFrame prázdný.")
                        if not order_quantity == delivered_quantity:
                            self.objednavka_kompletne_zpracovana = False
                            print(f"\nPoložka {get_ids} objednávky ještě není kompletně vykryta !\n")
                    else:
                        print(f"Objednací číslo '{get_ids}' bylo nalezeno.")
                        radek = order_products_df[order_products_df['Číslo produktu'] == get_ids]
                        mnozstvi = radek['Množství'].iloc[0]
                        print(f"\nMnožství ve faktuře je: {mnozstvi}\n")
                        if int(order_quantity) - int(delivered_quantity) == mnozstvi:
                            print(f"\nPoložka {get_ids} objednávky je zcela vykryta !\n")
                        else:
                            self.objednavka_kompletne_zpracovana = False
                        
                        print(order_products_df[order_products_df['Číslo produktu'] == get_ids])
                
                print(f"\nVýsledný DataFrame ve shodě s objednávkou je:\n{order_products_df}\n")
                print(f"Objednávka {cislo_objednavky} {'je' if self.objednavka_kompletne_zpracovana else 'není'} kompletně vykrytá.\n")
                return order_products_df, text_objednavky
            else:
                print(f"\nChyba v HTTP požadavku na objednávku \"{cislo_objednavky}\"\n generovala HTTP kód \"{response.status_code}\"")
                return False
        except Exception as e:
            print(f"\nChyba v HTTP požadavku na objednávku číslo \"{cislo_objednavky}\" skončila chybovým kódem: \"{e}\"")
            return False


class StockXMLValidator:
    """Validátor pro XML odpovědi při načítání seznamu skladových zásob."""
    
    @staticmethod
    def validate(response_text):
        """Kontroluje XML odpověď pro načtení skladových zásob."""
        try:
            from xml.dom import minidom
            doc = minidom.parseString(response_text)
            
            namespaces = {
                'rsp': 'http://www.stormware.cz/schema/version_2/response.xsd',
                'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd',
                'lStk': 'http://www.stormware.cz/schema/version_2/list_stock.xsd',
                'stk': 'http://www.stormware.cz/schema/version_2/stock.xsd'
            }
            
            response_items = doc.getElementsByTagNameNS(namespaces['rsp'], 'responsePackItem')
            if not response_items:
                return False, "Chybí element responsePackItem v odpovědi"
            
            has_critical_error = False
            error_messages = []
            warning_messages = []
            stock_count = 0
            items_processed = 0
            items_with_errors = 0
            
            for response_item in response_items:
                state = response_item.getAttribute('state')
                item_id = response_item.getAttribute('id')
                
                if state == 'error':
                    has_critical_error = True
                    items_with_errors += 1
                    error_messages.append(f"Chyba v položce {item_id}")
                elif state == 'ok':
                    items_processed += 1
                    
                    list_stocks = response_item.getElementsByTagNameNS(namespaces['lStk'], 'listStock')
                    for list_stock in list_stocks:
                        list_stock_state = list_stock.getAttribute('state')
                        if list_stock_state == 'ok':
                            stocks = list_stock.getElementsByTagNameNS(namespaces['lStk'], 'stock')
                            stock_count += len(stocks)
            
            result_parts = []
            
            if stock_count > 0:
                result_parts.append(f"✅ Načteno {stock_count} skladových zásob")
                if items_processed > 0:
                    result_parts.append(f"📦 Úspěšně zpracováno {items_processed} požadavků")
            elif items_processed > 0:
                result_parts.append(f"⚠️ Zpracováno {items_processed} požadavků, ale nenalezeny žádné zásoby")
            
            if items_with_errors > 0:
                if result_parts:
                    result_parts.append("")
                result_parts.append(f"❌ Počet chybných požadavků: {items_with_errors}")
            
            if error_messages:
                for msg in error_messages:
                    result_parts.append(f"  - {msg}")
            
            if warning_messages:
                if result_parts:
                    result_parts.append("")
                result_parts.append("⚠️ Varování:")
                for msg in warning_messages:
                    result_parts.append(f"  - {msg}")
            
            if stock_count == 0 and items_processed == 0 and not error_messages:
                return True, "XML odpověď je v pořádku, ale nebyly nalezeny žádné požadavky k zpracování"
            
            return not has_critical_error, "\n".join(result_parts)
            
        except Exception as e:
            return False, f"Chyba při kontrole XML odpovědi: {str(e)}"


class StockXMLBuilder:
    """Vytváří XML pro načtení skladových zásob."""
    
    def __init__(self):
        ET.register_namespace('dat', 'http://www.stormware.cz/schema/version_2/data.xsd')
        ET.register_namespace('stk', 'http://www.stormware.cz/schema/version_2/stock.xsd')
        ET.register_namespace('ftr', 'http://www.stormware.cz/schema/version_2/filter.xsd')
        ET.register_namespace('lStk', 'http://www.stormware.cz/schema/version_2/list_stock.xsd')
        ET.register_namespace('typ', 'http://www.stormware.cz/schema/version_2/type.xsd')
        
        self.ns = {
            'dat': 'http://www.stormware.cz/schema/version_2/data.xsd',
            'stk': 'http://www.stormware.cz/schema/version_2/stock.xsd',
            'ftr': 'http://www.stormware.cz/schema/version_2/filter.xsd',
            'lStk': 'http://www.stormware.cz/schema/version_2/list_stock.xsd',
            'typ': 'http://www.stormware.cz/schema/version_2/type.xsd'
        }
    
    def build(self, df):
        """Vytvoří XML pro načtení skladových zásob."""
        if 'Číslo produktu' not in df.columns:
            print(f"❌ CHYBA: DataFrame neobsahuje sloupec 'Číslo produktu'!")
            return ""
        
        print(f"Vytvářím XML pro načtení {len(df)} položek ze skladu...")
        
        root = ET.Element(f'{{{self.ns["dat"]}}}dataPack')
        root.set('version', '2.0')
        root.set('id', 'ExportujZasoby')
        root.set('ico', '75973502')
        root.set('application', 'Export vybranych zasob')
        root.set('note', 'Exportuj vybrane zasoby podle Kodu')
        
        for index, row in df.iterrows():
            cislo_materialu = str(row['Číslo produktu']).strip()
            
            if not cislo_materialu or cislo_materialu.lower() in ['nan', 'none', '']:
                continue
            
            pack_item = ET.SubElement(root, f'{{{self.ns["dat"]}}}dataPackItem')
            pack_item.set('version', '2.0')
            pack_item.set('id', f'ZAS{index + 1:02d}')
            
            list_stock_request = ET.SubElement(pack_item, f'{{{self.ns["lStk"]}}}listStockRequest')
            list_stock_request.set('version', '2.0')
            list_stock_request.set('stockVersion', '2.0')
            
            request_stock = ET.SubElement(list_stock_request, f'{{{self.ns["lStk"]}}}requestStock')
            filter_element = ET.SubElement(request_stock, f'{{{self.ns["ftr"]}}}filter')
            code_filter = ET.SubElement(filter_element, f'{{{self.ns["ftr"]}}}code')
            code_filter.text = cislo_materialu
        
        return ET.tostring(root, encoding='unicode', method='xml')


class StockDataProcessor:
    """Třída pro zpracování dat skladových zásob."""
    
    @staticmethod
    def extract_from_xml(xml_content):
        """Extrahuje data skladových zásob z XML odpovědi."""
        try:
            from io import StringIO
            
            try:
                root = ET.fromstring(xml_content)
            except ET.ParseError as e:
                print(f"Chyba při parsování XML: {e}")
                root = ET.parse(StringIO(xml_content)).getroot()
        
            data = []
            namespaces = {
                'rsp': 'http://www.stormware.cz/schema/version_2/response.xsd',
                'lStk': 'http://www.stormware.cz/schema/version_2/list_stock.xsd',
                'stk': 'http://www.stormware.cz/schema/version_2/stock.xsd',
                'typ': 'http://www.stormware.cz/schema/version_2/type.xsd'
            }
            
            for response_pack_item in root.findall('.//rsp:responsePackItem', namespaces):
                item_id = response_pack_item.get('id', 'Neznámé ID')
                item_state = response_pack_item.get('state', 'Neznámý stav')
                
                for stock in response_pack_item.findall('.//lStk:stock', namespaces):
                    stock_header = stock.find('stk:stockHeader', namespaces)
                    
                    if stock_header is not None:
                        stock_data = {
                            'responsePackItem_id': item_id,
                            'responsePackItem_state': item_state,
                            'id': StockDataProcessor._get_element_text(stock_header, 'stk:id', namespaces),
                            'Číslo produktu': StockDataProcessor._get_element_text(stock_header, 'stk:code', namespaces),
                            'EAN': StockDataProcessor._get_element_text(stock_header, 'stk:EAN', namespaces),
                            'Název': StockDataProcessor._get_element_text(stock_header, 'stk:name', namespaces),
                            'Jednotka': StockDataProcessor._get_element_text(stock_header, 'stk:unit', namespaces),
                        }
                        
                        # Extrakce DPH sazeb
                        purch_vat = stock_header.find('stk:purchasingRateVAT', namespaces)
                        if purch_vat is not None:
                            stock_data['Nákup DPH text'] = purch_vat.text if purch_vat.text else ''
                            stock_data['Nákup DPH value'] = purch_vat.get('value', '')
                        else:
                            stock_data['Nákup DPH text'] = ''
                            stock_data['Nákup DPH value'] = ''
                        
                        sell_vat = stock_header.find('stk:sellingRateVAT', namespaces)
                        if sell_vat is not None:
                            stock_data['Prodej DPH text'] = sell_vat.text if sell_vat.text else ''
                            stock_data['Prodej DPH value'] = sell_vat.get('value', '')
                        else:
                            stock_data['Prodej DPH text'] = ''
                            stock_data['Prodej DPH value'] = ''
                        
                        # Extrakce členění (storage)
                        storage = stock_header.find('stk:storage', namespaces)
                        if storage is not None:
                            storage_id = storage.find('typ:id', namespaces)
                            storage_ids = storage.find('typ:ids', namespaces)
                            stock_data['Členění ID'] = storage_id.text if storage_id is not None else ''
                            stock_data['Členění'] = storage_ids.text if storage_ids is not None else ''
                        else:
                            stock_data['Členění ID'] = ''
                            stock_data['Členění'] = ''
                        
                        # Extrakce cenové skupiny (typePrice)
                        type_price = stock_header.find('stk:typePrice', namespaces)
                        if type_price is not None:
                            price_id = type_price.find('typ:id', namespaces)
                            price_ids = type_price.find('typ:ids', namespaces)
                            stock_data['Cenová skupina ID'] = price_id.text if price_id is not None else ''
                            stock_data['Cenová skupina'] = price_ids.text if price_ids is not None else ''
                        else:
                            stock_data['Cenová skupina ID'] = ''
                            stock_data['Cenová skupina'] = ''
                        
                        # Extrakce nákupní ceny
                        purch_price = stock_header.find('stk:purchasingPrice', namespaces)
                        stock_data['Nákupní cena'] = float(purch_price.text) if purch_price is not None and purch_price.text else 0.0
                        
                        # Extrakce prodejní ceny včetně atributu payVAT
                        sell_price = stock_header.find('stk:sellingPrice', namespaces)
                        if sell_price is not None:
                            stock_data['Prodejní cena'] = float(sell_price.text) if sell_price.text else 0.0
                            stock_data['Prodejní cena payVAT'] = sell_price.get('payVAT', '')
                        else:
                            stock_data['Prodejní cena'] = 0.0
                            stock_data['Prodejní cena payVAT'] = ''
                        
                        # Extrakce fixace ceny
                        fixation = stock_header.find('stk:fixation', namespaces)
                        stock_data['Fixace'] = fixation.text if fixation is not None and fixation.text else ''
                        
                        # Extrakce ID dodavatele
                        supplier = stock_header.find('stk:supplier', namespaces)
                        if supplier is not None:
                            supplier_id = supplier.find('typ:id', namespaces)
                            stock_data['Dodavatel ID'] = supplier_id.text if supplier_id is not None else ''
                        else:
                            stock_data['Dodavatel ID'] = ''
                        
                        data.append(stock_data)
            
            return pd.DataFrame(data)
        
        except Exception as e:
            print(f"Chyba při extrakci dat: {str(e)}")
            return pd.DataFrame()
    
    @staticmethod
    def _get_element_text(parent, xpath, namespaces):
        """Pomocná funkce pro získání textu z XML elementu."""
        element = parent.find(xpath, namespaces)
        return element.text if element is not None else None


class StockManager:
    """Třída pro správu zásob - načítání informací o zásobách z POHODA"""
    
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.username = "Monzar"
        self.password = "Tonda"
        self.url = "http://localhost:444/xml"
        self.xml_builder = StockXMLBuilder()
        self.data_processor = StockDataProcessor()
    
    def _get_headers(self):
        """Vytvoří autentizační hlavičku"""
        credentials = f"{self.username}:{self.password}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        return {
            'User-Agent': 'MujPythonKlient/1.0',
            'STW-Authorization': f"Basic {encoded_credentials}",
            'Content-Type': 'text/xml',
            'Accept-Encoding': 'gzip, deflate'
        }
    
    def nacti_zasoby_podle_objednacich_id(self, df_shody):
        """
        Na základě ObjednacíID z DataFrame načte informace o zásobách z POHODA
        
        Args:
            df_shody: DataFrame s informacemi o shodách, musí obsahovat sloupec 'Číslo produktu'
        
        Returns:
            DataFrame s informacemi o zásobách
        """
        if 'Číslo produktu' not in df_shody.columns:
            print("Chyba: DataFrame neobsahuje sloupec 'Číslo produktu'")
            return pd.DataFrame()
        
        print(f"\n{'='*60}")
        print("NAČÍTÁM SKLADOVÉ ZÁSOBY Z mSERVERU")
        print(f"{'='*60}")
        
        # Vytvoření XML požadavku
        xml_request = self.xml_builder.build(df_shody)
        
        if not xml_request:
            print("Chyba při vytváření XML požadavku")
            return pd.DataFrame()
        
        # Uložení vstupního XML
        input_file = f"{self.output_dir}Nacitani_zasob-vstup.xml"
        FileHelper.zapis_do_souboru(input_file, xml_request, 'utf-8')
        print(f"XML požadavek uložen do: {input_file}")
        
        try:
            # Odeslání požadavku
            response = requests.post(self.url, headers=self._get_headers(), data=xml_request.encode('utf-8'))
            
            if response.ok:
                print(f"XML požadavek zpracován, status: {response.status_code}")
                
                response_text = response.content.decode('windows-1250', errors='replace')
                
                # Uložení odpovědi
                output_file = f"{self.output_dir}Nacitani_zasob-vystup.xml"
                FileHelper.zapis_do_souboru(output_file, response_text, 'windows-1250')
                print(f"Ukládám odpověď do: {output_file}")
                
                # Validace odpovědi
                je_bez_chyb, zprava = StockXMLValidator.validate(response_text)
                print(zprava)
                
                if je_bez_chyb:
                    # Extrakce dat
                    df_zasoby = self.data_processor.extract_from_xml(response_text)
                    
                    if not df_zasoby.empty:
                        stock_file = f"{self.output_dir}Skladove_zasoby.xlsx"
                        print(f"Skladové zásoby ukládám do: {stock_file}")
                        df_zasoby.to_excel(stock_file, index=False)
                    
                    return df_zasoby
                else:
                    print("❌ Chyba při validaci XML odpovědi")
                    return pd.DataFrame()
            else:
                print(f"Chyba při komunikaci s mServerem: {response.status_code}")
                return pd.DataFrame()
        
        except Exception as e:
            print(f"Chyba při načítání zásob: {str(e)}")
            return pd.DataFrame()


class XMLGenerator:
    """Třída pro generování XML elementů"""
    
    @staticmethod
    def vytvor_xml_elementy_polozek(prijemkaDetail, matched_df):
        """Vytvoří XML elementy "prijemkaItem" pro každou položku příjemky"""
        for _, row in matched_df.iterrows():
            prijemkaItem = ET.SubElement(prijemkaDetail, 'pri:prijemkaItem')
            
            link = ET.SubElement(prijemkaItem, 'pri:link')
            sourceAgenda = ET.SubElement(link, 'typ:sourceAgenda')
            sourceAgenda.text = 'issuedOrder'
            sourceItemId = ET.SubElement(link, 'typ:sourceItemId')
            sourceItemId.text = str(row['ObjednacíID'])
            
            settingsSourceDocumentItem = ET.SubElement(link, 'typ:settingsSourceDocumentItem')
            linkIssuedOrderToReceiptVoucher = ET.SubElement(settingsSourceDocumentItem, 'typ:linkIssuedOrderToReceiptVoucher')
            linkIssuedOrderToReceiptVoucher.text = '2'
            
            ET.SubElement(prijemkaItem, 'pri:text').text = str(row['Název produktu'])
            ET.SubElement(prijemkaItem, 'pri:quantity').text = str(MathHelper.preved_na_desetinne_cislo(str(row['Množství'])))
            ET.SubElement(prijemkaItem, 'pri:unit').text = 'ks'
            ET.SubElement(prijemkaItem, 'pri:payVAT').text = 'false'
            
            rateVAT = ET.SubElement(prijemkaItem, 'pri:rateVAT')
            price_vat = '0.00'
            price_sum = '0.00'
            
            if str(row['Kód DPH']) == 'A1':
                rateVAT.text = 'high'
                price = MathHelper.zaokrouhli(row['Částka celkem'])
                price_vat = price * Decimal('0.21')
                price_sum = str(MathHelper.zaokrouhli(price + price_vat))
                price_vat = str(MathHelper.zaokrouhli(price_vat))
            
            sleva = MathHelper.preved_na_desetinne_cislo(str(row['Sleva']))
            sleva = abs(sleva)
            ET.SubElement(prijemkaItem, 'pri:discountPercentage').text = str(sleva)
            
            homeCurrency = ET.SubElement(prijemkaItem, 'pri:homeCurrency')
            
            mnozstvi = row['Množství']
            jednotkova_cena_po_sleve = MathHelper.preved_na_desetinne_cislo(row['Jednotková cena po slevě'])
            po_sleve_za_vsechny = MathHelper.zaokrouhli(jednotkova_cena_po_sleve * mnozstvi)
            castka_celkem = MathHelper.zaokrouhli(row['Částka celkem'])
            koeficient_slevy = 100 - sleva
            
            print('Proměnné:', castka_celkem, ',', koeficient_slevy, ',', MathHelper.zaokrouhli(po_sleve_za_vsechny,3), ',', mnozstvi, ',', jednotkova_cena_po_sleve, '\n')
            
            cenikova_cena_za_kus = MathHelper.zaokrouhli(castka_celkem * 100 / koeficient_slevy / mnozstvi, 3)
            ET.SubElement(homeCurrency, 'typ:unitPrice').text = str(cenikova_cena_za_kus)
            ET.SubElement(homeCurrency, 'typ:price').text = str(MathHelper.zaokrouhli(row['Částka celkem']))
            
            celkova_dan = castka_celkem / 100 * 21
            ET.SubElement(homeCurrency, 'typ:priceVAT').text = str(MathHelper.zaokrouhli(celkova_dan))
            ET.SubElement(prijemkaItem, 'pri:code').text = str(row['Číslo produktu'])
        
        return prijemkaDetail


class ReceiptManager:
    """Třída pro správu příjemek"""
    
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.username = "Monzar"
        self.password = "Tonda"
        self.url = "http://localhost:444/xml"
        self.zpracovane_objednavky = None
    
    def _get_headers(self):
        """Vytvoří autentizační hlavičku"""
        credentials = f"{self.username}:{self.password}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        return {
            'User-Agent': 'MujPythonKlient/1.0',
            'STW-Authorization': f"Basic {encoded_credentials}",
            'Content-Type': 'text/xml',
            'Accept-Encoding': 'gzip, deflate'
        }
    
    def kontrola_xml_odpovedi(self, xml_odpoved):
        """Kontroluje XML odpověď po odeslání příjemky"""
        try:
            root = ET.fromstring(xml_odpoved)
            
            if root.get('state') != 'ok':
                return False, "Hlavní stav odpovědi není 'ok'"
            
            response_pack_item = root.find('.//rsp:responsePackItem', namespaces={'rsp': 'http://www.stormware.cz/schema/version_2/response.xsd'})
            if response_pack_item is None or response_pack_item.get('state') != 'ok':
                return False, "Stav responsePackItem není 'ok' nebo element chybí"
            
            prijemka_response = root.find('.//pri:prijemkaResponse', namespaces={'pri': 'http://www.stormware.cz/schema/version_2/prijemka.xsd'})
            if prijemka_response is None or prijemka_response.get('state') != 'ok':
                return False, "Stav prijemkaResponse není 'ok' nebo element chybí"
            
            produced_details = root.find('.//rdc:producedDetails', namespaces={'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd'})
            if produced_details is None:
                return False, "Element producedDetails chybí"
            
            import_details = root.findall('.//rdc:importDetails/rdc:detail', namespaces={'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd'})
            warnings = []
            for detail in import_details:
                state = detail.find('rdc:state', namespaces={'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd'})
                if state is not None and state.text == 'warning':
                    errno = detail.find('rdc:errno', namespaces={'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd'})
                    note = detail.find('rdc:note', namespaces={'rdc': 'http://www.stormware.cz/schema/version_2/documentresponse.xsd'})
                    warnings.append(f"Varování {errno.text if errno is not None else 'N/A'}: {note.text if note is not None else 'Bez popisu'}")
            
            if warnings:
                return True, f"XML odpověď obsahuje varování:\n" + "\n".join(warnings)
            
            return True, "XML odpověď je bez chyb."
        
        except ET.ParseError:
            return False, "Chyba při parsování XML"
        except Exception as e:
            return False, f"Neočekávaná chyba: {str(e)}"
    
    def odesli_pozadavek_serveru(self, xml_str, cislo_faktury):
        """Odešle data mServeru a vytiskne informaci o výsledku importu"""
        try:
            response = requests.post(self.url, headers=self._get_headers(), data=xml_str.encode('utf-8'))
            if response.ok:
                print("Import Příjemky byl zpracován a odpověď Serveru je:", response.status_code)
                nazev_souboru = f"Odpoved-po-odeslani-prijemky-k-fakture-{os.path.basename(cislo_faktury)}.xml"
                response_text = response.content.decode('windows-1250', errors='replace')
                FileHelper.zapis_do_souboru(nazev_souboru, response_text, 'windows-1250')
                
                je_bez_chyb, zprava = self.kontrola_xml_odpovedi(response.text)
                print("Kontrola XML odpovědi:", zprava)
                
                if je_bez_chyb:
                    return True, response.status_code
                else:
                    print("CHYBA při kontrole XML odpovědi.")
                    return False, response.status_code
            else:
                print("\nPři importu Příjemky nastala chyba:", response.status_code, response.text)
                return False, response.status_code
        except Exception as e:
            print(f"\nChyba v HTTP požadavku na import Příjemky s chybovým kódem: \"{e}\"")
            return False, None
    
    def zpracuj_vsechny_objednavky(self, cislo_faktury, objednavka, order_processor):
        """Zpracuje všechny objednávky a vytvoří příjemku"""
        is_shipping = None
        shipping_df = None
        self.zpracovane_objednavky = None
        
        dataPack = ET.Element('dat:dataPack', {
            'id': "ImportPrijemky",
            'ico': "75973502",
            'application': "ImportSTIHLfaktur_v45",
            'version': "2.0",
            'note': "XML Import Příjemky přenosem z Objednávky",
            'xmlns:dat': 'http://www.stormware.cz/schema/version_2/data.xsd',
            'xmlns:lst': 'http://www.stormware.cz/schema/version_2/list.xsd',
            'xmlns:pri': 'http://www.stormware.cz/schema/version_2/prijemka.xsd',
            'xmlns:typ': 'http://www.stormware.cz/schema/version_2/type.xsd'
        })
        
        dataPackItem = ET.SubElement(dataPack, 'dat:dataPackItem', {'id': "XMLprijemkaZobjednavky", 'version': "2.0"})
        prijemka = ET.SubElement(dataPackItem, 'pri:prijemka', {'version': "2.0", 'xmlns:pri': "http://www.stormware.cz/schema/version_2/prijemka.xsd"})
        prijemkaHeader = ET.SubElement(prijemka, 'pri:prijemkaHeader')
        
        date = ET.SubElement(prijemkaHeader, 'pri:date')
        date.text = datetime.now().strftime('%Y-%m-%d')
        
        partnerIdentity = ET.SubElement(prijemkaHeader, 'pri:partnerIdentity')
        id_elem = ET.SubElement(partnerIdentity, 'typ:id')
        id_elem.text = '13'
        
        prijemkaDetail = ET.SubElement(prijemka, 'pri:prijemkaDetail')
        
        pdf_parser = PDFParser(cislo_faktury)
        
        for cislo_objednavky, order_lines in objednavka.items():
            df = pdf_parser.extrahuj_produktove_informace()
            df_filtered = pd.DataFrame()
            
            for line in order_lines:
                for _, row in df.iterrows():
                    if str(row['Číslo produktu']) in line or str(row['Název produktu']) in line:
                        df_filtered = pd.concat([df_filtered, pd.DataFrame([row])], ignore_index=True)
                        break
            
            if df_filtered.empty:
                df_filtered = df
            
            output_filename = f"{self.output_dir}{cislo_objednavky}.xlsx"
            print("\nFakturované položky z objednávky číslo \"" + cislo_objednavky + "\" ukládám do adresáře \"" + self.output_dir +
                  "\"\nTabulka fakturovaných položek objednávky uložena do \"" + output_filename + "\"")
            
            try:
                df_filtered.to_excel(output_filename, index=False)
            except PermissionError:
                print(f"Chyba při zápisu do souboru: '{output_filename}'. Soubor může být otevřený v jiném programu nebo nemáte potřebná práva.")
            except Exception as e:
                print(f"Obecná chyba při zápisu do souboru: '{output_filename}'. Chyba: {e}")
            
            print("\nChystám se mServeru dotázat na objednávku číslo \"" + cislo_objednavky + "\"")
            
            df_prijemka, text_objednavky = order_processor.zpracuj_objednavku(cislo_objednavky, df_filtered)
            
            if not isinstance(df_prijemka, pd.DataFrame):
                continue
            
            print(f"\nTabulka shod mezi fakturou a objednávkou je:\n{df_prijemka}")
            
            if self.zpracovane_objednavky is None:
                self.zpracovane_objednavky = f"STIHL - {cislo_objednavky}"
            else:
                self.zpracovane_objednavky += f", {cislo_objednavky}"
            
            output_filename = f"{self.output_dir}{cislo_objednavky}-shody.xlsx"
            try:
                df_prijemka.to_excel(output_filename, index=False)
            except PermissionError:
                print(f"Chyba při zápisu do souboru: '{output_filename}'. Soubor může být otevřený v jiném programu nebo nemáte potřebná práva.")
            except Exception as e:
                print(f"Obecná chyba při zápisu do souboru: '{output_filename}'. Chyba: {e}")
            
            is_shipping = df_prijemka['Číslo produktu'].isna()
            is_product = df_prijemka['Číslo produktu'].notna()
            
            products_df = df_prijemka[is_product]
            
            if shipping_df is None:
                if is_shipping.any():
                    shipping_df = df_prijemka[is_shipping]
            
            prijemkaDetail = XMLGenerator.vytvor_xml_elementy_polozek(prijemkaDetail, products_df)
        
        text = ET.SubElement(prijemkaHeader, 'pri:text')
        text.text = self.zpracovane_objednavky
        
        if shipping_df is not None:
            print(f"\nDataFrame 'shipping_df' je naplněn daty u objednávky číslo {cislo_objednavky}.")
        else:
            print("\nDataFrame 'shipping_df' je prázdný, takže na faktuře není účtována Doprava.")
            prijemkaSummary = ET.SubElement(prijemka, 'pri:prijemkaSummary')
            homeCurrency = ET.SubElement(prijemkaSummary, 'pri:homeCurrency')
            priceNone = ET.SubElement(homeCurrency, 'typ:priceNone')
            priceNone.text = '0'
            priceLow = ET.SubElement(homeCurrency, 'typ:priceLow')
            priceLow.text = '0'
            priceLowVAT = ET.SubElement(homeCurrency, 'typ:priceLowVAT')
            priceLowVAT.text = '0'
            priceLowSum = ET.SubElement(homeCurrency, 'typ:priceLowSum')
            priceLowSum.text = '0'
            round_elem = ET.SubElement(homeCurrency, 'typ:round')
            priceRound = ET.SubElement(round_elem, 'typ:priceRound')
            priceRound.text = '0'
            
            xml_str = ET.tostring(dataPack, encoding='windows-1250', method='xml').decode('windows-1250')
            response_file_path = os.path.join(self.output_dir, f"{cislo_faktury}_Prijemka.xml")
            print(f"\nXML požadavek na vytvoření Příjemky ukládám do souboru \"{response_file_path}\"")
            FileHelper.zapis_do_souboru(response_file_path, xml_str, 'windows-1250')
            
            print(f"\nNa základě faktury {cislo_faktury} importuji příjemku")
            uspech, status_code = self.odesli_pozadavek_serveru(xml_str, cislo_faktury)
            if uspech and status_code == 200:
                return True
            else:
                return False
        
        prijemkaAccessoryChargesItem = ET.SubElement(prijemkaDetail, 'pri:prijemkaAccessoryChargesItem')
        quantity = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:quantity')
        quantity.text = '1.0'
        payVAT = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:payVAT')
        payVAT.text = 'false'
        rateVAT = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:rateVAT')
        rateVAT.text = 'high'
        discountPercentage = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:discountPercentage')
        discountPercentage.text = '0.0'
        
        homeCurrency = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:homeCurrency')
        
        unit_price_value = str(shipping_df['Částka celkem'].iloc[0])
        print(f"Hodnota vedlejších nákladů (Dopravy) je: {unit_price_value}")
        
        unitPrice = ET.SubElement(homeCurrency, 'typ:unitPrice')
        unitPrice.text = unit_price_value
        priceVAT = ET.SubElement(homeCurrency, 'typ:priceVAT')
        
        dopravaBezDPH = MathHelper.preved_na_desetinne_cislo(unit_price_value)
        dopravaDPH = dopravaBezDPH * Decimal('0.21')
        dopravaCelkem = dopravaBezDPH + dopravaDPH
        dopravaDPHstr = str(MathHelper.preved_na_desetinne_cislo(f"{dopravaDPH:.2f}"))
        dopravaCelkem = str(dopravaCelkem)
        
        priceVAT.text = dopravaDPHstr
        priceSum = ET.SubElement(homeCurrency, 'typ:priceSum')
        priceSum.text = dopravaCelkem
        note = ET.SubElement(prijemkaAccessoryChargesItem, 'pri:note')
        note.text = 'Dopravné'
        
        prijemkaSummary = ET.SubElement(prijemka, 'pri:prijemkaSummary')
        roundingDocument = ET.SubElement(prijemkaSummary, 'pri:roundingDocument')
        roundingDocument.text = 'none'
        roundingVAT = ET.SubElement(prijemkaSummary, 'pri:roundingVAT')
        roundingVAT.text = 'none'
        typeCalculateVATInclusivePrice = ET.SubElement(prijemkaSummary, 'pri:typeCalculateVATInclusivePrice')
        typeCalculateVATInclusivePrice.text = 'VATNewMethod'
        
        homeCurrency = ET.SubElement(prijemkaSummary, 'pri:homeCurrency')
        priceNone = ET.SubElement(homeCurrency, 'typ:priceNone')
        priceNone.text = '0'
        priceLow = ET.SubElement(homeCurrency, 'typ:priceLow')
        priceLow.text = '0'
        priceLowVAT = ET.SubElement(homeCurrency, 'typ:priceLowVAT')
        priceLowVAT.text = '0'
        priceLowSum = ET.SubElement(homeCurrency, 'typ:priceLowSum')
        priceLowSum.text = '0'
        round_elem = ET.SubElement(homeCurrency, 'typ:round')
        priceRound = ET.SubElement(round_elem, 'typ:priceRound')
        priceRound.text = '0'
        
        xml_str = ET.tostring(dataPack, encoding='windows-1250', method='xml').decode('windows-1250')
        response_file_path = os.path.join(self.output_dir, f"{cislo_faktury}_Prijemka.xml")
        print(f"\nXML požadavek na vytvoření Příjemky ukládám do souboru \"{response_file_path}\"")
        FileHelper.zapis_do_souboru(response_file_path, xml_str, 'windows-1250')
        
        print(f"\nNa základě faktury {cislo_faktury} importuji příjemku")
        uspech, status_code = self.odesli_pozadavek_serveru(xml_str, cislo_faktury)
        if uspech and status_code == 200:
            return True
        else:
            return False


class InvoiceWorkflow:
    """Hlavní třída pro orchestraci celého procesu zpracování faktury"""
    
    def __init__(self):
        self.root = Tk()
        self.root.withdraw()
        
        self.jmeno_faktury = askopenfilename(
            title='Zvol soubor s PDF fakturou STIHL',
            filetypes=[('PDF files', '*.pdf')]
        )
        
        if not self.jmeno_faktury:
            print("Nebyl vybrán žádný soubor s fakturou!")
            sys.exit(1)
        
        self.output_dir = os.path.dirname(self.jmeno_faktury) + '/'
        self.pdf_parser = PDFParser(self.jmeno_faktury)
    
    def run(self):
        """Spustí celý workflow zpracování faktury"""
        print(f"Zpracovávám fakturu \"{self.jmeno_faktury}\"")
        print(f"v adresáři \"{self.output_dir}\"")
        
        objednavka = self.pdf_parser.rozdel_fakturu_podle_objednavek()
        
        order_processor = OrderProcessor(self.output_dir)
        receipt_manager = ReceiptManager(self.output_dir)
        stock_manager = StockManager(self.output_dir)
        
        # Slovník pro ukládání DataFrame shod pro každou objednávku
        shody_dfs = {}
        
        # Zpracuj objednávky a uložení shod
        for cislo_objednavky, order_lines in objednavka.items():
            pdf_parser_temp = PDFParser(self.jmeno_faktury)
            df = pdf_parser_temp.extrahuj_produktove_informace()
            df_filtered = pd.DataFrame()
            
            for line in order_lines:
                for _, row in df.iterrows():
                    # Přidej řádek, pokud obsahuje číslo produktu NEBO je to dopravné (None)
                    if pd.isna(row['Číslo produktu']):
                        # Je to dopravné - přidej ho
                        df_filtered = pd.concat([df_filtered, pd.DataFrame([row])], ignore_index=True)
                        break
                    elif str(row['Číslo produktu']) in line or str(row['Název produktu']) in line:
                        df_filtered = pd.concat([df_filtered, pd.DataFrame([row])], ignore_index=True)
                        break
            
            if df_filtered.empty:
                df_filtered = df
            
            df_prijemka, text_objednavky = order_processor.zpracuj_objednavku(cislo_objednavky, df_filtered)
            
            if isinstance(df_prijemka, pd.DataFrame):
                shody_dfs[cislo_objednavky] = df_prijemka
                
                # Ulož shody do Excel
                output_filename = f"{self.output_dir}{cislo_objednavky}-shody.xlsx"
                try:
                    df_prijemka.to_excel(output_filename, index=False)
                    print(f"Tabulka shod uložena do: {output_filename}")
                except Exception as e:
                    print(f"Chyba při ukládání shod: {e}")
        
        # Načti informace o zásobách pro všechny objednávky
        if shody_dfs:
            print("\n" + "="*60)
            print("NAČÍTÁNÍ INFORMACÍ O ZÁSOBÁCH Z POHODA")
            print("="*60)
            
            # Spojení všech DataFrame shod do jednoho
            df_vsechny_shody = pd.concat(shody_dfs.values(), ignore_index=True)
            
            # Načti informace o zásobách
            df_zasoby = stock_manager.nacti_zasoby_podle_objednacich_id(df_vsechny_shody)
            
            if not df_zasoby.empty:
                print("\n✅ Úspěšně načteny informace o zásobách")
                print(f"Celkem načteno: {len(df_zasoby)} položek")
            else:
                print("\n⚠️ Nepodařilo se načíst informace o zásobách")
        
        # Pokračuj ve vytváření příjemky
        print("\n" + "="*60)
        print("VYTVÁŘENÍ PŘÍJEMKY")
        print("="*60)
        
        vysledek_zpracovani = receipt_manager.zpracuj_vsechny_objednavky(
            self.jmeno_faktury,
            objednavka,
            order_processor
        )
        
        if vysledek_zpracovani:
            print("\nZpracování faktury proběhlo v pořádku.")
        else:
            print("\nZpracování faktury skončilo s chybou.")
        
        subprocess.Popen(r'\\POHODA\Pohoda\Pohoda.exe /http stop "Firma"', shell=True)
        input("Stiskněte Enter pro ukončení...")


def main():
    """Hlavní vstupní bod programu"""
    try:
        status = MServerStatus.zjisti_stav_serveru()
        print(f"mServer je spuštěný")
        print(f"Počet požadavků ve frontě ke zpracování: {status['processing']}")
    except Exception as e:
        MServerInitializer.inicializace_mServeru()
    
    workflow = InvoiceWorkflow()
    workflow.run()


if __name__ == "__main__":
    main()