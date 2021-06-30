import xml.etree.ElementTree as ET
import datetime as dt
import pyexcel as pe

def converter():
    sheet = pe.get_sheet(file_name="name.xlsx", start_row=1, row_limit=-1)

    time = dt.datetime.now().replace(microsecond=0).isoformat() + "+03:00"
    root = ET.Element("realty-feed")
    root.set("xmlns", "http://webmaster.yandex.ru/schemas/feed/realty/2010-06")
    gdate = ET.SubElement(root, "generation-date")
    gdate.text = time
    x = 1 
    for row in sheet:
        offer = ET.SubElement(root, "offer")
        if row[0] == '':
            offer.set("internal-id", str(row[7]) + ' ' + 'оффер '+ str(x))
        else:
            offer.set("internal-id", row[0])

        creation = ET.SubElement(offer, "creation-date")
        creation.text = time
        typer = ET.SubElement(offer, "type")
        typer.text = "продажа"
        category = ET.SubElement(offer, "category")
        category.text = "квартира"
        proptype = ET.SubElement(offer, "property-type")
        proptype.text = "жилая"

        dealstatus = ET.SubElement(offer, "deal-status")
        dealstatus.text = "первичная продажа"
        newflat = ET.SubElement(offer, "new-flat")
        newflat.text = "1"

        salesagent = ET.SubElement(offer, "sales-agent")
        salescat = ET.SubElement(salesagent, "category")
        salescat.text = row[1].strip().lower()
        phone = ET.SubElement(salesagent, "phone")
        phone.text = str(row[2]).strip()

        location = ET.SubElement(offer, "location")
        country = ET.SubElement(location, "country")
        country.text = "Россия"
        locality = ET.SubElement(location, "locality-name")
        locality.text = row[3]
        address = ET.SubElement(location, "address")
        address.text = row[4]

        bid = ET.SubElement(offer, "yandex-building-id")
        bid.text = str(row[5]).strip()
        hid = ET.SubElement(offer, "yandex-house-id")
        hid.text = str(row[6]).strip()

        name = ET.SubElement(offer, "building-name")
        name.text = row[7]
        year = ET.SubElement(offer, "built-year")
        year.text = str(row[8]).strip()
        quart = ET.SubElement(offer, "ready-quarter")
        quart.text = str(row[9]).strip()
        state = ET.SubElement(offer, "building-state")
        state.text = row[10].strip().lower()

        area = ET.SubElement(offer, "area")
        value = ET.SubElement(area, "value")
        value.text = str(row[11]).strip()
        unit = ET.SubElement(area, "unit")
        unit.text = "кв.м"

        if row[12] == '':
            pass
        else:
            livingspace = ET.SubElement(offer, "living-space")
            livingvalue = ET.SubElement(livingspace, "value")
            livingvalue.text = str(row[12]).strip()
            livingunit = ET.SubElement(livingspace, "unit")
            livingunit.text = "кв.м"

        price = ET.SubElement(offer, "price")
        pricevalue = ET.SubElement(price, "value")
        pricevalue.text = str(row[13]).strip()
        currency = ET.SubElement(price, "currency")
        currency.text = "RUR"

        floor = ET.SubElement(offer, "floor")
        floor.text = str(row[14]).strip()
        totalfloors = ET.SubElement(offer, "floors-total")
        totalfloors.text = str(row[15]).strip()
        if row[16] == '':
            pass
        else:
            rooms = ET.SubElement(offer, "rooms")
            rooms.text = str(row[16]).strip()

        image1 = ET.SubElement(offer, "image")
        image1.text = row[17]

        if row[18] == '':
            pass
        else:
            studio = ET.SubElement(offer, 'studio')
            studio.text = str(row[18]).strip()
        
        if row[19] == '':
            pass
        else:
            description = ET.SubElement(offer, "description")
            description.text = str(row[19])

        if row[20] == '':
            aparts = ET.SubElement(offer, 'apartaments')
            aparts.text = '0'         

        else:
            aparts = ET.SubElement(offer, 'apartaments')
            aparts.text = str(row[20]).strip()

        x += 1

    tree = ET.ElementTree(root)
    tree.write("filename.xml", encoding='utf-8')
