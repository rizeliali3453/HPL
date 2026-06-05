import os
import time
import uno
from com.sun.star.beans import PropertyValue


REPO_DIR = os.getcwd()
EXCEL_PATH = os.path.join(REPO_DIR, "HPL_Cost_Master.xlsx")
CSV_PATH = os.path.join(REPO_DIR, "HPL_Prices.csv")

DEBUG_ZERO_LIMIT = 10
debug_zero_count = 0


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect_libreoffice():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx
    )

    ctx = None

    for _ in range(30):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
            )
            break
        except Exception:
            time.sleep(1)

    if ctx is None:
        raise Exception("LibreOffice bağlantısı kurulamadı.")

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop


def set_cell(sheet, address, value):
    cell = sheet.getCellRangeByName(address)

    if isinstance(value, str):
        cell.String = value
    else:
        cell.String = ""
        cell.Value = float(value)


def read_price(cell):
    value = cell.Value

    if value is not None and value > 0:
        return float(value)

    text = str(cell.String)
    text = text.replace("€", "")
    text = text.replace(" ", "")
    text = text.replace("\xa0", "")
    text = text.strip()

    # Avrupa formatı: 11.589,62 -> 11589.62
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return 0.0


def unit_code(model, direction, supply, ret, cooling, outdoor, preheater, bms):
    return (
        "HPL"
        + f"{int(model):02d}"
        + str(direction)
        + f"{int(supply):02d}"
        + f"{int(ret):02d}"
        + str(cooling)
        + str(int(outdoor))
        + str(int(preheater))
        + str(int(bms))
    )


def write_inputs(sheet, model, direction, supply, ret, cooling, outdoor, preheater, bms):
    set_cell(sheet, "B2", model)
    set_cell(sheet, "C2", direction)
    set_cell(sheet, "D2", supply)
    set_cell(sheet, "E2", ret)

    if cooling == "C":
        set_cell(sheet, "F2", "C")
    else:
        set_cell(sheet, "F2", 0)

    set_cell(sheet, "G2", outdoor)
    set_cell(sheet, "H2", preheater)
    set_cell(sheet, "I2", bms)


def add_price_line(doc, sheet, rows, model, direction, supply, ret, cooling, outdoor, preheater, bms):
    global debug_zero_count

    write_inputs(sheet, model, direction, supply, ret, cooling, outdoor, preheater, bms)

    try:
        doc.enableAutomaticCalculation(True)
    except Exception:
        pass

    doc.calculateAll()
    time.sleep(0.2)

    price_cell = sheet.getCellRangeByName("P3")
    price = read_price(price_cell)

    code = unit_code(model, direction, supply, ret, cooling, outdoor, preheater, bms)

    if price <= 0:
        if debug_zero_count < DEBUG_ZERO_LIMIT:
            print(
                "DEBUG ZERO PRICE | "
                + f"Code={code} | "
                + f"P3.Value={price_cell.Value} | "
                + f"P3.String='{price_cell.String}'"
            )
            debug_zero_count += 1
        return

    rows.append({
        "UnitCode": code,
        "Price": f"{price:.2f}",
        "Currency": "EUR"
    })


def main():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel dosyası bulunamadı: {EXCEL_PATH}")

    desktop = connect_libreoffice()

    file_url = uno.systemPathToFileUrl(os.path.abspath(EXCEL_PATH))

    doc = desktop.loadComponentFromURL(
        file_url,
        "_blank",
        0,
        (
            prop("Hidden", True),
            prop("ReadOnly", False),
            prop("UpdateDocMode", 3),
        )
    )

    try:
        sheet = doc.Sheets.getByName("MIF")

        rows = []

        models = [7, 11, 15, 22, 32, 42, 53, 64, 90]
        directions = ["R", "L"]
        bms_values = [0, 1]

        for model in models:
            if model in [7, 11, 15]:
                for direction in directions:
                    for bms in bms_values:
                        add_price_line(
                            doc, sheet, rows,
                            model, direction,
                            7, 5, "0", 0, 0, bms
                        )
            else:
                supply_filters = [47, 49, 5]
                return_filters = [5]
                cooling_values = ["0", "C"]
                outdoor_values = [0, 1]
                preheater_values = [0, 1]

                for direction in directions:
                    for supply in supply_filters:
                        for ret in return_filters:
                            for cooling in cooling_values:
                                for outdoor in outdoor_values:
                                    for preheater in preheater_values:
                                        for bms in bms_values:
                                            add_price_line(
                                                doc, sheet, rows,
                                                model, direction,
                                                supply, ret, cooling,
                                                outdoor, preheater, bms
                                            )

        if len(rows) == 0:
            raise RuntimeError(
                "Hiç fiyat satırı üretilemedi. "
                "Excel formülleri GitHub/LibreOffice ortamında hesaplanamamış olabilir. "
                "Bu nedenle HPL_Prices.csv dosyası güncellenmedi."
            )

        with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
            f.write("UnitCode;Price;Currency\n")

            for row in rows:
                f.write(
                    str(row["UnitCode"]) + ";" +
                    str(row["Price"]) + ";" +
                    str(row["Currency"]) + "\n"
                )

        print(f"HPL_Prices.csv oluşturuldu. Kayıt sayısı: {len(rows)}")

    finally:
        doc.close(True)


if __name__ == "__main__":
    main()
