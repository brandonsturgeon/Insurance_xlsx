# Used to format crop insurance data into a dynamically created spreadsheet

import xlsxwriter
import psycopg2


# Converts Row,Col notation to LetterNum notation
def rc_to_ln(r, c):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[c]+str(r+1)


# Converts Row,Col,Row,Col range notation to LetterNum:LetterNum notation
def rc_to_ln_range(r, c, r1, c1):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[c]+str(r+1)+":"+letters[c1]+str(r1+1)


# Main creation class
class Create():
    def __init__(self, name, data):
        self.name = name
        self.data = data

        # Creates the actual file
        self.workbook = xlsxwriter.Workbook(self.name+".xlsx")

        # Formats #
        self.unlocked = self.workbook.add_format({"locked": 0})
        self.bold = self.workbook.add_format({"bold": True})
        self.underline = self.workbook.add_format({"underline": True})

        # Bolded, Bordered, Grey, Centered horizontally and vertically
        self.format_01 = self.workbook.add_format({"bold": True,
                                                   "border": 1,
                                                   "fg_color": "#555555",
                                                   "align": "center",
                                                   "valign": "vcenter"})
        # Bolded, Bordered, Centered H+Z
        self.format_02 = self.workbook.add_format({"bold": True,
                                                   "border": 1,
                                                   "align": "center",
                                                   "valign": "vcenter"})
        # Bolded, Top Border
        self.format_03 = self.workbook.add_format({"bold": True,
                                                   "top": 1,
                                                   "hidden": 1})

        self.main()

    # Builds all pages
    def main(self):
        # Creates and formats the worksheets
        self.make_policy_info(self.data["policy_info"])

        if "enterprise_units" in self.data:
            self.make_enterprise_units(self.data["enterprise_units"])
        elif "optional_units" in self.data:
            self.make_optional_units(self.data["optional_units"])

        if "hpp_units" in self.data:
            self.make_hpp_units(self.data["hpp_units"])

        self.workbook.close()

    # Creates and formats the Policy Information sheet
    def make_policy_info(self, data):
        page = self.workbook.add_worksheet()
        page.protect()

        # Row/Column sizes
        page.set_row(0, 25)

        # Total width should be 175, we merge them later so each needs to be ~half
        page.set_column(0, 1, 25)

        # Header
        page.merge_range("A1:B1", "Insurance Policy Info", self.format_01)

        # Table
        # Walks through the key,value pairs and writes a list of the two to a row
        r = 1
        for k, v in data.iteritems():
            page.write_row(r, 0, [k, v])
            r += 1

    # Creates and formats the Enterprise Unit sheet
    def make_enterprise_units(self, data):
        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:G1", "Enterprise Units", self.format_01)
        page.merge_range("A3:B3", self.data["policy_info"]["County"], self.format_02)

        # Setting Column Width
        page.set_column(2, 0, 15)
        page.set_column(3, 0, 15)
        page.set_column(4, 0, 15)

        # General info
        page.write_row(3, 1, data["gen"].keys(), self.format_01)
        page.write_row(4, 1, data["gen"].values())

        # Row counter
        r = 6
        for name, unit in data["units"].iteritems():
            page.write(r, 1, name, self.format_01)
            # The general information for this unit
            page.write_row(r+1, 2, unit["gen"].keys(), self.format_01)
            page.write_row(r+2, 2, unit["gen"].values())

            # ! This will have to change, because dictionaries aren't ordered, which means .keys() won't be ! #
            # Sets our headers for the zone data
            headers = unit["zones"][0].keys()
            page.write_row(r+4, 2, headers, self.format_01)
            r += 5

            # Writes the data for each zone in the unit
            for zone in unit["zones"]:
                page.write_row(r, 2, zone.values())
                r += 1

            # Total calculations
            page.write(r, 2, "Totals: ", self.format_03)
            for c in range(4, 6):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.unlocked)
            r += 3

    # Formats the Optional Units sheet
    def make_optional_units(self, data):
        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:H1", "Optional Units", self.format_01)

        # Row counter
        r = 2
        for name, unit in data["units"].iteritems():
            page.write(r, 0, name, self.format_01)

            # General unit information
            page.write_row(r+1, 1, unit["gen"].keys())
            page.write_row(r+2, 1, unit["gen"].values())

            # Sets the headers for our zone columns
            headers = unit["zones"][0].keys()
            page.write_row(r+3, 3, headers, self.format_01)

            r += 5
            # Creates the table by just writing each zone's values as a list
            for zone in unit["zones"]:
                page.write_row(r, 1, zone.values())
                r += 1

            # Total calculation
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(3, 5):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

    # Formats the HPP Units sheet
    def make_hpp_units(self, data):
        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:F1", "HPP Units", self.format_01)

        page.set_column(1, 0, 15)
        page.set_column(2, 0, 15)
        page.set_column(3, 0, 15)

        # Row counter
        r = 2
        for name, unit in data["units"].iteritems():
            page.write(r, 0, name, self.format_01)

            # General unit information
            page.write_row(r+1, 1, unit["gen"].keys(), self.format_01)
            page.write_row(r+2, 1, unit["gen"].values())

            # Sets the headers for our zone columns
            headers = unit["zones"][0].keys()
            page.write_row(r+4, 1, headers, self.format_01)

            r += 5
            # Creates the table by just writing each zone's values as a list
            for zone in unit["zones"]:
                page.write_row(r, 1, zone.values())
                r += 1

            # Total calculations
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(3, 5):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.unlocked)
            r += 3

if __name__ == "__main__":
    doing_connections = True

    dic = {"policy_info":
                {"Crop": "Corn",
                 "County": "Adair, IA",
                 "Units": "optional",
                 "MPCI Coverage": "60%",
                 "Practice": "non_irrigated"},
            "enterprise_units": {"gen": {"Total": 493, "Total2": 86700, "Total3": 19720},
                                 "units": {"unit1": {"gen": {"totalacres": 10,
                                                             "total2": 70,
                                                             "total3": 32},
                                                     "zones": [{"Field-Zone": "zone1",
                                                                "Acres": 200,
                                                                "Actual Production": 275000,
                                                                "Actual Yield": 550}]},
                                           "unit2": {"gen": {"totalacres": 20,
                                                             "total2": 50,
                                                             "total3": 75},
                                                     "zones": [{"Field-Zone": "zone1",
                                                                "Acres": 200,
                                                                "Actual Production": 275000,
                                                                "Actual Yield": 550}]}}},
            "hpp_units": {"units": {"unit1": {"gen": {"total acre": 720, "modified APH": 550},
                                              "zones": [{"Field-Zone": "zone1",
                                                                "Acres": 200,
                                                                "Actual Production": 275000,
                                                                "Actual Yield": 550}]}}}}

    if doing_connections:

        # Lookups
        market_symbols = {
            "alfalfa": ["alfalfa"],
            "cane": ["cane"],
            "corn": ["corn_enogen", "corn_enogen_dryland", "corn_white",
                     "corn_white_dryland", "corn_yellow", "corn_yellow_dryland"],
            "corn_pink": ["corn_pink"],
            "cotton": ["cotton"],
            "oats": ["oats", "oats_dryland"],
            "soybeans": ["soybeans", "soybean_dryland", "soybean_meal",
                         "soybean_meal_dryland", "soybean_oil", "soybean_oil_dryland"],
            "wheat": ["wheat", "wheat_dryland", "wheat_red",
                      "wheat_red_dryland", "wheat_spring", "wheat_spring_dryland"]
        }

        # DB Connection
        conn = psycopg2.connect("dbname=DB user=brandonsturgeon password=brandon1 host=localhost")
        cur = conn.cursor()

        # Creates object with policy info
        policy_id = "24"
        headers = ["id", "farm_id", "farm_crop_id", "units", "combined_market_symbol"]

        # Converts headers list into a string to plug into a query
        t_str = str(headers).replace("'", "").strip("[]")

        # Query to get our policy dictionary
        a = "SELECT " + t_str + " FROM insurances WHERE id = %s;"
        cur.execute(a, (policy_id,))
        policy = dict(zip(headers, cur.fetchone()))
        print policy

        # Generates the Array to be used in farm_crop query
        m_symb = policy["combined_market_symbol"]
        t_str = "ANY(ARRAY"
        t_str += str(market_symbols[m_symb]).replace("\"", "'") + ")"
        print t_str

        # Gets the farm_crop IDs with same farm_id and market symbols
        a = "SELECT DISTINCT farm_crops.id " \
            "FROM farm_crops, crops " \
            "WHERE crops.market_symbol = " + t_str + \
            "AND farm_crops.crop_id = crops.id " \
            "AND farm_crops.farm_id = %s;"
        cur.execute(a, (policy["farm_id"],))
        farm_crops = [x[0] for x in cur.fetchall()]

        # Finds zones
        t_str = "ANY(ARRAY["
        t_str += str(farm_crops).strip("[]") + "])"
        print t_str
        a = "SELECT DISTINCT ON (zones.id) zones " \
            "FROM insurances, farms, fields, zones " \
            "WHERE insurances.id = %s " \
            "AND farms.id = insurances.farm_id " \
            "AND fields.farm_id = farms.id " \
            "AND zones.field_id = fields.id " \
            "AND zones.county_id = insurances.county_id " \
            "AND zones.farm_crop_id = " + t_str + ";"
        cur.execute(a, (policy_id,))
        zones = cur.fetchall()
        print len(zones)

    else:
        # Otherwise just use pre-created model
        Create("test_file", dic)

    #insurance county ID
    #zones county ID
