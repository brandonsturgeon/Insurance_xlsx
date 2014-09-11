# Used to format crop insurance data into a dynamically created spreadsheet

import xlsxwriter


# Converts Row,Col notation to LetterNum notation
def rc_to_ln(r, c):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[c]+str(r+1)


# Converts Row,Col,Row,Col range notation to LetterNum:LetterNum notation
def rc_to_ln_range(r, c, r1, c1):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[c]+str(r+1)+":"+letters[c1]+str(r+1)

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
        page.set_row(0, 35)

        # Total width should be 175, we merge them later so each needs to be ~half
        page.set_column(0, 1, 87)

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

        # General info
        page.write(2, 1, data["gen"].keys(), self.format_01)
        page.write(3, 1, data["gen"].values(), self.format_01)

        # Row counter
        r = 6
        for name, unit in data["units"].iteritems():
            page.write(r, 1, name, self.format_01)

            # The general information for this unit
            page.write(r+1, 2, unit["gen"].keys(), self.format_01)
            page.write(r+2, 2, unit["gen"].values())

            # ! This will have to change, because dictionaries aren't ordered, which means .keys() won't be ! #
            # Sets our headers for the zone data
            headers = unit["zones"][0].keys()
            page.write(r+4, 2, headers, self.format_01)
            r += 5

            # Writes the data for each zone in the unit
            for zone in unit["zones"]:
                page.write(r, 2, zone.values())
                r += 1

            # Total calculations
            page.write(r, 2, "Totals: ", self.format_03)
            for c in range(4, 6):
                _range = rc_to_ln_range(r, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
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
            page.write(r+1, 1, unit["gen"].keys())
            page.write(r+2, 1, unit["gen"].values())

            # Sets the headers for our zone columns
            headers = unit["zones"][0].keys()
            page.write(r+3, 3, headers, self.format_01)

            r += 5
            # Creates the table by just writing each zone's values as a list
            for zone in unit["zones"]:
                page.write(r, 1, zone.values())
                r += 1

            # Total calculation
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(3, 5):
                _range = rc_to_ln_range(r, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

    # Formats the HPP Units sheet
    def make_hpp_units(self, data):
        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:J1", "HPP Units", self.format_01)

        # Row counter
        r = 2
        for name, unit in data["units"].iteritems():
            page.write(r, 0, name, self.format_01)

            # General unit information
            page.write(r+1, 1, unit["gen"].keys())
            page.write(r+2, 1, unit["gen"].values())

            # Sets the headers for our zone columns
            headers = unit["zones"][0].keys()
            page.write(r+3, 3, headers, self.format_01)

            r += 5
            # Creates the table by just writing each zone's values as a list
            for zone in unit["zones"]:
                page.write(r, 1, zone.values())
                r += 1

            # Total calculations
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(3, 5):
                _range = rc_to_ln_range(r, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

if __name__ == "__main__":
    dic = {"policy_info":
                {"Crop": "Corn",
                 "County": "Adair, IA",
                 "Units": "optional",
                 "MPCI Coverage": "60%",
                 "Practice": "non_irrigated"},
            "enterprise_units": {"gen": True, "units": {"unit1": {"gen": {"totalacres": 10}}}}}
    Create("test_file", dic)




