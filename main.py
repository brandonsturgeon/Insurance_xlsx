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
            z_headers = unit["zones"][0].keys()
            page.write_row(r+4, 2, z_headers, self.format_01)
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
            z_headers = unit["zones"][0].keys()
            page.write_row(r+3, 3, z_headers, self.format_01)

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
            z_headers = unit["zones"][0].keys()
            page.write_row(r+4, 1, z_headers, self.format_01)

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


class Generate():
    def __init__(self):
        self.dictionary = {}
        self.main()

    def main(self):
#        Lookups
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

        data_set = {"policy_info":
                    {"Crop": "Corn",
                     "County": "Adair, IA",
                     "Units": "optional",
                     "MPCI Coverage": "60%",
                     "Practice": "non_irrigated"}}

        # Creates object with policy info
        policy_id = "24"
        headers = ["id", "farm_id", "farm_crop_id", "units",
                   "combined_market_symbol", "hpp_coverage",
                   "units", "county_id", "practice", "hpp_practice", "mpci_coverage",
                   "percent_of_spring_price"]

        # Converts headers list into a string to plug into a query
        t_str = str(headers).replace("'", "").strip("[]")

        # Query to get our policy dictionary
        a = "SELECT " + t_str + " FROM insurances WHERE id = %s;"
        cur.execute(a, (policy_id,))
        policy = dict(zip(headers, cur.fetchone()))

        # Setting words to be the same as what we use for zones
        for k, v in policy.iteritems():
            if v == "irrigated":
                policy[k] = True
            elif v == "non irrigated":
                policy[k] = False
        #print "Policy: " + str(policy_id) + " " + str(policy)

        # Puts HPP coverage shell into the data set if our policy has HPP
        # policy["hpp_coverage"] is either an integer if it exists, or None if it doesn't
        usable_units = []
        if policy["hpp_coverage"] is not None:
            data_set["hpp_units"] = {"units": {}}
            usable_units.append("hpp_units")

        # Puts either enterprise_units or optional_units shell into data set
        u = policy["units"]+"_units"
        data_set[u] = {"units": {}}
        usable_units.append(u)
        print "Usable units: " + str(usable_units)

        # Generates the crop name Array to be used in farm_crop query
        m_symb = policy["combined_market_symbol"]
        t_str = "ANY(ARRAY"
        t_str += str(market_symbols[m_symb]).replace("\"", "'") + ")"
       # print "Crop Names: " + t_str

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
        #print "Farm Crop IDs: " + t_str

        hpp = "SELECT DISTINCT zones.id " \
              "FROM insurances, farms, fields, zones " \
              "WHERE insurances.id = %s " \
              "AND farms.id = insurances.farm_id " \
              "AND fields.farm_id = farms.id " \
              "AND zones.field_id = fields.id " \
              "AND zones.county_id = insurances.county_id " \
              "AND zones.irrigated = " + str(policy["hpp_practice"]) + " " \
              "AND zones.farm_crop_id = " + t_str + ";"
        cur.execute(hpp, (policy_id,))
        hpp_zones = [x[0] for x in cur.fetchall()]

        optional = "SELECT DISTINCT zones.id " \
                   "FROM insurances, farms, fields, zones " \
                   "WHERE insurances.id = %s " \
                   "AND farms.id = insurances.farm_id " \
                   "AND fields.farm_id = farms.id " \
                   "AND zones.field_id = fields.id " \
                   "AND zones.county_id = insurances.county_id " \
                   "AND zones.irrigated = " + str(policy["practice"]) + " " \
                   "AND zones.farm_crop_id = " + t_str + ";"
        cur.execute(optional, (policy_id,))
        optional_zones = [x[0] for x in cur.fetchall()]

        enterprise = "SELECT DISTINCT zones.id " \
                     "FROM insurances, farms, fields, zones " \
                     "WHERE insurances.id = %s " \
                     "AND farms.id = insurances.farm_id " \
                     "AND fields.farm_id = farms.id " \
                     "AND zones.field_id = fields.id " \
                     "AND zones.county_id = insurances.county_id " \
                     "AND zones.farm_crop_id = " + t_str + ";"
        cur.execute(enterprise, (policy_id,))
        enterprise_zones = [x[0] for x in cur.fetchall()]

        # General information for each legal unit
        unit_gens = {"hpp_units": ["total_acres", "modified_aph", "mcpi_yield_guarantee",
                                   "covered_bushels", "guarantee/acre", "loss%",
                                   "potential_bushel_loss", "potential_$_loss", "actual_$_loss"],
                     "optional_units": ["total_acres", "APH", "yield_guarantee",
                                        "guarantee/acre", "total_bu_guarantee", "mcpi_bu_loss/acre", "mcpi_loss"],
                     "enterprise_units": ["total_acres", "APH", "yield_guarantee",
                                          "guarantee/acre", "total_bu_guarantee", "mcpi_bu_loss/acre", "mcpi_loss"]}

        check_pre = [(hpp_zones, "hpp_units"),
                     (enterprise_zones, "enterprise_units"),
                     (optional_zones, "optional_units")]
        check_l = [x for x in check_pre if x[1] in usable_units]

        for page in check_l:
            print page
            a = "SELECT (zones.section, zones.township, zones.range), array_to_string(array_agg(zones.id), ',') " \
                "FROM zones " \
                "WHERE zones.id = ANY(ARRAY[" + str(page[0]) + "]) " \
                "GROUP BY (zones.section, zones.township, zones.range);"
            cur.execute(a)
            legals = cur.fetchall()

            # Generates the units (legal definitions) for this sheet
            units = {}

            for u in legals:
                total_acres = 0
                actual_production_total = 0
                # Turns the string of (Int, North/South, East/West) into a proper legal name
                legal_name = u[0].strip("()").split(",")
                name = "Unit - " + str(legal_name[0]) + " " + str(legal_name[1]) + " " + str(legal_name[2])

                # Gets the general info for this section from unit_gens, and then sets them all to 0 in a dictionary
                unit_gen = unit_gens[page[1]]
                gens = dict(zip(unit_gen, [0 for x in range(len(unit_gen))]))
                units[name] = {"gen": gens, "zones": [int(x) for x in u[1].split(",")]}

                new_l = []
                for k, x in enumerate(units[name]["zones"]):
                    a = "SELECT fields.name, zones.name, zones.acres, " \
                        "zones.yield_goal, zones.fsa_acres, zones.loss_percent, zones.aph, zones.id " \
                        "FROM fields, zones " \
                        "WHERE zones.id = " + str(x) + " " \
                        "AND zones.field_id = fields.id;"
                    cur.execute(a)
                    result = cur.fetchone()
                    name_key = result[0] + "-" + result[1]

                    # Math Stuff
                    # Actual Yield = Yield_goal - (yield_goal * loss_percent)
                    actual_yield = result[3] - (result[3] * (result[5]/100))

                    # Actual Production = actual_yield * fsa_acres
                    actual_production = actual_yield * result[4]

                    new_l.append({"Field-Zone": name_key,
                                  "Acres": result[2],
                                  "Actual Production": actual_production,
                                  "Actual Yield": actual_yield})
                    total_acres += int(result[2])
                    actual_production_total += actual_production

                gen_dict = units[name]["gen"]
                gen_dict["total_acres"] = total_acres
                gen_dict["mpci_yield_guarantee"] = result[6] * (policy["mpci_coverage"] / 100)

                q = "SELECT farm_crops.harvest_price_cents, farm_crops.spring_price_cents " \
                    "FROM zones,farm_crops " \
                    "WHERE zones.id = " + str(result[7]) + " " \
                    "AND farm_crops.id = zones.farm_crop_id;"
                cur.execute(q)
                res = cur.fetchone()
                harvest_price = res[0]
                spring_price = res[1]

                # If we're working with Optional or Enterprise
                if page[1] != "hpp_units":
                    gen_dict["aph"] = result[6]
                    gen_dict["guarantee/acre"] = spring_price * gen_dict["mpci_yield_guarantee"]
                    gen_dict["total_bu_guarantee"] = gen_dict["mpci_yield_guarantee"] * total_acres

                    if harvest_price < spring_price:
                        a = gen_dict["guarantee/acre"] / harvest_price
                    else:
                        a = gen_dict["mpci_yield_guarantee"]
                    trigger_yield = a

                    gen_dict["MPCI_bu_loss/acre"] = trigger_yield - (actual_production_total / total_acres)
                    gen_dict["MPCI_loss"] = harvest_price * gen_dict["MPCI_bu_loss/acre"] * total_acres

                # If we're working with HPP
                else:
                    percent_spring_price = spring_price * (policy["percent_of_spring_price"] / 100)

                    gen_dict["modified_aph"] = result[6] * (policy["hpp_coverage"] / 100)
                    gen_dict["covered_bushels"] = gen_dict["modified_aph"] - result[6] * (policy["mpci_coverage"] / 100)
                    gen_dict["guarantee/acre"] = percent_spring_price * gen_dict["covered_bushels"]
                    gen_dict["loss%"] = result[5] / 100

                    total_bushel_loss = gen_dict["modified_aph"] * gen_dict["loss%"]
                    if gen_dict["covered_bushels"] > total_bushel_loss:
                        gen_dict["potential_bushel_loss"] = total_bushel_loss
                    else:
                        gen_dict["potential_bushel_loss"] = gen_dict["covered_bushels"]

                    a = percent_spring_price * gen_dict["potential_bushel_loss"] * gen_dict["total_acres"]
                    gen_dict["potential_$_loss"] = a

                    total_actual_yield = actual_production_total / total_acres
                    if total_actual_yield > gen_dict["modified_aph"] - gen_dict["potential_bushel_loss"]:
                        if gen_dict["modified_aph"] - total_actual_yield < gen_dict["potential_bushel_loss"]:
                            a = percent_spring_price * (gen_dict["modified_aph"] - total_actual_yield) * total_acres
                        else:
                            a = percent_spring_price * gen_dict["potential_bushel_loss"] * total_acres
                    else:
                        a = gen_dict["potential_$_loss"]
                    gen_dict["actual_$_loss"] = a

                # Adds the list of zones for this unit
                units[name]["zones"] = new_l
            data_set[page[1]]["units"] = units
        self.dictionary = data_set

if __name__ == "__main__":
    doing_connections = True

    dic = {"policy_info":
          {"Crop": "Corn",
           "County": "Adair, IA",
           "Units": "optional",
           "MPCI Coverage": "60%",
           "Practice": "non_irrigated"},

           "enterprise_units":
           {"gen": {"Total": 493, "Total2": 86700, "Total3": 19720},
            "units": {"unit1": {"gen": {"total_acres": 10, "total2": 70, "total3": 32},
                                "zones": [{"Field-Zone": "zone1",
                                           "Acres": 200,
                                           "Actual Production": 275000,
                                           "Actual Yield": 550}]},
                      "unit2": {"gen": {"total_acres": 20, "total2": 50, "total3": 75},
                                "zones": [{"Field-Zone": "zone1",
                                           "Acres": 200,
                                           "Actual Production": 275000,
                                           "Actual Yield": 550}]}}},
           "hpp_units":
           {"units": {"unit1": {"gen": {"total acre": 720, "modified APH": 550},
                                "zones": [{"Field-Zone": "zone1",
                                           "Acres": 200,
                                           "Actual Production": 275000,
                                           "Actual Yield": 550}]}}}}

    if doing_connections:
        Create("test_file2", Generate().dictionary)
    else:
        Create("test_file", dic)
