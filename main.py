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
    def __init__(self, name, data, verbose):
        self.name = name
        self.data = data
        self.verbose = verbose

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
                                                   "valign": "vcenter",
                                                   "text_wrap": 1})
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
        self.v_print("Creating policy_info sheet..")

        page = self.workbook.add_worksheet()
        page.protect()

        # Row/Column sizes
        page.set_row(0, 25)

        # Total width should be 175, we merge them later so each needs to be ~half
        page.set_column(0, 1, 25)

        # Header
        page.merge_range("A1:B1", "Insurance Policy Info", self.format_01)

        # Table
        r = 1

        # This is the order we want the data to be displayed in, we remove values if they're not in data.keys()
        h_order = ["County", "Units", "MPCI Coverage", "Practice", "HPP Coverage",
                   "HPP Practice", "Harvest Price", "Spring Price", "Percent of Spring Price"]
        h_order = [x for x in h_order if x in data.keys()]

        # Walks through h_order and gets the value or k from data to form a list, which is then written
        for k in h_order:
            page.write_row(r, 0, [k, data[k]])
            r += 1

    # Creates and formats the Enterprise Unit sheet
    def make_enterprise_units(self, data):
        self.v_print("Creating enterprise_units sheet..")

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
            # Sets the headers for our zone columns
            h_order = ["Field-Zone", "Acres", "Actual Production", "Actual Yield"]
            page.write_row(r+4, 1, h_order, self.format_01)

            r += 5
            # Creates the table by just writing each zone's ordered values as a list
            for zone in unit["zones"]:
                t_list = []
                for h in h_order:
                    t_list.append(zone[h])
                page.write_row(r, 1, t_list)
                r += 1

            # Total calculations
            page.write(r, 2, "Totals: ", self.format_03)
            for c in range(4, 6):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

    # Formats the Optional Units sheet
    def make_optional_units(self, data):
        self.v_print("Creating optional_units sheet..")

        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:F1", "Optional Units", self.format_01)

        page.set_column(1, 0, 20)
        page.set_column(2, 0, 20)
        page.set_column(3, 0, 20)
        page.set_column(4, 0, 15)
        page.set_column(5, 0, 5)
        page.set_column(7, 0, 15)
        page.set_column(8, 0, 15)
        page.set_column(9, 0, 15)

        # Row counter
        r = 2
        for name, unit in data["units"].iteritems():
            page.write(r, 0, name, self.format_01)

            # General unit information
            page.write_row(r+1, 1, unit["gen"].keys()[:4], self.format_01)
            page.write_row(r+2, 1, unit["gen"].values()[:4])

            page.write_row(r+4, 1, unit["gen"].keys()[4:], self.format_01)
            page.write_row(r+5, 1, unit["gen"].values()[4:])

            # Sets the headers for our zone columns
            h_order = ["Field-Zone", "Acres", "Actual Production", "Actual Yield"]
            page.write_row(r+8, 1, h_order, self.format_01)

            r += 9
            # Creates the table by just writing each zone's ordered values as a list
            for zone in unit["zones"]:
                t_list = []
                for h in h_order:
                    t_list.append(zone[h])
                page.write_row(r, 1, t_list)
                r += 1

            # Total calculation
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(2, 5):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

    # Formats the HPP Units sheet
    def make_hpp_units(self, data):
        self.v_print("Creating hpp_units sheet..")

        page = self.workbook.add_worksheet()
        page.protect()

        # Header
        page.merge_range("A1:F1", "HPP Units", self.format_01)

        page.set_column(1, 0, 20)
        page.set_column(2, 0, 12)
        page.set_column(3, 0, 15)
        page.set_column(4, 0, 20)
        page.set_column(5, 0, 15)
        page.set_column(7, 0, 15)
        page.set_column(8, 0, 15)
        page.set_column(9, 0, 15)

        # Row counter
        r = 2
        for name, unit in data["units"].iteritems():
            page.write(r, 0, name, self.format_01)

            # General unit information
            page.write_row(r+1, 1, unit["gen"].keys()[:4], self.format_01)
            page.write_row(r+2, 1, unit["gen"].values()[:4])

            page.write_row(r+4, 1, unit["gen"].keys()[4:], self.format_01)
            page.write_row(r+5, 1, unit["gen"].values()[4:])

            # Sets the headers for our zone columns
            h_order = ["Field-Zone", "Acres", "Actual Production", "Actual Yield"]
            page.write_row(r+8, 1, h_order, self.format_01)

            r += 9
            # Creates the table by just writing each zone's ordered values as a list
            for zone in unit["zones"]:
                t_list = []
                for h in h_order:
                    t_list.append(zone[h])
                page.write_row(r, 1, t_list)
                r += 1

            # Total calculations
            page.write(r, 1, "Totals: ", self.format_03)
            for c in range(2, 5):
                _range = rc_to_ln_range(r-1, c, r-len(unit["zones"]), c)
                page.write_formula(r, c, "=SUM("+_range+")", self.format_03)
            r += 3

    # Used for printing status messages if self.verbose is enabled
    def v_print(self, message):
        if self.verbose:
            print message


# Generates our data set to pass over to the Create class
class Generate():
    def __init__(self, verbose):
        self.verbose = verbose
        self.dictionary = {}
        self.main()

    def main(self):
        self.v_print("Beginning data set fabrication..")

        # Market symbol lookups
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
        self.v_print("Attempting database connection..")

        # Attempts database connection, errors out if it fails
        try:
            conn = psycopg2.connect("dbname=DB2 user=brandonsturgeon password=brandon1 host=localhost")
            cur = conn.cursor()
            self.v_print("Database connection successful..")
        except psycopg2.OperationalError as e:
            self.return_error("Something went wrong when trying to connect to the database.", e)

        data_set = {"policy_info":
                    {"Crop": "",
                     "County": "",
                     "Units": "",
                     "MPCI Coverage": "",
                     "Practice": ""}}

        # Creates object with policy info
        policy_id = "20"
        headers = ["id", "farm_id", "farm_crop_id", "units",
                   "combined_market_symbol", "hpp_coverage",
                   "county_id", "practice", "hpp_practice", "MPCI_coverage",
                   "percent_of_spring_price", "county_id"]

        # Converts headers list into a string to plug into a query
        t_str = str(headers).replace("'", "").strip("[]")

        # Query to get our policy dictionary
        a = "SELECT " + t_str + " FROM insurances WHERE id = %s;"
        cur.execute(a, (policy_id,))
        policy = dict(zip(headers, cur.fetchone()))

        # Plugging info into dictionary for the policy_info page
        p_info = data_set["policy_info"]
        p_info["Crop"] = policy["combined_market_symbol"]
        p_info["Units"] = policy["units"]
        p_info["MPCI Coverage"] = str(policy["MPCI_coverage"]) + "%"
        p_info["Practice"] = policy["practice"]
        # Generating the County,State string for "County" key
        a = "SELECT (county_name, state) "\
            "FROM counties "\
            "WHERE id = %s;"
        cur.execute(a, (policy["county_id"],))
        p_info["County"] = cur.fetchone()[0].strip("()")
        p_info["Percent of Spring Price"] = str(policy["percent_of_spring_price"]) + "%"
        # Policy info stuff that only shows up if HPP exists
        if policy["hpp_coverage"] is not None:
            p_info["HPP Coverage"] = str(policy["hpp_coverage"]) + "%"
            p_info["HPP Practice"] = policy["hpp_practice"]
        # Spring Price and Harvest price are set later on in the calculations portion

        # Setting policy words to be the same as what we use for zones
        # We do this after creating the policy_info set because we want "(non-)irrigated" in the policy_info page
        for k, v in policy.iteritems():
            if v == "irrigated":
                policy[k] = True
            elif v == "non irrigated":
                policy[k] = False

        # Puts HPP coverage shell into the data set if our policy has HPP
        # policy["hpp_coverage"] is either an integer if it exists, or None if it doesn't
        usable_units = []
        if policy["hpp_coverage"] is not None:
            data_set["hpp_units"] = {"units": {}}
            usable_units.append("hpp_units")
            self.v_print("HPP_Units exist, adding key to data set..")

        # Puts either enterprise_units or optional_units shell into data set
        u = policy["units"]+"_units"
        self.v_print(u+" exists, adding key to data set..")
        data_set[u] = {"units": {}}
        usable_units.append(u)

        # Generates the crop name Array to be used in farm_crop query
        m_symb = policy["combined_market_symbol"]
        t_str = "ANY(ARRAY"
        t_str += str(market_symbols[m_symb]).replace("\"", "'") + ")"
        # print "Crop Names: " + t_str

        self.v_print("Doing DB lookup to retrieve farm_crop ID's..")
        # Gets the farm_crop IDs with same farm_id and market symbols
        a = "SELECT DISTINCT farm_crops.id " \
            "FROM farm_crops, crops " \
            "WHERE crops.market_symbol = " + t_str + \
            "AND farm_crops.crop_id = crops.id " \
            "AND farm_crops.farm_id = %s;"
        cur.execute(a, (policy["farm_id"],))
        farm_crops = [x[0] for x in cur.fetchall()]

        self.v_print("Creating SQL array of farm_crop ID's..")
        # Creates an string representing an SQL array with all farm_crop_id's that we need to look for specific zones
        t_str = "ANY(ARRAY["
        t_str += str(farm_crops).strip("[]") + "])"

        # This list is used later on to loop through and do unit-specific calculations
        check_l = []

        # Gets the zones for each sheet (*_units) if they're in usable_units
        if "hpp_units" in usable_units:
            self.v_print("Generating hpp_units zones..")

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
            check_l.append((hpp_zones, "hpp_units"))

        if "optional_units" in usable_units:
            self.v_print("Generating optional_units zones..")

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
            check_l.append((optional_zones, "optional_units"))

        elif "enterprise_units" in usable_units:
            self.v_print("Generating enterprise_units zones..")

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
            check_l.append((enterprise_zones, "enterprise_units"))

        # General information shells for each legal unit
        unit_gens = {"hpp_units": ["Total Acres", "Modified APH", "MPCI Yield Guarantee",
                                   "Covered Bushels", "guarantee/acre", "Loss Percent",
                                   "Potential Bushel Loss", "Potential Dollar Loss", "Actual Dollar Loss"],

                     "optional_units": ["Total Acres", "APH", "Yield Guarantee",
                                        "guarantee/acre", "Total Bushel Guarantee",
                                        "MPCI Bushel Loss per acre", "MPCI Loss"],

                     "enterprise_units": ["Total Acres", "APH", "Yield Guarantee",
                                          "guarantee/acre", "Total Bushel Guarantee",
                                          "MPCI Bushel Loss per acre", "MPCI Loss"]}

        self.v_print("Beginning primary calculations loop..")
        # Loops through check_l
        # Inside check_l there are tuples with pairs of (list of zone ids), (string name for the unit)
        for page in check_l:

            # SQL query to get legal names of zones using the list of zone ID's we have in check_l
            a = "SELECT (zones.section, zones.township, zones.range), array_to_string(array_agg(zones.id), ',') " \
                "FROM zones " \
                "WHERE zones.id = ANY(ARRAY[" + str(page[0]) + "]) " \
                "GROUP BY (zones.section, zones.township, zones.range);"
            cur.execute(a)
            legals = cur.fetchall()

            # Generates the units (legal definitions) for this sheet #

            units = {}
            # Looping through all results of the zones query
            for u in legals:
                # Running totals used for "gen" information
                total_acres = 0
                actual_production_total = 0

                # Turns the string of (Int, North/South, East/West) into a proper legal name
                # Strips off the ()'s and splits by commas
                legal_name = u[0].strip("()").split(",")
                name = "Unit - " + str(legal_name[0]) + " " + str(legal_name[1]) + " " + str(legal_name[2])

                # Gets the general info for this section from unit_gens, and then sets them all to 0 in a dictionary
                unit_gen = unit_gens[page[1]]
                gens = dict(zip(unit_gen, [0 for x in range(len(unit_gen))]))
                units[name] = {"gen": gens, "zones": [int(x) for x in u[1].split(",")]}

                new_l = []
                for k, x in enumerate(units[name]["zones"]):
                    # Finds and returns zone information
                    a = "SELECT fields.name, zones.name, zones.acres, " \
                        "zones.yield_goal, zones.fsa_acres, zones.loss_percent, zones.aph, zones.id " \
                        "FROM fields, zones " \
                        "WHERE zones.id = " + str(x) + " " \
                        "AND zones.field_id = fields.id;"
                    cur.execute(a)
                    result = cur.fetchone()

                    # Formats Field Name - Zone Name
                    name_key = result[0] + " - " + result[1]

                    # Math Stuff #
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

                # Generating general parts of the data set used by all units
                gen_dict = units[name]["gen"]
                gen_dict["Total Acres"] = total_acres
                gen_dict["MPCI Yield Guarantee"] = result[6] * (policy["MPCI_coverage"] / 100)

                # Query to get the harvest prices and spring prices (in cents)
                q = "SELECT farm_crops.harvest_price_cents, farm_crops.spring_price_cents " \
                    "FROM zones,farm_crops " \
                    "WHERE zones.id = " + str(result[7]) + " " \
                    "AND farm_crops.id = zones.farm_crop_id;"
                cur.execute(q)
                res = cur.fetchone()
                harvest_price = res[0]
                spring_price = res[1]

                # While we're here, we go ahead and set some more policy attributes to show on the policy_info sheet
                data_set["policy_info"]["Harvest Price"] = "$" + str(harvest_price / 100.0)
                data_set["policy_info"]["Spring Price"] = "$" + str(spring_price / 100.0)

                # Generating unit-specific data
                # If we're working with Optional or Enterprise
                if page[1] != "hpp_units":
                    self.v_print("Generating enterprise and optional calculations..")

                    gen_dict["APH"] = result[6]
                    gen_dict["guarantee/acre"] = spring_price * gen_dict["MPCI Yield Guarantee"]
                    gen_dict["Total Bushel Guarantee"] = gen_dict["MPCI Yield Guarantee"] * total_acres

                    # Calculates trigger_yield, which is used in future calculations
                    if harvest_price < spring_price:
                        a = gen_dict["guarantee/acre"] / harvest_price
                    else:
                        a = gen_dict["MPCI Yield Guarantee"]
                    trigger_yield = a

                    gen_dict["MPCI Bushel Loss per acre"] = trigger_yield - (actual_production_total / total_acres)
                    gen_dict["MPCI Loss"] = harvest_price * gen_dict["MPCI Bushel Loss per acre"] * total_acres

                    # Converting to currency - Doing it after everything because calculations require solid numbers
                    gen_dict["guarantee/acre"] = self.to_currency(gen_dict["guarantee/acre"])
                    gen_dict["MPCI Loss"] = self.to_currency(gen_dict["MPCI Loss"])

                    # Rounding - We round after calculations to ensure accuracy in the calculations
                    gen_dict["MPCI Bushel Loss per acre"] = round(gen_dict["MPCI Bushel Loss per acre"], 2)

                # If we're working with HPP
                else:
                    self.v_print("Generating hpp calculations..")

                    percent_spring_price = spring_price * (policy["percent_of_spring_price"] / 100)

                    gen_dict["Modified APH"] = result[6] * (policy["hpp_coverage"] / 100)
                    gen_dict["Covered Bushels"] = gen_dict["Modified APH"] - result[6] * (policy["MPCI_coverage"] / 100)
                    gen_dict["guarantee/acre"] = percent_spring_price * gen_dict["Covered Bushels"]
                    gen_dict["Loss Percent"] = result[5]

                    # Calculates potential_bushel_loss
                    total_bushel_loss = gen_dict["Modified APH"] * gen_dict["Loss Percent"]
                    if gen_dict["Covered Bushels"] > total_bushel_loss:
                        a = total_bushel_loss
                    else:
                        a = gen_dict["Covered Bushels"]
                    gen_dict["Potential Bushel Loss"] = a

                    # Uses "a" as a temporary variable to save space and abide by PEP8 line-length standards
                    a = percent_spring_price * gen_dict["Potential Bushel Loss"] * gen_dict["Total Acres"]
                    gen_dict["Potential Dollar Loss"] = a

                    # Calculates actual_$_loss
                    total_actual_yield = actual_production_total / total_acres
                    if total_actual_yield > gen_dict["Modified APH"] - gen_dict["Potential Bushel Loss"]:
                        if gen_dict["Modified APH"] - total_actual_yield < gen_dict["Potential Bushel Loss"]:
                            a = percent_spring_price * (gen_dict["Modified APH"] - total_actual_yield) * total_acres
                        else:
                            a = percent_spring_price * gen_dict["Potential Bushel Loss"] * total_acres
                    else:
                        a = gen_dict["Potential Dollar Loss"]
                    gen_dict["Actual Dollar Loss"] = a

                    # Converting to currency - Doing it after everything because calculations require solid numbers
                    gen_dict["guarantee/acre"] = self.to_currency(gen_dict["guarantee/acre"])
                    gen_dict["Potential Dollar Loss"] = self.to_currency(gen_dict["Potential Dollar Loss"])
                    gen_dict["Actual Dollar Loss"] = self.to_currency(gen_dict["Actual Dollar Loss"])

                # Adds the list of zones for this unit
                units[name]["zones"] = new_l

                self.v_print("Finished calculations for " + page[1] + "..")

            # Adds the units for this sheet to the final data_set
            data_set[page[1]]["units"] = units

        self.v_print("Putting data into dictionary attribute..")
        # Finally sets self.dictionary, which is what we use in the Create class
        self.dictionary = data_set

    # Used for printing status messages if self.verbose is enabled
    def v_print(self, message):
        if self.verbose:
            print message

    # Takes a number and converts it to a string with currency formatting
    @staticmethod
    def to_currency(number):
        if type(number) not in (float, int):
            return number

        is_neg = False
        if number < 0:
            is_neg = True
            number *= -1

        a = "%0.2f" % number
        pre = a[:-3]
        post = a[-3:]

        new_l = []
        div = (len(pre) / 3) + 1
        for x in range(div):
            if len(pre) > 3:
                new_l.append(pre[-3:])
                pre = pre[:-3]
            else:
                new_l.append(pre)
        new_l.reverse()
        pre = ",".join(new_l)

        res = "$" + "".join(pre) + post
        if is_neg:
            res = "-" + res

        return res

    # Used to return errors from the data creation process
    @staticmethod
    def return_error(message, error):
        print message
        print "Error message: ", str(error)
        quit()


# Intro function
def main():
    doing_connections = True
    verbose = False

    # Sample data set
    dic = {"policy_info":
          {"Crop": "Corn",
           "County": "Adair, IA",
           "Units": "optional",
           "MPCI Coverage": "60%",
           "Practice": "non_irrigated"},

           "enterprise_units":
           {"gen": {"Total": 493, "Total2": 86700, "Total3": 19720},
            "units": {"unit1": {"gen": {"Total Acres": 10, "total2": 70, "total3": 32},
                                "zones": [{"Field-Zone": "zone1",
                                           "Acres": 200,
                                           "Actual Production": 275000,
                                           "Actual Yield": 550}]},
                      "unit2": {"gen": {"Total Acres": 20, "total2": 50, "total3": 75},
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

    # Generate a data set to use
    if doing_connections:
        Create("test_file2", Generate(verbose).dictionary, verbose)

    # Use the pre-created data set
    else:
        Create("test_file", dic, verbose)

if __name__ == "__main__":
    main()
