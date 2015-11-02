#!/usr/bin/python

import csv
import itertools
import string
import sys

from datetime import datetime

class CSVMapper:

    VENDORS = ["Adidas", "Brooks", "Diadora", "Hummel", "Kelme", "Kwik Goal", "Mueller", "Nike", "Nomis", "Puma", "Reebok", "Uhlsport", "Umbro"]

    # (Shopify, Volusion) CSV structure mappings
    CUSTOMERS_SHOPIFY = [("First Name","firstname"),("Last Name","lastname"),("Email","emailaddress"),("Company","companyname"),("Address1","billingaddress1"),
                         ("Address2","billingaddress2"),("City","city"),("Province","state"),("Province Code",""),("Country","country"),("Country Code",""),
                         ("Zip","postalcode"),("Phone","phonenumber"),("Accepts Marketing","emailsubscriber"),("Total Spent",""),("Total Orders",""),("Tags",""),
                         ("Note","customer_notes"),("Tax Exempt","")]

    PRODUCTS_SHOPIFY = [("Handle","productname"),("Title","productname"),("Body (HTML)","productdescription"),("Vendor","productname"),("Type","categorytree"),("Tags",""),
                        ("Published","hideproduct"),("Option1 Name",""),("Option1 Value",""),("Option2 Name",""),("Option2 Value",""),("Option3 Name",""),
                        ("Option3 Value",""),("Variant SKU",""),("Variant Grams","productweight"),("Variant Inventory Tracker",""),("Variant Inventory Quantity",""),
                        ("Variant Inventory Policy",""),("Variant Fulfillment Service",""),("Variant Price","saleprice"),("Variant Compare at Price","productprice"),
                        ("Variant Requires Shipping",""),("Variant Taxable","taxableproduct"),("Variant Barcode","upc_code"),("Variant Weight Unit",""),
                        ("Image Src","photourl"),("Image Alt Text","photo_alttext"),("Gift Card",""),("SEO Title","productname"),("SEO Description","productname"),
                        ("Google Shopping / MPN",""),("Google Shopping / Age Group",""),("Google Shopping / Gender",""),("Google Shopping / Google Product Category",""),
                        ("Google Shopping / Adwords Grouping",""),("Google Shopping / Adwords Labels",""),("Google Shopping / Condition",""),
                        ("Google Shopping / Custom Product",""),("Variant Image","photourl"),("Collection","categorytree")]

    def __init__(self, args):
        self.importtype = args[1]
        self.shopifyfile = self.importtype + '-' + datetime.now().isoformat('-') + ".csv"            # File to import into Shopify
        self.shopifyCSV = None
        if self.importtype == "customers":
            self.customersfile = args[2]
            self.customersCSV = None
        elif self.importtype == "products":
            self.productsfile = args[2]
            self.productsCSV = None
            self.optionsfile = args[3]
            self.optionsCSV = None
            self.optioncatsfile = args[4]
            self.optioncatsCSV = None
        # Open files for reading/writing
        self.openall()

    # Initiate mapping of CSV files
    def mapCSV(self):
        if self.importtype == "customers":
            self.map_customers()
        elif self.importtype == "products":
            self.map_products()
        else:
            print("Import type must invalid.")
        self.closeall()

    # Opens all required CSV files
    def openall(self):
        self.shopifyCSV = open(self.shopifyfile, 'wt')
        if self.importtype == "customers":
            self.customersCSV = open(self.customersfile, 'rt')
        elif self.importtype == "products":
            self.productsCSV = open(self.productsfile, 'rt')
            self.optionsCSV = open(self.optionsfile, 'rt')
            self.optioncatsCSV = open(self.optioncatsfile, 'rt')

    # Closes all open CSV files
    def closeall(self):
        self.shopifyCSV.close()
        if self.importtype == "customers":
            self.customersCSV.close()
        elif self.importtype == "products":
            self.productsCSV.close()
            self.optionsCSV.close()
            self.optioncatsCSV.close()

    # Map Volusion product data to Shopify product data
    def map_products(self):
        try:
            reader = csv.DictReader(self.productsCSV)
            optionreader = csv.DictReader(self.optionsCSV)
            optioncatreader = csv.DictReader(self.optioncatsCSV)
            writer = csv.writer(self.shopifyCSV)

            # Write headers to Shopify CSV
            headers = []
            for item in self.PRODUCTS_SHOPIFY:
                headers.append(item[0])
            writer.writerow(headers)

            # Map Volusion product fields to Shopify product fields
            reader.next()
            for row in reader:                                                          # For every Volusion product
                mappedvals = []
                for key, val in self.PRODUCTS_SHOPIFY:                                  # For every Shopify mapping
                    if key == "Handle" and val == "productname":                        # Product handle
                        prodhandle = row[val].lower()
                        for char in prodhandle:
                            if char not in "abcdefghijklmnopqrstuvwxyz0123456789-":
                                prodhandle = prodhandle.replace(char, "-")
                        mappedvals.append(prodhandle)
                    elif key == "Vendor":
                        matchvendor = list(set(self.VENDORS).intersection(set(row[val].split())))
                        if matchvendor:
                            mappedvals.append(matchvendor[0])
                        else:
                            mappedvals.append("")
                    elif key == "Variant Price":                                        # Fix for items that are not on sale
                        if not row["saleprice"]:
                            mappedvals.append(row["productprice"])
                        else:
                            mappedvals.append(row["saleprice"])
                    elif val == "hideproduct":                                          # Product publish
                        if row[val] == "Y":
                            mappedvals.append("FALSE")
                        else:
                            mappedvals.append("TRUE")
                    elif val == "taxableproduct":                                       # Taxable products
                        if row[val] == "Y":
                            mappedvals.append("TRUE")
                        else:
                            mappedvals.append("FALSE")
                    elif val == "productweight":                                        # Pound to gram conversion
                        if row[val]:
                            mappedvals.append(float(row[val]) * 454)
                        else:
                            mappedvals.append(0)
                    elif key == "Variant Inventory Policy":                             # Deny sales when inventory reaches 0
                        mappedvals.append("deny")
                    elif key == "Variant Inventory Quantity":                           # All 0 initially
                        mappedvals.append(0)
                    elif key == "Variant Fulfillment Service":                          # Manual fulfillment for imports
                        mappedvals.append("manual")
                    elif key == "Gift Card":                                            # Product is gift card?
                        mappedvals.append("false")
                    elif key == "Google Shopping / Condition":                          # Google product condition
                        mappedvals.append("new")
                    elif key == "Google Shopping / Custom Product":                     # Google unique product ID
                        mappedvals.append("FALSE")
                    elif (key == "Collection" or key == "Type") and val == "categorytree":    # Categories to Collections
                        if row[val]:
                            mappedvals.append(row[val].split("> ").pop())
                        else:
                            mappedvals.append("")
                    elif key and val:
                        mappedvals.append(row[val])
                    else:
                        mappedvals.append("")

                # Determine if product has variants (size, color, etc.)
                optionvaldict = {}
                if row["optionids"]:
                    optionids = row["optionids"].split(",")
                    for optionid in optionids:
                        self.optionsCSV = open(self.optionsfile, 'rt')
                        optionreader = csv.DictReader(self.optionsCSV)
                        for row in optionreader:
                            if row["id"] == optionid:
                                optioncatid = row["optioncatid"]
                                optionval = row["optionsdesc"]
                                self.optionsCSV.close()
                                self.optioncatsCSV = open(self.optioncatsfile, 'rt')
                                optioncatreader = csv.DictReader(self.optioncatsCSV)
                                for catrow in optioncatreader:
                                    if catrow["id"] == optioncatid:
                                        optioncat = catrow["optioncategoriesdesc"]
                                        self.optioncatsCSV.close()
                                        if not optionvaldict.has_key(optioncat):
                                            optionvaldict[optioncat] = [optionval]
                                        else:
                                            optionvaldict[optioncat].append(optionval)
                                        break                               # Got the option description
                                break
                    # Add a new variant line for every permutation of optioncat and optionval
                    optioncombos = list(itertools.product(*optionvaldict.values()))
                    for index, combo in enumerate(optioncombos):
                        # Add new variant line in Shopify import with optioncat and optionval
                        mappedvals[7] = optionvaldict.keys()[0]                     # Option1 Name
                        mappedvals[8] = combo[0]                                    # Option1 Value
                        if len(combo) > 1:
                            mappedvals[9] = optionvaldict.keys()[1]                 # Option2 Name
                            mappedvals[10] = combo[1]                               # Option2 Name
                            if len(combo) > 2:
                                mappedvals[11] = optionvaldict.keys()[2]            # Option3 Value
                                mappedvals[12] = combo[2]                           # Option3 Value
                        # Skip "Title", "Body (HTML)", "Vendor" and "Tags" for all but first variant
                        if index > 0:
                            mappedvals[1] = ""
                            mappedvals[2] = ""
                            mappedvals[3] = ""
                            mappedvals[5] = ""
                        # Write variant to Shopify import
                        writer.writerow(mappedvals)
                else:
                    writer.writerow(mappedvals)

        finally:
            self.productsCSV.close()
            self.shopifyCSV.close()
            self.optionsCSV.close()
            self.optioncatsCSV.close()

    # Map Volusion customer data to Shopify customer data
    def map_customers(self):
        try:
            reader = csv.DictReader(self.customersCSV)
            writer = csv.writer(self.shopifyCSV)
            # Write headers to Shopify CSV
            headers = []
            for item in self.CUSTOMERS_SHOPIFY:
                headers.append(item[0])
            writer.writerow(headers)

            # Map Volusion customer fields to Shopify customer fields
            reader.next()
            for row in reader:
                mappedvals = []
                for key,val in self.CUSTOMERS_SHOPIFY:
                    if val:
                        # Massage equivalent terms
                        if row[val] == "Y":
                            mappedvals.append("yes")
                        else:
                            mappedvals.append(row[val])
                    else:
                        mappedvals.append("")
                writer.writerow(mappedvals)

        finally:
            self.customersCSV.close()
            self.shopifyCSV.close()

# Print the script usage
def usage():
    print "Usage: "
    print "\tpython cartmapper.py [ \"customers\" | \"products\" ] [ <customer_csv> | <product_csv> <options_csv> <optioncategories_csv> ]"
    sys.exit(0)

# Check command-line args for correctness
def args():
    if len(sys.argv) < 3:
        usage()
    elif sys.argv[1] not in ["customers", "products"]:
        usage()

def main():
    args()
    mapper = CSVMapper(sys.argv)
    mapper.mapCSV()

if __name__ == "__main__":
    main()
