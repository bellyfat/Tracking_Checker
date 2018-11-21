# Tracking_Checker
Tracking_Checker.py
This purpose of the script is to scan emails for shipping email sent by various vendors (Newegg, BestBuy, Amazon, etc.), extract the order number, shipment date, shipping address and tracking number.
The script will also determine which of the major carriers the tracking number belongs to, navigate to their website and lookup tracking information and estimated delivery date on the carrier website. The script does so by using the function “check_carrier_name_and_add_record”. 
This function takes in several parameters including the order number, tracking number and shipping address.
This information can then be posted to a database via REST API and an excel sheet containing the information can be generated as well.


