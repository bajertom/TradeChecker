import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import csv

# account credentials
username = "ENTER YOUR EMAIL"
password = "ENTER YOUR PASSWORD"

# create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL("imap.gmail.com")
# authenticate
imap.login(username, password)
print("Logged in!")

status, messages = imap.select("INBOX")
# number of top email to fetch
N = 50

# total number of emails
messages = int(messages[0])
traded_companies = ["ROYAL DUTCH SHELL PLC"]
for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    try:
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email int oa message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    # if it's bytes, decode to str
                    subject = subject.decode()
                # email sender
                from_ = msg.get("From")
                for company in traded_companies:
                    if company in subject:
                        # if the email message is multipar
                        if msg.is_multipart():
                            # iterate over email parts
                            for part in msg.walk():
                                # extract content of email
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content-Disposition"))
                                try:
                                    # get the email body
                                    body = part.get_payload(decode=True).decode()
                                except:
                                    pass
                                if content_type == "text/plain" and "attachment" not in content_disposition:
                                    # saves content into temporary files and skip attachments
                                    path = "C:\\Tom\\test_api_gmail\\" + company + str(i) + '.txt'
                                    print("Creating log file in " + path)
                                    with open(path, "w") as log:
                                        log.write(body)
                                elif "attachment" in content_disposition:
                                    pass

                        else:
                            # extract content type of email
                            content_type = msg.get_content_type()
                            # get the email body
                            body = msg.get_payload(decode=True).decode()
                            if content_type == "text/plain":
                                # print only text email parts
                                print(body)
                                print("Not multipart")
                        print("="*100)
    except:
        pass
imap.close()
imap.logout()
print("Logged out!" + "\n")

# checks if csv file already exists, creates if not
for company in traded_companies:
    if not os.path.exists(company + ".csv"):
        with open(company + ".csv", "w",newline="") as file:
                        fieldnames = ["Buy date", "Sell date", "Volume", "Buy ratio","Buy fee",
                     "Sell ratio", "Sell fee", "Fee", "Buy price", "Sell price",
                     "Diff", "Profit($)", "Profit(€)", "Net(€)", "IDBUY", "IDSELL"]
                        writer = csv.DictWriter(file, fieldnames=fieldnames)
                        writer.writeheader()


# extracts data from temporary files
for company in traded_companies:
    dictionary = {}
    for i in range(messages-N+1, messages+1):
        
        path = "C:\\Tom\\test_api_gmail\\" + company + str(i) + '.txt'
        if os.path.exists(path):
            print("Checking " + path)
            with open(path, "r") as pars:
                compound_volume = 0
                for line in pars:
                    if "Datum" in line:
                        date = line[17:-11]
                    elif "Nákup" in line:
                        action = line[7:-2]
                    elif "Prodej" in line:
                        action = line[7:-2]
                    elif "Počet" in line:
                        volume = int(line[7:-2])
                        compound_volume += volume
                    elif "Cena" in line:
                        price = float((line[10:-2]).replace(",","."))
                    elif "Směnný kurz" in line:
                        ratio = float(line[13:-2].replace(",","."))
                    elif "ID objednávky" in line:
                        ID = line[15:-2]


            if action == "Nákup":
                buy_date = date
                buy_price = price
                buy_id = ID
                buy_ratio = ratio
                buy_fee = round(((0.5 / buy_ratio) + (0.004 * volume)), 2)

            else:
                sell_date = date
                sell_price = price
                sell_id = ID
                sell_ratio = ratio
                sell_fee = round(((0.5 / sell_ratio) + (0.004 * volume)), 2)

            # feed data to dictionary
            dictionary["Volume"] = compound_volume
            try:
                dictionary["Buy fee"] = buy_fee
            except NameError:
                pass
            try:
                dictionary["Sell fee"] = sell_fee
            except NameError:
                pass    
            try:
                dictionary["Buy ratio"] = buy_ratio
            except NameError:
                pass
            try:
                dictionary["Sell ratio"] = sell_ratio
            except NameError:
                pass
            try:
                fee = buy_fee + sell_fee
            except NameError:
                pass
            try:
                diff = abs(round((buy_price - sell_price), 2))
            except NameError:
                pass
            try:
                profitUSD = round(((compound_volume * diff) - fee), 2)
            except NameError:
                pass
            try:
                profitEUR = round((profitUSD * ratio), 2)
            except NameError:
                pass
            try:
                netEUR = round((profitEUR * 0.85), 2)
            except NameError:
                pass
            try:
                dictionary["Buy date"] = buy_date
            except NameError:
                pass
            try:
                dictionary["IDBUY"] = buy_id
            except NameError:
                pass
            try:
                dictionary["Sell date"] = sell_date
            except NameError:
                pass
            try:
                dictionary["IDSELL"] = sell_id
            except NameError:
                pass
            try:
                dictionary["Fee"] = fee
            except NameError:
                pass
            try:
                dictionary["Buy price"] = buy_price
            except NameError:
                pass
            try:
                dictionary["Sell price"] = sell_price
            except NameError:
                pass
            try:
                dictionary["Diff"] = diff
            except NameError:
                pass
            try:
                dictionary["Profit($)"] = profitUSD
            except NameError:
                pass
            try:
                dictionary["Profit(€)"]= profitEUR
            except NameError:
                pass
            try:
                dictionary["Net(€)"] = netEUR
            except NameError:
                pass

            # checks wheter the buy-sell duo is loaded in dictionary 
            if "Sell date" in dictionary and "Buy date" in dictionary:
                
                ID_checker = 0

                # checks whether buy-sell duo is already written
                with open(company + ".csv","r") as file:
                    reader = csv.reader(file)
                    for line in reader:
                        if dictionary["IDSELL"] and dictionary['IDBUY'] in line:
                            ID_checker = 1
                # if not, appends to the file
                if ID_checker == 0:
                    print("Buy-Sell duo found! Appending to " + company + ".csv!")

                    with open(company + ".csv", "a",newline="") as file:
                        fieldnames = ["Buy date", "Sell date", "Volume", "Buy ratio", "Buy fee",
                     "Sell ratio", "Sell fee", "Fee", "Buy price", "Sell price",
                     "Diff", "Profit($)", "Profit(€)", "Net(€)", "IDBUY", "IDSELL"]
                        writer = csv.DictWriter(file, fieldnames=fieldnames)
                        writer.writerow(dictionary)
                # renews the dictionary for new buy-sell duo
                dictionary = {}
                del sell_date
                del buy_date
            # removes temporary files
            print("Removing " + path)
            os.remove(path)
            print("Removed!")
            print("*"*100)


print("Done")
