import os
from csv import reader
from email.message import EmailMessage

def natural_sort_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'([0-9]+)', s)]

def input_yes_no(prompt):
    while True:
        response = input(prompt).strip().lower()
        if response in ['yes', 'no', 'y', 'n']:
            return response == 'yes' or response == 'y'
        print("Please answer with 'yes' or 'no'.")

def get_sorted_ticket_files(ticket_folder):
    files = [f for f in os.listdir(ticket_folder) if f.endswith('.pdf')]
    files.sort(key=natural_sort_key)
    return files

# Paths (adjust as needed)
ticket_folder = "auto-email-sender/pdf-split/output_tickets"
csv_path = "auto-email-sender/pdf-split/shop.csv"
output_email_dir = "auto-email-sender/output_emails"
os.makedirs(output_email_dir, exist_ok=True)

ticket_files = get_sorted_ticket_files(ticket_folder)
available_tickets = ticket_files.copy()

show_name = input("Enter show name for email subject title (eg Phantom of the Opera): ")
location = input("What theatre?: ")
theatre_time = input("What time is the meet up for show? Leave 20 min before show starts (enter for default 7:10pm): ")
if not theatre_time:
    theatre_time = "7:10pm"
meet_up = input_yes_no("Campus meet-up? (yes/no): ")

subject = f"{show_name} tickets!"

if meet_up:
    library_time = input("What time is the meet-up? Take the time to travel + 5 min waiting at library + 5 min buffer: ")
    committee = input("Which committee members will be at the library? (phrase like 'Beth, Justin, and Anh): ")
    committee2 = input("Which committee members will meet at the theatre?: ")
    template_file = open("auto-email-sender/email_library.html", "r")
else:
    committee = input("Which committee members will be there?: ")
    template_file = open("auto-email-sender/email_theatre.html", "r")

general_email = template_file.read()
general_email = general_email.replace('{ShowName}', show_name)
general_email = general_email.replace('{Location}', location)
general_email = general_email.replace('{Committee}', committee)
general_email = general_email.replace('{Time}', theatre_time)

if meet_up:
    general_email = general_email.replace('{Time2}', library_time)
    general_email = general_email.replace('{Committee2}', committee2)

while True:
    print("Final email draft: ")
    print(general_email)
    change_template = input_yes_no("Would you like to change the email? (yes/no)")
    if not change_template:
        break

    else:
        if change_template:
            what_change = input("What would you like to change? /n a) Show name /n b) Location /n c) Committee members /n d) Theatre time: /n e) Library time (if applicable)").strip().lower()
            if what_change == 'a':
                show_name = input("Enter new show name: ")
                general_email = general_email.replace('{ShowName}', show_name)
            elif what_change == 'b':
                location = input("Enter new theatre location: ")
                general_email = general_email.replace('{Location}', location)
            elif what_change == 'c' and meet_up:
                committee = input("Enter new committee members at library: ")
                general_email = general_email.replace('{Committee}', committee)
                committee2 = input("Enter new committee members at theatre: ")
                general_email = general_email.replace('{Committee2}', committee2)
            elif what_change == 'c' and not meet_up:
                committee = input("Enter new committee members: ")
                general_email = general_email.replace('{Committee}', committee)
            elif what_change == 'd':
                theatre_time = input("Enter new theatre time: ")
                general_email = general_email.replace('{Time}', theatre_time)
            elif what_change == 'e' and meet_up:
                library_time = input("Enter new library time: ")
                general_email = general_email.replace('{Time2}', library_time)

customer_info_file = open(csv_path)
customer_info_reader = reader(customer_info_file)
customer_info_header = next(customer_info_reader)
num_customer_header = len(customer_info_header)

email_count = 0

for customer_info in customer_info_reader:
    custom_email = general_email
    address = ''
    name = ''
    num_tickets = 1
    seat_numbers = []

    for i in range(num_customer_header):
        header = customer_info_header[i]
        value = customer_info[i]
        if header == "Email":
            address = value
        elif header == "First Name":
            name = value
        elif header == "Quantity":
            try:
                num_tickets = int(value)
            except (ValueError, TypeError):
                num_tickets = 1
        elif header.lower().startswith("seat"):
            seat_numbers.append(value)
        custom_email = custom_email.replace('{' + header + '}', value)

    if "noreply" in address.lower():
        print(f"Skipping: Email address for {name} contains 'noreply': {address}")
        continue

    if name.lower() not in address.lower():
        print(f"Warning: Name '{name}' does not appear in email '{address}'.")

    # Create the email message
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = "artsoc@imperial.ac.uk"
    msg['To'] = f'"{name}" <{address}>'
    msg.set_content(custom_email, subtype='html')

    # Attach correct tickets by matching name and seat in filename
    attached_files = []
    for seat in seat_numbers[:num_tickets]:
        found = False
        for ticket_file in available_tickets:
            if name.lower() in ticket_file.lower() and seat and str(seat) in ticket_file:
                ticket_path = os.path.join(ticket_folder, ticket_file)
                if os.path.exists(ticket_path):
                    with open(ticket_path, 'rb') as f:
                        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=ticket_file)
                    attached_files.append(ticket_file)
                    available_tickets.remove(ticket_file)
                    found = True
                    break
        if not found:
            print(f"Warning: No ticket found for {name} (seat {seat})")

    # Save as .eml file
    eml_filename = os.path.join(output_email_dir, f"{name}_{address}.eml".replace("/", "_"))
    with open(eml_filename, 'wb') as eml_file:
        eml_file.write(bytes(msg))

    email_count += 1
    print(f"Created email draft {email_count} for {address}")
    print(f"  attached tickets: {', '.join(attached_files)}")

template_file.close()
customer_info_file.close()

print(f"\nFinished creating {email_count} email drafts")