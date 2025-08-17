from csv import reader
import os
import tempfile
import shutil
import re
import subprocess

def create_mail_draft(subject, body_html, to_address, attachments):
    applescript = f'''
    set theSubject to "{subject}"
    set theContent to "{body_html.replace('"', "'")}"
    set theAddress to "{to_address}"
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject:theSubject, content:theContent, visible:true}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address:theAddress}}
    '''
    for attachment in attachments:
        applescript += f'\n            make new attachment with properties {{file name:"{os.path.abspath(attachment)}"}} at after the last paragraph'
    applescript += '''
        end tell
        delay 1
        save newMessage
    end tell
    '''
    subprocess.run(['osascript', '-e', applescript])

def render_email(template_content, show_name, location, committee, theatre_time, meet_up, library_time=None, committee2=None):
    email = template_content
    email = email.replace('{ShowName}', show_name)
    email = email.replace('{Location}', location)
    email = email.replace('{Committee}', committee)
    email = email.replace('{Time}', theatre_time)
    if meet_up:
        email = email.replace('{Time2}', library_time)
        email = email.replace('{Committee2}', committee2)
    return email

def natural_sort_key(s): # makes sure the order is A1, A2, ... , A10 etc
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
    # print("Files will be attached in this order:")
    # for f in files:
    #     print(f"  {f}")
    # print()
    return files

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# sender email address
sender_email = "artsoc@imperial.ac.uk" 

# specify folder with ticket PDFs
ticket_folder = os.path.join(os.path.dirname(__file__), 'output_tickets')
ticket_folder = os.path.abspath(ticket_folder)

ticket_files = get_sorted_ticket_files(ticket_folder)
current_ticket_index = 0

# take inputs from user
show_name = input("Enter show name for email subject title (eg Phantom of the Opera): ")
location = input("What theatre?: ")
theatre_time = input("What time is the meet up for show? Leave 20 min before show starts (enter for default 7:10pm): ")
if not theatre_time:
    theatre_time = "7:10pm"  # default time if not specified
meet_up = input_yes_no("Campus meet-up? (yes/no): ")

subject = f"{show_name} tickets!"


if meet_up:
    library_time = input("What time is the meet-up? Take the time to travel + 5 min waiting at library + 5 min buffer: ")
    committee = input("Which committee members will be at the library? (phrase like 'Beth, Justin, and Anh): ")
    committee2 = input("Which committee members will meet at the theatre?: ")
    template_file = open("templates/email_library.html", "r") 
else:
    committee = input("Which committee members will be there?: ")
    template_file = open("templates/email_theatre.html", "r") # ensure email template is correct and updated

general_email = template_file.read()

# general_email = general_email.replace('{ShowName}', show_name)
# general_email = general_email.replace('{Location}', location)
# general_email = general_email.replace('{Committee}', committee)
# general_email = general_email.replace('{Time}', theatre_time)

# if meet_up:
#     general_email = general_email.replace('{Time2}', library_time)
#     general_email = general_email.replace('{Committee2}', committee2)

while True:
    print("Final email draft: ")
    if meet_up:
        print(render_email(general_email, show_name, location, committee, theatre_time, meet_up, library_time, committee2))
    else:
        print(render_email(general_email, show_name, location, committee, theatre_time, meet_up))
    change_template = input_yes_no("Would you like to change the email? (yes/no)")
    if not change_template:
        break

    else:
        what_change = input("What would you like to change? /n a) Show name /n b) Location /n c) Committee members /n d) Theatre time: /n e) Library time (if applicable)").strip().lower()
        if what_change == 'a':
            show_name = input("Enter new show name: ")
        elif what_change == 'b':
            location = input("Enter new theatre location: ")
        elif what_change == 'c' and meet_up:
            committee = input("Enter new committee members at library: ")
            committee2 = input("Enter new committee members at theatre: ")
        elif what_change == 'c' and not meet_up:
            committee = input("Enter new committee members: ")
        elif what_change == 'd':
            theatre_time = input("Enter new theatre time: ")
        elif what_change == 'e' and meet_up:
            library_time = input("Enter new library time: ")

general_email = general_email.replace('{ShowName}', show_name)
general_email = general_email.replace('{Location}', location)
general_email = general_email.replace('{Committee}', committee)
general_email = general_email.replace('{Time}', theatre_time)

if meet_up:
    general_email = general_email.replace('{Time2}', library_time)
    general_email = general_email.replace('{Committee2}', committee2)

customer_info_file = open("shop.csv")
customer_info_reader = reader(customer_info_file)
customer_info_header = next(customer_info_reader)
num_customer_header = len(customer_info_header)

email_count = 0

total_tickets_needed = 0
for customer_info in reader(open("shop.csv")):
    try:
        quantity_index = customer_info_header.index("Quantity")
        total_tickets_needed += int(customer_info[quantity_index])
    except Exception:
        # produce error and stop execution
        print("Warning: error in customer info in quantity field: " + str(customer_info))

if total_tickets_needed != len(ticket_files):
    raise ValueError(f"ERROR: {len(ticket_files)} tickets available, but {total_tickets_needed} needed.")

available_tickets = ticket_files.copy()

# custom email for each
for customer_info in customer_info_reader:
    custom_email = general_email
    num_tickets = 1  # default to 1 ticket
    seat_numbers = []

    # replace with custom name and email for each
    for i in range(num_customer_header):
        if (customer_info_header[i] == "Email"):
            address = customer_info[i]
        elif (customer_info_header[i] == "First Name"):
            name = customer_info[i]
        elif (customer_info_header[i] == "Quantity"):
            try:
                num_tickets = int(customer_info[i])
            except (ValueError, TypeError):
                num_tickets = 1
        elif customer_info_header[i].lower().startswith("seat"):
            seat_numbers.append(customer_info[i])

        custom_email = custom_email.replace(
            '{' + customer_info_header[i] + '}', customer_info[i])
        
    if name.lower() not in address.lower():
        print(f"Warning: Name '{name}' does not appear in email '{address}'.")
    if "noreply" in address.lower():
        address = ""

    # Add attachments based on number of tickets
    attached_files = []
    for seat in seat_numbers[:num_tickets]:
        # Find ticket file matching both name and seat
        found = False
        for ticket_file in available_tickets:
            if name.lower() in ticket_file.lower() and seat and str(seat) in ticket_file:
                ticket_path = os.path.join(ticket_folder, ticket_file)
                if os.path.exists(ticket_path):
                    attached_files.append(ticket_file)
                    available_tickets.remove(ticket_file)
                    current_ticket_index += 1
                    found = True
                    break
        if not found:
            print(f"Warning: No ticket found for {name} (seat {seat})")

    # saves as draft - go to apple mail to check and send
    create_mail_draft(subject, custom_email, address, attached_files)
    email_count += 1
    print(f"created draft email {email_count} for {address}")
    print(f"  attached tickets: {', '.join(attached_files)}")

template_file.close()
customer_info_file.close()

print(f"\nFinished creating {email_count} draft emails")
print(f"Used {current_ticket_index} tickets out of {len(ticket_files)} available")