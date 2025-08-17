from appscript import app, k
from csv import reader
import tempfile
import shutil
import os

# Connect to Outlook for Mac
outlook = app('Microsoft Outlook')

# Read the template email
with open("email_template.html", "r") as template_file:
    general_email = template_file.read()

# Read general info
with open("General_Info.csv") as gen_info_file:
    gen_info_reader = reader(gen_info_file)
    gen_info_header = next(gen_info_reader)
    gen_info = next(gen_info_reader)

subject = "Test Email"
for i in range(len(gen_info_header)):
    if gen_info_header[i] == "Musical Name":
        subject = gen_info[i] + " tickets!!"
    general_email = general_email.replace('{' + gen_info_header[i] + '}', gen_info[i])

# Read customer info
with open("Purchase_Summary_Dummy.csv") as customer_info_file:
    customer_info_reader = reader(customer_info_file)
    customer_info_header = next(customer_info_reader)
    num_customer_header = len(customer_info_header)

    for customer_info in customer_info_reader:
        custom_email = general_email
        address = ''
        name = ''

        # Replace keywords with custom ones
        for i in range(num_customer_header):
            if customer_info_header[i] == "Email":
                address = customer_info[i]
            elif customer_info_header[i] == "First Name":
                name = customer_info[i]
            custom_email = custom_email.replace('{' + customer_info_header[i] + '}', customer_info[i])

        # Create the draft message
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: subject,
                k.content: custom_email
            })

        # Attach PDF if available (uncomment and set attachment_path as needed)
        # attachment_path = "/Users/youruser/path/to/ticket.pdf"
        # if os.path.exists(attachment_path):
        #     temp_dir = tempfile.mkdtemp()
        #     temp_attachment_path = os.path.join(temp_dir, os.path.basename(attachment_path))
        #     shutil.copyfile(attachment_path, temp_attachment_path)
        #     msg.make(
        #         new=k.attachment,
        #         with_properties={k.file: temp_attachment_path})

        # Add recipient
        msg.make(
            new=k.recipient,
            with_properties={
                k.email_address: {
                    k.name: name,
                    k.address: address
                }
            })

        # Open the draft for review
        msg.open()
        msg.activate()
        print(f"Draft created for {name} <{address}>")
