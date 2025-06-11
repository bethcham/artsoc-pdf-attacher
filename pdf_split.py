import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

def process_pdf(pdf_file, csv_file, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # load csv file (shop.csv)
    df = pd.read_csv(csv_file, encoding='latin1')  # Adjust encoding if needed
    
    # load pdf (tickets.pdf)
    pdf_reader = PdfReader(pdf_file)

    page_index = 0
    
    show_name = input("Enter show name for ticket to be named, use short-form (eg BOM): ")

    used_seats = set()  # Track used seat numbers

    for index, row in df.iterrows():
        quantity = row['Quantity']
        first_name = row['First Name']

        seat_numbers  = []

        for seat in quantity:
            seat_col = f"Seat {seat}"
            seat_numbers.append(row[seat_col].strip())

            if seat in used_seats:
                print(f"Warning: Duplicate seat number '{seat}' for {first_name}")
            else:
                used_seats.add(seat)
            
            if pd.isna(row[seat_col]):
                print(f"Warning: Seat {seat_col} is empty for {first_name}")

            output_pdf(pdf_reader, page_index, first_name, seat_numbers[seat], output_dir, show_name)
            page_index += 1 

            if page_index >= len(pdf_reader.pages):
                raise IndexError("Page index exceeds the number of pages in the PDF.") 

def output_pdf(pdf_reader, page_index, first_name, seat, output_dir, show_name):
    pdf_writer = PdfWriter()
    page = pdf_reader.pages[page_index] 
    pdf_writer.add_page(page)

    # output file name: change to correct show here!!
    output_filename = f"{show_name}_{seat}_{first_name}.pdf"
    output_filepath = os.path.join(output_dir, output_filename)

    # write to a single page pdf
    with open(output_filepath, "wb") as output_pdf_file:
        pdf_writer.write(output_pdf_file)

if __name__ == "__main__":
    process_pdf("tickets.pdf", "shop.csv", "output_tickets") # change these if you need to