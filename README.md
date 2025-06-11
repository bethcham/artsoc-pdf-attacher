# pdf-attacher
Takes a pdf and splits it into individual pages, then attaches each page to a personalised email. Created to do ticketing for ArtSoc at Imperial College London.

Inspired by
[seat-distributor](https://github.com/Tuna521/seat-distributor/blob/main/distributor.py)
and [auto-email_sender](https://github.com/Tuna521/auto-email-sender)

Some warnings:
1. You must ensure that the order of the tickets in `tickets.pdf` matches in `shop.csv`. This is **usually** in alphanumerical order, but not always. 
2. Manually add seats in the format 'Seat 1', 'Seat 2', ... as separate columns. This is case sensitive. [To be added: debug seat distributor]