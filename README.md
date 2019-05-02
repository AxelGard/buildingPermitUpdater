# buildingPermitUpdater

Update a excel file bast on dates

The bot will send a email with the info on when dates is outofdate or close to. 

buildingPermitUpdater was written by Axel Gard in the summer of 2018.
The program was made for municipality of täby in stocholm sweden.
It was written in the programing lanuge python 3.6.

## What dose the program do?

The program gose and looks in the designated excel file.
It looks ate the coloumn 'G' and compres it's to days date.
If the date is has passt it will make the font red and if
the date is with in two months of passing it will make it orange.

Then if it finds any dates it will send an email
with the bulding name and expirison date.

## What I learned 

- [x] openpyxl
- [x] smtplib
- [x] email.mime
