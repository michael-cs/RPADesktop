import csv
import os


def cria_csv(file_path):
    header = ['First Name',
              'Last Name',
              'Company Name',
              'Role in Company',
              'Address',
              'Email',
              'Phone Number']

    if os.path.exists(file_path):
        os.remove(file_path)

    with open((file_path), 'w', encoding='UTF8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(header)


def escreve_csv(file_path, row):
    with open((file_path), 'a', encoding='UTF8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(row)
