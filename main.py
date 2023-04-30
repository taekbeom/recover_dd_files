import pandas as pd
import openpyxl
import os

bytes_size = 512

file_names = ['Archive', 'Audio', 'Documents', 'Graphic', 'Video']

for file_name in file_names:

    df = pd.read_excel('xlsx_files/L5_' + file_name + '.xlsx')

    df = df[df['File'] != 'Fill (Заполнение)']

    df['Start Sector'] = df['Start Sector'].str.replace(',', '')
    df['End Sector'] = df['End Sector'].str.replace(',', '')
    df['File Size'] = df['File Size'].str.replace(',', '')

    file_chunks = []

    for file_drop in df['File']:
        substr = file_drop.split('(')[0].strip()
        file_chunks.append(substr)

    file_chunks = list(set(file_chunks))

    os.mkdir('output/L5_' + file_name)

    for file_chunk in file_chunks:

        sections = []

        for index, row in df.iterrows():

            if file_chunk in row['File']:

                sections.append((int(row['Start Sector'])*bytes_size,
                                 int(row['File Size'])))

                # sections.append((int(row['Start Sector']) * bytes_size,
                #                  (int(row['End Sector']) - int(row['Start Sector']) + 1) * bytes_size))

        with open('dd_files/L5_' + file_name + '.dd', 'rb') as input_file:

            with open('output/L5_' + file_name + '/' + file_chunk, 'wb') as output_file:

                for (offset, size) in sections:

                    input_file.seek(offset)
                    data = input_file.read(size)
                    output_file.write(data)
