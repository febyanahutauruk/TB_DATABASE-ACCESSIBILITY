"""
Nama: CIKY FEBYANA TAMARIA HUTAURUK 
NIM: 1202210411
KELAS: SI-45-03
"""

print("TUGAS BESAR_ALGORITMA PEMROGRAMAN")
print("===== Program Database Accessibility PT.CENDEKIA =====")


def mainmenu():
    print(""" >>> Main Menu <<<
1. Pendataan informasi pribadi karyawan.
2. Pendataan jumlah karayawan .
3. Mengupdate data karyawan.
4. Melihat statistik data karyawan(Jenis Kelamin).
5. Exit.
""")
pilihan =int(input("Masukkan pilihan anda : "))
match pilihan:
    case 1:
        print('=== Pendataan Informasi Pribadi Karyawan PT.CENDEKIA ===')
        import PySimpleGUI as sg
        import pandas as pd

        sg.theme("DarkGreen4")

        Excelfile = 'Database.xlsx'

        df = pd.read_excel(Excelfile)

        layout = [
        [sg.Text('Masukkan Data Karyawan:')],
        [sg.Text('Nama', size=(20,1)), sg.InputText(key='Nama')],
        [sg.Text('Jenis Kelamin', size=(20,1)), sg.Combo(['pria', 'wanita'],key='Jenis Kelamin')],
        [sg.Text('Usia', size=(20,1)), sg.InputText(key='Usia')],
        [sg.Text('Alamat', size=(20,1)), sg.Multiline(key='Alamat')],
        [sg.Text('No.Telp', size=(20,1)), sg.InputText(key='No.Telp')],
        [sg.Text('Status', size=(20,1)), sg.InputText(key='Status')],
        [sg.Text('Jumlah anak', size=(20,1)), sg.Combo(['belum ada','1','2','3','4','5'],key='Jumlah anak')],
        [sg.Text('Gaji Pokok', size=(20,1)), sg.Combo(['Rp1.000.000','Rp.1.500.000','Rp3.000.000','Rp.5.200.000','Rp6.800.000','Rp10.000.000'],key='Gaji Pokok')],

        [sg.Submit(), sg.Button('clear'), sg.Exit()]

        ]

        window = sg.Window('Database', layout)

        def clear_input():
            for key in values:
                window[key]('')
                return None

        while True:
            event, values = window.read()
            if event == 'Exit':
                break
            if event == 'Clear':
                clear_input
            if event == 'Submit':
                df = df.append(values, ignore_index = True)
                df.to_excel(Excelfile, index = False)
                sg.popup('Data Berhasil Di Simpan')
                clear_input()
 
        window.close()

    case 2:
        print("=== Jumlah Karyawan PT.CENDEKIA ===")
        import pandas as pd

        jumlah = 10
        df = pd.read_excel('Database.xlsx')

        for i in df.index:
            print(f"Nama Karyawan ke- {i+1}: "+df['Nama'][i])
            
        print(f'Jumlah Karyawan PT.CENDEKIA adalah {jumlah} orang')

        
    case 3:
        print("=== Rekap Data Karyawan PT.CENDEKIA ===")
        import pandas as pd
        df = pd.read_excel('Database.xlsx')
        lst=[]

        for i in df.index:
            data = {}
            data['Nama'] = df['Nama'][i]
            data['Jenis Kelamin'] = df['Jenis Kelamin'][i]
            data['Usia'] = df['Usia'][i]
            data['Alamat'] = df['Alamat'][i]
            data['No.Telp'] = df['No.Telp'][i]
            data['Status'] = df['Status'][i]
            data['Jumlah anak'] = df['Jumlah anak'][i]
            data['Gaji Pokok'] = df['Gaji Pokok'][i]

            lst.append(data)
            print(lst)
    
    case 4:
        print("=== Statistik Data Karyawan PT.CENDEKIA BERDASARKAN JENIS KELAMIN ===")
        import matplotlib.pyplot as plt
        import numpy as np
        import pandas as pd

        data = pd.read_excel('Database.xlsx')
        data
        df = pd.read_excel('Database.xlsx')
 
        data_jeniskelamin = df[df['Jenis Kelamin'] == 'Jenis Kelamin']
        data_jeniskelamin = [np.count_nonzero(data['Jenis Kelamin'] == 'wanita'), np.count_nonzero(data['Jenis Kelamin']== 'pria')]
        label = ['wanita','pria']
        plt.title("Perbandingan Jenis Kelamin Karyawan PT.CENDEKIA")
        plt.pie(data_jeniskelamin, labels = label, radius = 1.3, startangle=60, autopct='%.1f%%', shadow=True)
        plt.show()

    case 5:
        print("Selesai")