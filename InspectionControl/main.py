#InspectionControl - проверяет наличие файлов и видеоматериалов в папках

import os
import pandas as pd
import time

start = time.time()
month = 'май 2024'
path = 'C:\\Users\\mde_KrivonosovDA\\Desktop\\Нормативы\\05. май\\'
save_path = 'C:\\Users\\mde_KrivonosovDA\\Desktop\\01. Нормативы личного участия по ОТ\\01. Проверки наличия\\Проверки {month}.xlsx'
   
df_check = pd.DataFrame(columns = ["ЭЧ", "Руководитель", "Норматив", "Где проводились", "Наличие видео ОП"])  
df_normativ = pd.DataFrame(columns =["ЭЧ", "Руководитель", "Норматив", "Наличие материалов"])


def create_line(path):
    line = []
    Flag = False
    for ech in os.scandir(path):
        ech_path = os.path.join(path, ech.name)
        for person in os.scandir(ech_path):
            person_path = os.path.join(ech_path, person.name)
            for normativ in os.scandir(person_path):
                normativ_path = os.path.join(person_path, normativ.name)
                check_count = 0
                if normativ.name == '1. Оперативные проверки':
                    for check in os.scandir(normativ_path):
                        check_path = os.path.join(normativ_path, check.name)
                        if os.path.isdir(check_path): # если не папка - нужно пропустить!
                            for files in os.scandir(check_path):
                                if os.path.splitext(files.path)[1] in ('.mov', '.avi', '.mp4', '.mpeg','.MP4', '.MOV', 'AVI','MPEG'):
                                    Flag = True
                            if Flag == True:
                                line = [ech.name, person.name, normativ.name, check.name,  1]
                                check_count += 1 # есть видео
                            else:
                                if check.name != '01.08 ЭЧК-№':
                                    line = [ech.name, person.name, normativ.name, check.name,  0]
                                    check_count += 1# нет видео
                            Flag = False
                            df_check.loc[len(df_check)] = line # Добавляем в df_check
                    while check_count < 3:
                        line = [ech.name, person.name, normativ.name, 'Нет проверки',  0]
                        check_count += 1
                        df_check.loc[len(df_check)] = line  # Добавляем в df_check отсутствующие проверки, для ЭЧ и остальных
                    if (check_count < 4) and (('ЭЧ ' not in str(person.name)) or ('ЭЧ-% ' not in str(person.name))):
                        line = [ech.name, person.name, normativ.name,'Нет проверки',  0]
                        check_count += 1
                        df_check.loc[len(df_check)] = line # Добавляем в df_check отсутствующие проверки для всех, кроме ЭЧ
                else:   # Остальные нормативы
                    if len(os.listdir(normativ_path)) > 0:
                        line = [ech.name, person.name, normativ.name, 1]
                    else:
                        line = [ech.name, person.name, normativ.name, 0]
                    df_normativ.loc[len(df_normativ)] = line # Добавляем к нормативам


create_line(path)

with pd.ExcelWriter(save_path) as writer:
    df_check.to_excel(writer, sheet_name = 'Оперативные', index = False)
    df_normativ.to_excel(writer, sheet_name = 'Нормативы', index = False)

finish  = time.time()-  start

print(f'Completed for {finish} seconds!')
