import shutil
import zipfile
import os

try:
    file = str(input())
    os.rename(file, f'{file}.zip')
    zip_file_path = f'{file}.zip'
except FileNotFoundError:
    print('неверное имя файла')
    breakpoint()


def modify_files_in_zip(zip_file_path):
    files_list = []
    os.makedirs('temp_dir', exist_ok=True)
    # Открываем архив для чтения и записи
    with zipfile.ZipFile(zip_file_path, 'a') as zip_ref:
        # Получаем список всех файлов в архиве
        zip_ref.extractall('temp_dir')
        for i in os.listdir('temp_dir/xl/worksheets'):
            files_list.append(f'temp_dir/xl/worksheets/{i}')
        files_list.append('temp_dir/xl/workbook.xml')
        for file in files_list:
            # with open('temp_dir/xl/workbook.xml', 'r') as file_name:
            with open(file, 'r') as file_name:
                content = file_name.read()
                if file == 'temp_dir/xl/workbook.xml':
                    modified_content = modify_content(content, '<workbookProtection', 1)
                else:
                    modified_content = modify_content(content, '<sheetProtection', 0)
                # print(modified_content)
            # with open('temp_dir/xl/workbook.xml', 'w') as file_name:
            with open(file, 'w') as file_name:
                file_name.write(modified_content)
                file_name.close()
    with zipfile.ZipFile(zip_file_path, 'w') as zip_ref:
        for root, dirs, files in os.walk('temp_dir'):
            for file in files:
                file_path = os.path.join(root, file)
                zip_ref.write(file_path, os.path.relpath(file_path, 'temp_dir'))
    shutil.rmtree('temp_dir')
    os.rename(zip_file_path, f'{zip_file_path.replace(".zip", "")}')


def modify_content(content, param_1, last_file):
    if 'Protection' in content:
        first_index = content.find(param_1)
        last_index = content.find('>', first_index)
        new_content = content[:first_index] + content[last_index + 1:]
        return new_content
    elif last_file == 1:
        print('шифрование не обнаружено')
        shutil.rmtree('temp_dir')
        os.rename(zip_file_path, f'{zip_file_path.replace(".zip", "")}')
        breakpoint()
    else:
        pass


# Пример использования функци

modify_files_in_zip(zip_file_path)
