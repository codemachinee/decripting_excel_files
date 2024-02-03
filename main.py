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
    os.makedirs('temp_dir', exist_ok=True)
    # Открываем архив для чтения и записи
    with zipfile.ZipFile(zip_file_path, 'a') as zip_ref:
        # Получаем список всех файлов в архиве
        zip_ref.extractall('temp_dir')
        with open('temp_dir/xl/workbook.xml', 'r') as file_name:
            content = file_name.read()
            modified_content = modify_content(content)
            # print(modified_content)
        with open('temp_dir/xl/workbook.xml', 'w') as file_name:
            file_name.write(modified_content)
            file_name.close()
    with zipfile.ZipFile(zip_file_path, 'w') as zip_ref:
        for root, dirs, files in os.walk('temp_dir'):
            for file in files:
                file_path = os.path.join(root, file)
                zip_ref.write(file_path, os.path.relpath(file_path, 'temp_dir'))
    shutil.rmtree('temp_dir')
    os.rename(zip_file_path, f'{zip_file_path.replace(".zip", "")}')


def modify_content(content):
    if 'Protection' in content:
        first_index = content.find('<workbookProtection')
        last_index = content.find('>', first_index)
        new_content = content[:first_index] + content[last_index + 1:]
        return new_content
    else:
        print('шифрование не обнаружено')
        shutil.rmtree('temp_dir')
        os.rename(zip_file_path, f'{zip_file_path.replace(".zip", "")}')
        breakpoint()


# Пример использования функци

modify_files_in_zip(zip_file_path)
