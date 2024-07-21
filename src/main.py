import shutil
import traceback
import os
import zipfile
from lxml import etree


def main():
    try:
        xlsx_paths = get_file_paths(file_extension='.xlsx', directory_path=os.getcwd())
        for path in xlsx_paths:
            unzip_dir = unzip(zip_file_path=path)
            xml_dir_in_unzip_dir = unzip_dir + '\\xl\\worksheets\\'
            xml_file_paths = get_file_paths(file_extension='.xml', directory_path=xml_dir_in_unzip_dir)
            for xml_file_path in xml_file_paths:
                cut_tag(file_path=xml_file_path, tag='sheetProtection')
            zip_file_path = zip_files_in_directory(
                input_directory_path=unzip_dir,
                output_zip_file_path=unzip_dir + '_unprotected.zip'
            )
            create_xlsx(file_path=zip_file_path)
            shutil.rmtree(unzip_dir)
            os.remove(zip_file_path)
    except Exception as e:
        print(f'{e}')


def zip_files_in_directory(*, input_directory_path, output_zip_file_path) -> str:
    with zipfile.ZipFile(output_zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, _, files in os.walk(input_directory_path):
            for file in files:
                file_path = str(os.path.join(root, file))
                file = os.path.relpath(file_path, input_directory_path)
                zip_file.write(file_path, file)
    return output_zip_file_path


def cut_tag(file_path, tag):
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(file_path, parser)
    root = tree.getroot()
    namespaces = root.nsmap
    for elem in root.xpath(f'//ns:{tag}', namespaces={'ns': namespaces.get(None)}):
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
    tree.write(file_path, encoding='utf-8', xml_declaration=True, pretty_print=True)


def unzip(*, zip_file_path: str) -> str:
    dest_folder = None
    if zipfile.is_zipfile(zip_file_path):
        dest_folder = os.path.join(os.getcwd(), get_filename_without_extension(file_path=zip_file_path))
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(dest_folder)
    return dest_folder


def get_filename_without_extension(*, file_path):
    filename_with_extension = os.path.basename(file_path)
    filename_without_extension, _ = os.path.splitext(filename_with_extension)
    return filename_without_extension


def create_xlsx(*, file_path: str) -> str:
    destination = file_path.replace('.zip', '.xlsx')
    shutil.copy(file_path, destination)
    return destination


def get_file_paths(*, file_extension: str, directory_path: str) -> [str]:
    files = os.listdir(directory_path)
    paths = []
    for file in files:
        if (os.path.isfile(os.path.join(directory_path, file))
                and file_extension in get_path_and_extension(file_path=os.path.join(directory_path, file))):
            paths.append(os.path.join(directory_path, file))
    return paths


def get_path_and_extension(*, file_path) -> tuple[str, str]:
    return os.path.splitext(file_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        print(f'main: Возникло исключение: {ex}')
        tb = traceback.format_exc()
