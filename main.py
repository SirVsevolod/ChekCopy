import os
import win32com.client
import win32api
import xlsxwriter


def GetExceptions():
    exception = []
    Path = input('Введите путь к файлу исключений: ')
    file = open(Path, 'r')
    for line in  file:
        exception.append(line[0:-1])
    file.close()
    return exception


def get_file_metadata(path, filename, metadata):
    sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
    #print(dir(sh))
    ns = sh.NameSpace(path)
    #print(dir(ns))
    #signer = ns.Signer
    #print(signer.Certificate.IssuerName, signer.Certificate.SerialNumber)
    # Enumeration is necessary because ns.GetDetailsOf only accepts an integer as 2nd argument
    file_metadata = dict()
    item = ns.ParseName(str(filename))
    for ind, attribute in enumerate(metadata):
        attr_value = ns.GetDetailsOf(item, ind)
        if attr_value:
            file_metadata[attribute] = attr_value
        else:
            file_metadata[attribute] = None
    return file_metadata


#1) Проверка существования цифровой подписи у EXE, DLL, DRX + исключения
def DigitalSignature(path, metadata, exceptions=[]):
    print('Проверка существования цифровой подписи у EXE, DLL, DRX')
    result = []
    tree = os.walk(path)
    for i in tree:
        for file in i[2]:
            file_name, file_extension = os.path.splitext(i[0] + file)
            if file_extension in ('.exe', '.dll', '.drx') and file not in exceptions:
                meta_data = get_file_metadata(i[0], file, metadata)
                if meta_data['Company'] == None:
                    result.append({'path': i[0] + '\\' + file,
                                   'copyright': meta_data['Copyright'],
                                   'company': meta_data['Company']})
    return result


#2) Проверка владельца цифровой подписи у EXE, DLL, DRX + input владелец цифровой подписи + исключения
def CheckOwner(path, metadata, owner, exceptions=[],):
    print('Проверка владельца цифровой подписи у EXE, DLL, DRX')
    result = []
    tree = os.walk(path)
    for i in tree:
        for file in i[2]:
            file_name, file_extension = os.path.splitext(i[0] + file)
            if file_extension in ('.exe', '.dll', '.drx') and file not in exceptions:
                meta_data = get_file_metadata(i[0], file, metadata)
                if meta_data['Company'] != owner:

                    result.append({'path': i[0] + '\\' + file,
                                   'copyright': meta_data['Copyright'],
                                   'company': meta_data['Company']})
    return result


#3) Проверить авторские права в файлах директории + input (c)
def CheckCopyright(path, metadata, copyright,):
    print('Проверить авторские права в файлах директории')
    result = []
    tree = os.walk(path)
    for i in tree:
        for file in i[2]:
            file_name, file_extension = os.path.splitext(i[0] + file)
            if file_extension in ('.exe', '.dll', '.drx'):
                meta_data = get_file_metadata(i[0], file, metadata)
                if meta_data['Copyright'] != copyright:
                    result.append({'path': i[0] + '\\' + file,
                                   'copyright': meta_data['Copyright'],
                                   'company': meta_data['Owner']})
    return result


#4) Проверить номера сборки в файлах директории + input номер актуальной сборки
def CheckAssembly(path, version,):
    result = []
    tree = os.walk(path)
    for i in tree:
        for file in i[2]:
            file_name, file_extension = os.path.splitext(i[0] + file)
            if file_extension in ('.exe', '.dll', '.drx'):
                print(i[0] + '\\' + file)
                fixedInfo = win32api.GetFileVersionInfo(i[0] + '\\' + file, '\\')
                print(fixedInfo)
                result.append({"Path": i[0] + file,
                               })

    return result

#5)Запись результатов
def WriteResult(DS, CO, CC):
    print('Запись результатов')
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet(name="ЦП")
    i = 1
    for file in DS:
        worksheet.write('A' + str(i), file['path'])
        worksheet.write('B' + str(i), file['copyright'])
        worksheet.write('C' + str(i), file['company'])
        i += 1
    i = 1

    worksheet1 = workbook.add_worksheet(name="ВЦП")
    for file in CO:
        worksheet1.write('A' + str(i), file['path'])
        worksheet1.write('B' + str(i), file['copyright'])
        worksheet1.write('C' + str(i), file['company'])
        i += 1
    i = 1

    worksheet2 = workbook.add_worksheet(name="АП")
    for file in CC:
        worksheet2.write('A' + str(i), file['path'])
        worksheet2.write('B' + str(i), file['copyright'])
        worksheet2.write('C' + str(i), file['company'])
        i += 1

    workbook.close()


if __name__ == "__main__":
    #inputs
    exceptions = GetExceptions()
    print(exceptions)
    metadata = ['Name', 'Size', 'Item type', 'Date modified', 'Date created', 'Date accessed', 'Attributes', 'Offline status', 'Availability', 'Perceived type', 'Owner', 'Kind', 'Date taken', 'Contributing artists', 'Album', 'Year', 'Genre', 'Conductors', 'Tags', 'Rating', 'Authors', 'Title', 'Subject', 'Categories', 'Comments', 'Copyright', '#', 'Length', 'Bit rate', 'Protected', 'Camera model', 'Dimensions', 'Camera maker', 'Company', 'File description', 'Masters keywords', 'Masters keywords']
    Path = input("Введите путь: ")
    owner = input("Владелец цифровой подписи: ")
    copyright = input("Введите копирайт: ")
    #assembly = input("Введите номер сборки: ")

    DS = DigitalSignature(Path, metadata)
    CO = CheckOwner(Path,metadata, owner)
    CC = CheckCopyright(Path, metadata, copyright)
    #CA = CheckAssembly(Path, assembly)

    WriteResult(DS, CO, CC)
