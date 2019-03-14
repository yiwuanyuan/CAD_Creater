import shutil, os,re

def copy_recursion(path, target_path):
    if not os.path.exists(path):
        print(path + '不存在')
        return
    if not os.path.exists(target_path):
        temp_array = target_path.split('\\')
        temp_addr = ''
        for i in range(len(temp_array)):
            temp_addr += temp_array[i]
            if not os.path.exists(temp_addr):
                os.mkdir(temp_addr)
            temp_addr += '\\'

    for i in os.listdir(path):
        path_file = os.path.join(path, i)

        if os.path.isfile(path_file):
            shutil.copy(path_file, target_path)
        else:
            if path_file.find('厚度') != -1:
                temp_target_path = target_path+'\\'+path_file.split('\\')[-1]
                # print(path)
                # print(temp_target_path)
                copy_recursion(path_file, temp_target_path)
            else:
                copy_recursion(path_file, target_path)

def del_recursion(path):
    if not os.path.exists(path):
        print(path + '不存在')
        return
    for i in os.listdir(path):
        path_file = os.path.join(path, i)
        # print(path_file)
        if os.path.isfile(path_file):
            os.remove(path_file)
        else:
            del_recursion(path_file)
    os.rmdir(path)
    return

class FileOperator():
    def __init__(self,path):
        self.path = re.sub('/', '\\\\', path)
        self.target_path = re.sub('dxf', '汇总', path)

    @staticmethod
    def copy_method(path, target_path):
        try:
            FileOperator.del_method(target_path)
            copy_recursion(path, target_path)
        except Exception as e:
            print(e)
            raise Exception(e)

    @staticmethod
    def del_method(path):
        try:
            del_recursion(path)
        except Exception as e:
            print(e)
            raise Exception(e)



# path = 'C:/Users/JAMES/Desktop/V3/test/dxf'
# path = FileOperator(path).path
# target_path = FileOperator(path).target_path
# FileOperator.copy_method(path,target_path)