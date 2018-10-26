class Error:
    code = -1
    msg = ""
    info = [
        '文件或文件夹路径名不正确',
        '找不到csv表',
        '另一个程序打开xlsx',
        '打开xlsx出错,请检查xlsx格式',
        '打开索引文件出错',
        '打开csv表出错',
        'old.xlsx被其他程序打开',
        'xlsx没CheckOut或者被其他程序打开',
        '读取配置文件config.txt失败',
    ]

    @staticmethod
    def set_code(code, msg):
        Error.code = code
        Error.msg = msg

    @staticmethod
    def get_code():
        return Error.code

    @staticmethod
    def get_info(code):
        return Error.info[code] + ": " + Error.msg
