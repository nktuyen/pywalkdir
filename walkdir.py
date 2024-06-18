import os
import sys
import optparse
import xlsxwriter
import enum

class IOType(enum.IntEnum):
    '''IOType class'''
    UNKNOWN = 0
    FILE = 1
    DIR = 2
    LINK = 3

class IOItem:
    '''IOItem class'''
    def __init__(self, _type: IOType, _path: str = "", _name: str = ""):
        self._type: IOType = _type
        self._path: str = _path
        self._name: str = _name
    def __str__(self):
        return self.name
    
    @property
    def type(self)->IOType:
        '''Return item type'''
        return self._type
    
    @property
    def path(self)->str:
        '''Path getter'''
        return self._path
    @path.setter
    def path(self, value: str):
        self._path = value
    
    @property
    def name(self)->str:
        '''Name getter'''
        return self._name
    @name.setter
    def name(self, value: str):
        self._name = value

    @property
    def full_path(self)->str:
        '''Return full path'''
        return os.path.join(self._path, self._name)


class IOFile(IOItem):
    '''IOFile class'''
    def __init__(self, _path: str, _name: str = ""):
        IOItem.__init__(self, IOType.FILE, _path, _name)

class IOFolder(IOItem):
    '''IOFolder class'''
    def __init__(self, _path: str, _name: str):
        IOItem.__init__(self, IOType.DIR, _path, _name)


class Command:
    '''Command class'''
    def __init__(self, _name: str = "", _desc: str = ""):
        self._name: str = _name
        self._desc: str = _desc
        self._dir: IOFolder = None
        self._options: optparse.Option = None
        self._result: int = 0

    def __str__(self)->str:
        return self._name
    
    @property
    def name(self)->str:
        '''Name getter'''
        return str(self._name)
    
    @property
    def description(self)->str:
        '''Description getter'''
        return self._desc
    
    @property
    def directory(self)->IOFolder:
        '''Directory getter'''
        return self._dir
    @directory.setter
    def directory(self, value: IOFolder):
        self._dir = value

    @property
    def options(self)->optparse.Option:
        '''Options getter'''
        return self._options
    
    @property
    def result(self)->int:
        '''Result getter'''
        return self._result

    def _pre_options_parsing(self, _args: list, _parser: optparse.OptionParser)->bool:
        '''This will be called before parsing options'''
        _parser.add_option("-v", "--verbose", action="store_false", default=None, help="Verbose output information")
        return True

    def _post_options_parsing(self):
        print(self.options)

    def _pre_start(self) -> bool:
        print(f'Command {self.name} is started')
        return True
    
    def _post_start(self):
        if self._result == 0:
            print(f'Command {self.name} is finished(success)')
        else:
            print(f'Command {self.name} is finished(failed)')
    
    def _run(self) -> int:
        return 0

    def parse_args(self, _args: list) -> bool:
        '''Parse options'''
        _parser: optparse.OptionParser = optparse.OptionParser(usage="Usage: %prog directory print [options]")
        if not self._pre_options_parsing(_args, _parser):
            return False
        try:
            opts_, args_ = _parser.parse_args(_args)
            self._options = opts_
        except Exception as ex:
            print(ex)
            return False
        self._post_options_parsing()
        return True
    
    def start(self):
        '''Execute main job'''
        self._result = -1
        if not self._pre_start():
            self._post_start()
        self._result = self._run()
        self._post_start()

class PrintCommand(Command):
    '''PrintCommand class'''
    def __init__(self, _dir: str = ""):
        Command.__init__(self, 'print', 'Print directory content')
        _path, _name = os.path.split(_dir)
        self._directory = IOFolder(_path, _name)

if __name__ == "__main__":
    args = sys.argv[1:]
    commands: dict = {}

    printCmd: PrintCommand = PrintCommand()
    commands[printCmd.name] = printCmd

    cmd_help: str = ""
    for cmd in commands.values():
        cmd_help += f'\n    {cmd.name}:    {cmd.description}'
    parser: optparse.OptionParser = optparse.OptionParser(usage="Usage: %prog directory command [options]\nCommands:"+cmd_help)
    if len(args) <= 0:
        parser.print_help()
        exit(0)

    directory: str = args[0]
    if directory == "-h" or directory == "--help":
        parser.print_help()
        exit(0)

    if not os.path.isdir(directory):
        temp: str = os.path.abspath(directory)
        if not os.path.isdir(temp):
            print(f'{directory} is not a directory')
            exit(1)
        else:
            directory = temp
    
    if len(args) <= 1:
        print('No command specified')
        parser.print_help()
        exit(1)
    

    cmd_name: str = args[1]
    if cmd_name not in commands:
        print(f'Unknown command {cmd_name}')
        print(f'Supported commands are: {",".join(commands.keys())}')
        exit(1)
    
    cmd: Command = commands[cmd_name]
    if not cmd.parse_args(args[2::]):
        exit(1)
    
    cmd.start()
    exit(cmd.result)
