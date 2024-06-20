import sys
import os
import re
import optparse
import concurrent.futures
import subprocess
import shutil
import xlsxwriter
from enum import Enum, IntEnum

S_Root: str = 'Root'
S_Name: str = 'Name'
S_Path: str = 'Path'
S_Fullpath: str = 'Fullpath'
S_Type: str = 'Type'
S_Size: str = 'Size'
S_Extension: str = 'Extension'
S_Command: str = 'Command'
S_Result: str = 'Result'
S_Remark: str = 'Remark'
S_Success: str = 'Success'
S_Failed: str = 'Failed'
S_Empty: str = ''
S_Sharp: str = '#'

class XlsBorderStyle(IntEnum):
    NONE = 0,
    CONTINUOUS = 1
    DASH = 3,
    DOT = 4,
    DOUBLE = 6,
    DASH_DOT = 9,
    DASH_DOT_DOT = 11,
    SLANTDASH_DOT = 13
    
class XlsFillStyle(IntEnum):
    NONE = 0,
    SOLID = 1

class XlsHAlignment(Enum):
    LEFT = 'left',
    CENTER = 'center',
    RIGHT = 'right',
    FILL = 'fill',
    JUSTIFY = 'justify',
    CENTER_ACROSS = 'center_across',
    DISTRIBUTED = 'distributed'
    
class XlsVAlignment(Enum):
    TOP = 'top',
    VCENTER = 'vcenter',
    BOTTOM = 'bottom',
    VJUSTIFY = 'vjustify',
    VDISTRIBUTED = 'vdistributed'
    
class XlsBorderFormat:
    def __init__(self) -> None:
        self._style: XlsBorderStyle = XlsBorderStyle.NONE
        self._color: str = '#000000'
        
    @property
    def style(self) -> XlsBorderStyle:
        return self._style
    
    @style.setter
    def style(self, val: XlsBorderStyle):
        self._style = val
    
    @property
    def color(self) -> str:
        return self._color
    
    @color.setter
    def color(self, val: str):
        self._color = val
        
    def copy(self, other) -> bool:
        if other is None:
            return False
        
        self.style = other.style
        self.color = other.color
        
        return True

class XlsFontFormat:
    def __init__(self) -> None:
        self._name: str = ''
        self._size: int = 0
        self._bold: bool = False
        self._strike: bool = False
        self._italic: bool = False
        self._underline: bool = False
        self._color: str = '#000000'
        
    @property
    def name(self) -> str:
        return self._name
    
    @name.setter
    def name(self, val: str):
        self._name = val
    
    @property
    def size(self) -> int:
        return self._size
    
    @size.setter
    def size(self, val: int):
        self._size = val
    
    @property
    def bold(self) -> bool:
        return self._bold
    
    @bold.setter
    def bold(self, val: bool):
        self._bold = val
    
    @property
    def italic(self) -> bool:
        return self._italic
    
    @italic.setter
    def italic(self, val: bool):
        self._italic = val
    
    @property
    def strike(self) -> bool:
        return self._strike
    
    @strike.setter
    def strike(self, val: bool):
        self._strike = val
    
    @property
    def underline(self) -> bool:
        return self._underline
    
    @underline.setter
    def underline(self, val: bool):
        self._underline = val
    
    @property
    def color(self) -> bool:
        return self._color
    
    @color.setter
    def color(self, val: str):
        self._color = val
        
    def copy(self, other) -> bool:
        if other is None:
            return False
        
        self.name = other.name
        self.size = other.size
        self.bold = other.bold
        self.italic = other.italic
        self.underline = other.underline
        self.strike = other.strike
        
        return True

class XlsFillFormat:
    def __init__(self) -> None:
        self._style: XlsFillStyle = XlsFillStyle.NONE
        self._color: str = '#000000'

    @property
    def style(self) -> XlsFillStyle:
        return self._style
    
    @style.setter
    def style(self, val: XlsFillStyle):
        self._style = val
    
    @property
    def color(self) -> str:
        return self._color
    
    @color.setter
    def color(self, val: str):
        self._color = val
    
    def copy(self, other) -> bool:
        if other is None:
            return False
        
        self.style = other.style
        self.color = other.color
        
        return True
    
class XlsBorders:
    def __init__(self) -> None:
        self._top: XlsBorderFormat = XlsBorderFormat()
        self._bottom: XlsBorderFormat = XlsBorderFormat()
        self._left: XlsBorderFormat = XlsBorderFormat()
        self._right: XlsBorderFormat = XlsBorderFormat()
        
    @property
    def top(self) -> XlsBorderFormat:
        return self._top
    
    @property
    def bottom(self) -> XlsBorderFormat:
        return self._bottom
    
    @property
    def left(self) -> XlsBorderFormat:
        return self._left
    
    @property
    def right(self) -> XlsBorderFormat:
        return self._right
    
    def copy(self, other) -> bool:
        if other is None:
            return False
        self._left.copy(other._left)
        self._right.copy(other._right)
        self._top.copy(other._top)
        self._bottom.copy(other._bottom)
        
        return True
    
class XlsCellAlignments:
    def __init__(self) -> None:
        self._h: XlsHAlignment = XlsHAlignment.LEFT
        self._v: XlsVAlignment = XlsVAlignment.TOP
        
    @property
    def horizontal(self) -> XlsHAlignment:
        return self._h
    
    @horizontal.setter
    def horizontal(self, val: XlsHAlignment):
        self._h = val
    
    @property
    def vertical(self) -> XlsVAlignment:
        return self._v
    
    @vertical.setter
    def vertical(self, val: XlsVAlignment):
        self._v = val
        
    def copy(self, other) -> bool:
        if other is None:
            return False
        self.horizontal = other.horizontal
        self.vertical = other.vertical
        
        return True
    
class XlsCellFormat:
    def __init__(self, fmt: xlsxwriter.format.Format = None) -> None:
        self._borders: XlsBorders = XlsBorders()
        self._font: XlsFontFormat = XlsFontFormat()
        self._fill: XlsFillFormat = XlsFillFormat()
        self._align: XlsCellAlignments = XlsCellAlignments()
        self._fmt = fmt
    
    @property
    def border(self) -> XlsBorders:
        return self._borders
    
    @property
    def font(self) -> XlsFontFormat:
        return self._font
    
    @property
    def fill(self) -> XlsFillFormat:
        return self._fill
    
    @property
    def align(self) -> XlsCellAlignments:
        return self._align
    
    def build(self, opt_fmt: xlsxwriter.format.Format = None) -> xlsxwriter.format.Format:
        fmt: xlsxwriter.format.Format = opt_fmt
        if fmt is None:
            fmt = self._fmt
        if fmt is None:
            return None
        fmt.set_top(self.border.top.style)
        fmt.set_top_color(self.border.top.color)
        fmt.set_bottom(self.border.bottom.style)
        fmt.set_bottom_color(self.border.bottom.color)
        fmt.set_left(self.border.left.style)
        fmt.set_left_color(self.border.left.color)
        fmt.set_right(self.border.right.style)
        fmt.set_right_color(self.border.right.color)
        
        if(self.font.name != ''):
            fmt.set_font_name(self.font.name)
        if(self.font.size > 0):
            fmt.set_font_size(self.font.size)
        if(self.font.color != ''):
            fmt.set_font_color(self.font.color)
        fmt.set_bold(self.font.bold)
        fmt.set_italic(self.font.italic)
        fmt.set_underline(self.font.underline)
        fmt.set_font_strikeout(self.font.strike)
        
        fmt.set_pattern(self.fill.style)
        if self.fill.style != XlsFillStyle.NONE:
            fmt.set_bg_color(self.fill.color)
        
        if(self.align.horizontal is not None):
            fmt.set_align(self.align.horizontal.value[0])
        if(self.align.vertical is not None):
            fmt.set_align(self.align.vertical.value[0])
        return fmt
    
    def copy(self, other) -> bool:
        if other is None:
            return False
        self.align.copy(other.align)
        self.border.copy(other.border)
        self.fill.copy(other.fill)
        self.font.copy(other.font)
        
        return True
        
    
    def clone(self):
        res: XlsCellFormat = XlsCellFormat()
        res.copy(self)
        return res    
    
class XlsHeaderFormat(XlsCellFormat):
    def __init__(self, fmt: xlsxwriter.format.Format) -> None:
        super().__init__(fmt)
        #Borders format
        self.border.left.style = XlsBorderStyle.CONTINUOUS
        self.border.left.color = '#FFFFFF'
        
        self.border.right.style = XlsBorderStyle.CONTINUOUS
        self.border.right.color = '#FFFFFF'
        
        self.border.top.style = XlsBorderStyle.CONTINUOUS
        self.border.top.color = '#FFFFFF'
        
        self.border.bottom.style = XlsBorderStyle.CONTINUOUS
        self.border.bottom.color = '#FFFFFF'
        #Fill format
        self.fill.style = XlsFillStyle.SOLID
        self.fill.color = '#002060'
        
        #Font format
        self.font.bold = True
        self.font.color = '#FFFFFF'
        
        #Alignment
        self.align.horizontal = XlsHAlignment.CENTER
        self.align.vertical = XlsVAlignment.VCENTER

class IOKind(IntEnum):
    UNKNOWN = 0,
    FILE = 1,
    DIR = 2,
    LINK = 3

class IOItem:    
    def __init__(self,kind: IOKind, name: str, path: str, depth: int = 0, parent = None) -> None:
        self._kind: IOKind = kind
        self._name: str = name
        self._path: str = path
        self._size: int = 0
        self._childs: list = []
        self._depth: int = depth
        self._parent: IOItem = parent
        self._result: bool = False
        self._tag = None
        self._status: bool = False
        
    @property
    def name(self) -> str:
        return self._name
    @name.setter
    def name(self, val: str):
        self._name = val

    @property
    def path(self) -> str:
        return self._path
    @path.setter
    def path(self, val: str):
        self._path = val

    @property
    def status(self) -> bool:
        return self._status
    @status.setter
    def status(self, val: bool):
        self._status = val

    @property
    def full_path(self) -> str:
        return os.path.join(self._path, self._name)
    
    @property
    def kind(self) -> IOKind:
        return self._kind
    
    @property
    def extension(self) -> str:
        if self._kind == IOKind.FILE:
            if len(self._name) > 0:
                ext: str = ''
                try:
                    a,b = os.path.splitext(self._name)
                    pos: int = 0
                    while b[pos] == '.' and pos < len(b):
                        pos += 1
                    ext = b[pos:]
                except:
                    ext = ''
                return ext
        return ''
        
    @property
    def parent(self):
        return self._parent
    
    @parent.setter
    def parent(self, val):
        self._parent = val
        
    @property
    def children(self) -> list:
        if self._childs is None:
            self._childs = []
        return self._childs
        
    @property
    def size(self) -> int:
        return self._size
    @size.setter
    def size(self, val: int):
        self._size = val
        
    @property
    def depth(self) -> int:
        return self._depth
    @depth.setter
    def depth(self, val: int):
        self._depth = val
    
    @property
    def tag(self):
        return self._tag
    @tag.setter
    def tag(self, val):
        self._tag = val

    def __str__(self) -> str:
        return "{{name:{}, path:{}, type:{}}}".format(self.name, self.path, self.kind)
        

class IOFile(IOItem):
    def __init__(self, name: str, path: str, depth: int = 0, parent = None) -> None:
        super().__init__(IOKind.FILE, name, path, depth, parent)
        if os.path.exists(self._path):
            stat = os.stat(self._path)
            self._size = stat.st_size

class IOFolder(IOItem):
    def __init__(self, name: str, path: str, depth: int = 0, parent = None) -> None:
        super().__init__(IOKind.DIR, name, path, depth, parent)

class IOLink(IOItem):
    def __init__(self, name: str, path: str, depth: int = 0, parent = None) -> None:
        super().__init__(IOKind.LINK, name, path, depth, parent)

class Command:
    def __init__(self, name: str, desc: str = '', dir: str = '') -> None:
        self._dir: IOFolder = None
        self._dir_count: int = 0
        self._file_count: int = 0
        self._link_count: int = 0
        if len(dir) > 0:
             _path, _name = os.path.split(dir)
             self._dir = IOFolder(_path, _name)
        self._status: bool = False
        self._name: str = name
        self._desc: str = desc
        self._prev_working_dir: str = os.curdir
        self._options: optparse.Option = None
        self._org_working_dir: str = os.path.abspath(os.curdir)
    
    @property
    def name(self) -> str:
        if self._name is None:
            self._name = ''
        return self._name
    
    @property
    def description(self) -> str:
        if self._desc is None:
            self._desc = ''
        return self._desc
    
    @property
    def directory(self) -> IOFolder:
        return self._dir
    
    @property
    def dir_count(self) -> int:
        return self._dir_count
    
    @property
    def file_count(self) -> int:
        return self._file_count
    
    @property
    def link_count(self) -> int:
        return self._link_count

    @property
    def options(self) -> optparse.Option:
        if self._options is None:
            self._options = optparse.Option()
        return self._options
    
    @property
    def status(self) -> bool:
        return self._status
 
    @property
    def original_working_dir(self) -> str:
        if self._org_working_dir is None:
            self._org_working_dir = os.curdir
        return self._org_working_dir
    
    def _printProgressBar (self, iteration, total: int=100, prefix = '', suffix = '', decimals = 1, length = 25, fill = '|', printPercent: bool = True, printEnd = "\r"):
        """
        Call in a loop to create terminal progress bar
        @params:
            iteration   - Required  : current iteration (Int)
            total       - Required  : total iterations (Int)
            prefix      - Optional  : prefix string (Str)
            suffix      - Optional  : suffix string (Str)
            decimals    - Optional  : positive number of decimals in percent complete (Int)
            length      - Optional  : character length of bar (Int)
            fill        - Optional  : bar fill character (Str)
            printPercent- Optional  : Whether or not print percentage value
            printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
        """
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        if printPercent:
            print(f'\r{prefix}|{bar}| {percent}%', end = printEnd)
        else:
            print(f'\r{prefix}|{bar}| {suffix}', end = printEnd)
        # Print New Line on Complete
        if iteration == total: 
            print()

    def _shorten_path(self, path: str, max_len: int = 0) -> str:
        if not isinstance(path, str):
            return path
        
        if max_len <= 0:
            return path
        
        if len(path) <= max_len:
            return path
        
        dirs: list = path.split(os.sep)
        if(len(dirs) <= 2):
            return path
        
        count: int = len(dirs)-2
        new_path: str = path
        base_name: str = dirs[len(dirs)-1]
        shorten: str = ''
        root_dir: str = dirs[0]
        prefix: str = ''
        if not root_dir.endswith(os.sep):
            root_dir += os.sep
        while count > 1 and len(new_path) > max_len:
            prefix = ''
            shorten = ''
            for i in range(1, count):
                prefix += f'{dirs[i]}{os.sep}'
            for j in range(count, len(dirs)-1):
                shorten += f'...{os.sep}'
            if prefix.endswith(os.sep):
                prefix = prefix.removesuffix(os.sep)
            new_path = os.path.join(root_dir, prefix)
            new_path += f'..{os.sep}{base_name}'
            count -= 1
        return new_path
    
    def _print_stderr(self, p: subprocess.Popen):
        if p is not None:
            try:
                stdout,stderr = p.communicate( 10)
                if stderr is not None:
                    b: str = stderr.decode('utf-8')
                    if b:
                        print(b)
            except:
                pass
    
    def _change_working_dir(self, dir: str) -> str:
        self._prev_working_dir = os.curdir
        return os.chdir(dir)
    
    def _restore_working_dir(self) -> str:
        return os.chdir(self._prev_working_dir)

    def _onAddOptions(self, parser:optparse.OptionParser):
        parser.add_option('-v', '--verbose', action="store_false", help='Verbose output logs')
        parser.add_option('-o', '--output', help='Output file to store result. Default is result.xlsx')
        parser.add_option('-x', '--exclude', help='Exclude patterns. Comma separated')
        parser.add_option('-r', '--recursive', action="store_false", help='Walk recursively')

    def _onOptionsParsed(self):
        pass

    def _is_matched(self, name: str, pattern: str) -> bool:
        if pattern == '*':
            return True
        rex = re.compile('^' + pattern.replace('*',  '.*').replace('?', '.') + '$')
        if rex.match(name):
            return True
        return False
    
    def _walk(self, root : IOFolder) -> IOFolder:
        if self.options.verbose is not None:
            print(f'Walking on {root.full_path}')
        else:
            line: str = f'Walking on {self._shorten_path(root.full_path, 128)}'
            print(f'\r{line}{" "*(150-len(line))}', end='\r')
        subpaths = os.listdir(root.full_path)
        abs_path: str = ''
        ignored: bool = False
        for path in subpaths:
            if root.depth == 0:
                root.depth = 1
            ignored = False
            if isinstance(self.options.exclude, str):
                excludes: list = self.options.exclude.split(',')
                for x in excludes:
                    if self._is_matched(path, x.strip()):
                        ignored = True
                        break
            if ignored:
                if self.options.verbose is not None:
                    print(f'Ignoring {path}...')
                continue
            abs_path = os.path.join(root.full_path, path)
            current: IOItem = None
            if os.path.isdir(abs_path):
                current = IOFolder(path, root.full_path, 0, root)
                current.parent = root
                current.tag = self
                self._dir_count += 1
                if current.parent.depth <= current.depth:
                    current.parent.depth = current.depth + 1
                else:
                    current.depth = current.parent.depth - 1
                if self.options.recursive is not None:
                    current = self._walk(current)
            elif os.path.islink(abs_path):
                current = IOLink(path, root.full_path, 0, root)
                current.tag = self
                self._link_count += 1
            else:
                current = IOFile(path, root.full_path, 0, root)
                current.tag = self
                self._file_count += 1
            current.parent = root
            if current.parent.depth <= current.depth:
                current.parent.depth = current.depth + 1
            else:
                current.depth = current.parent.depth - 1
            current.parent.size += current.size
            root.children.append(current)
        return root

    def parse_args(self, options) -> bool:
        parser: optparse.OptionParser = optparse.OptionParser(f'%prog {self._name} [options]')
        self._onAddOptions(parser)

        try:
            opts, args = parser.parse_args(options)
            self._options = opts
        except Exception as ex:
            print(ex)
            parser.print_help()
            return False
        self._onOptionsParsed()
        return True   
    
    def execute(self, dir: str = None) -> bool:
        if isinstance(dir, str):
            _path, _name = os.path.split(dir)
            self._dir = IOFolder(_name, _path)
        if not self._preExecute():
            self._postExecute()
            return False
        self._status = self._onExecute()
        self._postExecute()
        return self._status
    
    def _preExecute(self) -> bool:
        print(f'\nCommand {self._name} is started')
        print('='*50)
        if self.directory is None:
            print(f'No directory specified')
            return False
        return True

    def _onExecute(self) -> bool:
        try:
            self._dir_count = 0
            self._file_count = 0
            self._link_count = 0
            self._dir = self._walk(self.directory)
        except Exception as ex:
            print(ex)
            return False
        return True
    
    def _postExecute(self):
        print()
        print('='*50)
        if self._status:
            print(f'Command {self._name} is finished (success)')
        else:
            print(f'Command {self._name} is finished (fail)')

class PrintCommand(Command):
    def __init__(self, dir: str = '') -> None:
        super().__init__('print', 'Print directory content', dir)
        self._fields_size: dict = {}
        self._fields_size[S_Sharp] = 0
        self._fields_size[S_Name] = 0
        self._fields_size[S_Path] = 0
        self._fields_size[S_Fullpath] = 0
        self._fields_size[S_Type] = 0
        self._fields_size[S_Size] = 0
        self._fields_size[S_Extension] = 0
        self._fields_size[S_Command] = 0
        self._fields_size[S_Result] = 0
        self._fields_size[S_Remark] = 0
    
    def _onAddOptions(self, parser: optparse.OptionParser):
        super()._onAddOptions(parser)
        parser.add_option('-n', '--print-name', action='store_false', help='Print item name')
        parser.add_option('-p', '--print-path', action='store_false', help='Print item path')
        parser.add_option('-t', '--print-type', action='store_false', help='Print item type')
        parser.add_option('-s', '--print-size', action='store_false', help='Print item size')
        parser.add_option('-e', '--print-ext', action='store_false', help='Print item extension')
        parser.add_option('-c', '--print-command', action='store_false', help='Print command name')
        parser.add_option('-u', '--print-result', action='store_false', help='Print command result')
        parser.add_option('-a', '--print-parent', action='store_false', help='Print parent item name before the item\'s name')
        parser.add_option('-m', '--print-remark', action='store_false', help='Print remark information')

    def _onExecute(self) -> bool:
        if not super()._onExecute():
            return False

        fields: list =[]
        if self.options.print_name is not None:
            fields.append(S_Name)
        if self.options.print_path is not None:
            fields.append(S_Path)
            fields.append(S_Fullpath)
        if self.options.print_type is not None:
            fields.append(S_Type)
        if self.options.print_size is not None:
            fields.append(S_Size)
        if self.options.print_ext is not None:
            fields.append(S_Extension)
        if self.options.print_command is not None:
            fields.append(S_Command)
        if self.options.print_result is not None:
            fields.append(S_Result)
        if self.options.print_remark is not None:
            fields.append(S_Remark)
        
        try:
            if self.options.output is None:
                #Print to console
                print()
                self._printDirectory(self.directory, self.directory.depth, ' ', fields, False)
                self._printDirectory(self.directory, self.directory.depth, ' ', fields, True)
            else:
                #Write to output file
                row: int = 0
                col: int = 0
                
                _wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.options.output)
                _ws: xlsxwriter.worksheet.Worksheet = _wb.add_worksheet(self.directory.name)
                
                hdr_fmt = XlsHeaderFormat(_wb.add_format())
                file_fmt = XlsCellFormat(_wb.add_format())
                file_fmt.border.top.style = XlsBorderStyle.CONTINUOUS
                file_fmt.border.left.style = XlsBorderStyle.CONTINUOUS
                dir_fmt = XlsCellFormat(_wb.add_format())
                dir_fmt.font.bold = True
                dir_fmt.border.top.style = XlsBorderStyle.CONTINUOUS
                dir_fmt.border.left.style = XlsBorderStyle.CONTINUOUS
                dir_fmt.fill.style = XlsFillStyle.SOLID
                dir_fmt.fill.color = '#EEEEEE'
                fmt = hdr_fmt.build()
                #Write header cells
                _ws.write(row, col, S_Root, fmt)
                for j in range(1, self.directory.depth + 1):
                    _ws.write(row, col + j, 'Sub-item level {}'.format(j), fmt)
                additional_col: int = self.directory.depth
                if S_Name in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Name, fmt)
                if S_Path in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Path, fmt)
                if S_Fullpath in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Fullpath, fmt)
                if S_Type in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col,S_Type, fmt)
                if S_Size in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Size, fmt)
                if S_Extension in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Extension, fmt)
                if S_Command in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Command, fmt)
                if S_Result in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Result, fmt)
                if S_Remark in fields:
                    additional_col += 1
                    _ws.write(row, col + additional_col, S_Remark, fmt)

                row += 1
                _ws.write(row, col, S_Name, fmt)
                for j in range(1, self.directory.depth + 1):
                    _ws.write(row, col + j, S_Name, fmt)
                additional_col: int = self.directory.depth
                if S_Name in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Name, fmt)
                if S_Path in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Path, fmt)
                if S_Fullpath in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Fullpath, fmt)
                if S_Type in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Type, fmt)
                if S_Size in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Size, fmt)
                if S_Extension in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Extension, fmt)
                
                if S_Command in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Command, fmt)
                
                if S_Result in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Result, fmt)

                if S_Remark in fields:
                    additional_col += 1
                    _ws.merge_range(row - 1, col + additional_col, row, col + additional_col, S_Remark, fmt)
                
                logparent: bool = False
                if self.options.print_parent is not None:
                    logparent = True
                progress: int = 0
                last_row = self._writeOutput(self.directory, _wb, _ws, row + 1, col, self.directory.depth, file_fmt, dir_fmt, fields, logparent, progress) + 1
                footer_fmt = XlsCellFormat(_wb.add_format())
                footer_fmt.border.top.style = XlsBorderStyle.CONTINUOUS
                additional_col += 1
                for c in range(col, col + additional_col):
                    _ws.write_blank(last_row, c, None, footer_fmt.build())
                footer_fmt.border.top.style = XlsBorderStyle.NONE
                footer_fmt.border.left.style = XlsBorderStyle.CONTINUOUS
                for r in range(row, last_row):
                    _ws.write_blank(r, col + additional_col, None, footer_fmt.build(_wb.add_format()))

                _wb.close()
        except Exception as ex:
            print(ex)
            return False
        return True
    
    def _writeOutput(self, item: IOItem, _wb: xlsxwriter.Workbook, _ws: xlsxwriter.worksheet.Worksheet, row: int, col: int, root_depth: int, fmt_file: XlsCellFormat = None, fmt_dir: XlsCellFormat = None, fields: list = [], logparent: bool = False, progress: int = None) -> int:
        if self.options.verbose is not None:
            print(f'Printing {item.full_path}')
        else:
            line: str = f'Printing {self._shorten_path(item.full_path, 128)}'
            print(f'\r{line}{" "*(140-len(line))}', end='\r')
        fmt: XlsCellFormat = None
        if item.kind == IOKind.DIR:
            if fmt_dir is not None:
                fmt = fmt_dir
        else:
            fmt = fmt_file
        _ws.write(row, col + root_depth - item.depth, item.name, fmt.build())
        additional_col: int = root_depth
        if S_Name in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.name, fmt.build())
        if S_Path in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.path, fmt.build())
        if S_Fullpath in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.full_path, fmt.build())
        if S_Type in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.kind.name, fmt.build())
        if S_Size in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.size, fmt.build())
        if S_Extension in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, item.extension, fmt.build())
        
        if S_Command in fields:
            additional_col += 1
            if item.tag is not None:
                _ws.write(row, col + additional_col, item.tag.name, fmt.build())
            else:
                _ws.write(row, col + additional_col, S_Empty, fmt.build())
        
        item.status = True
        if isinstance(progress, int):
            progress += 1
        if S_Result in fields:
            additional_col += 1
            if item.tag is not None:
                if item.status:
                    _ws.write(row, col + additional_col, S_Success, fmt.build())
                else:
                    _ws.write(row, col + additional_col, S_Failed, fmt.build())
            else:
                _ws.write(row, col + additional_col, S_Empty, fmt.build())
        
        if S_Remark in fields:
            additional_col += 1
            _ws.write(row, col + additional_col, S_Remark, fmt.build())
        
        first_row: int = row
        if item.depth > 0:
            new_fmt = fmt.clone()
            new_fmt.border.left.style = XlsBorderStyle.NONE
            new_fmt.border.right.style = XlsBorderStyle.NONE
            for c in range(1, item.depth + 1):
                _ws.write_blank(row, col + root_depth - item.depth + c,  None, new_fmt.build(_wb.add_format()))

        for child in item.children:
            row = self._writeOutput(child, _wb, _ws, row + 1, col, root_depth, fmt_file, fmt_dir, fields, logparent, progress)
        if first_row < row:
            new_fmt = fmt.clone()
            new_fmt.border.left.style = XlsBorderStyle.CONTINUOUS
            new_fmt.border.right.style = XlsBorderStyle.NONE
            new_fmt.border.top.style = XlsBorderStyle.NONE
            new_fmt.border.bottom.style = XlsBorderStyle.NONE
            for r in range(first_row + 1, row + 1):
                if logparent:
                    _ws.write(r, col + root_depth - item.depth, item.name, new_fmt.build(_wb.add_format()))
                else:
                    _ws.write_blank(r, col + root_depth - item.depth, None, new_fmt.build(_wb.add_format()))
        return row 
    
    def _printDirectory(self, folder: IOFolder, depth: int = 0, separator: str = ' ', fields: list = None, do_print: bool = True):
        if not do_print: #Only evaluate
            line = f'{(depth-folder.depth) * separator}{folder.name}'
            if self._fields_size.get(S_Sharp) < len(line):
                self._fields_size[S_Sharp] = len(line)
            
            if S_Name in fields:
                if S_Name in self._fields_size:
                    if self._fields_size.get(S_Name) < len(folder.name):
                        self._fields_size[S_Name] = len(folder.name)
                else:
                    self._fields_size[S_Name] = len(folder.name)
            
            if S_Path in fields:
                if S_Path in self._fields_size:
                    if self._fields_size.get(S_Path) < len(folder.path):
                        self._fields_size[S_Path] = len(folder.path)
                else:
                    self._fields_size[S_Path] = len(folder.path)
            
            if S_Fullpath in fields:
                if S_Path in self._fields_size:
                    if self._fields_size.get(S_Fullpath) < len(folder.full_path):
                        self._fields_size[S_Fullpath] = len(folder.full_path)
                else:
                    self._fields_size[S_Fullpath] = len(folder.full_path)
            
            if S_Type in fields:
                if S_Type in self._fields_size:
                    if self._fields_size.get(S_Type) < len(folder.kind.name):
                        self._fields_size[S_Type] = len(folder.kind.name)
                else:
                    self._fields_size[S_Type] = len(folder.kind.name)

            if S_Size in fields:
                if S_Size in self._fields_size:
                    if self._fields_size.get(S_Size) < len(str(folder.size)):
                        self._fields_size[S_Size] = len(str(folder.size))
                else:
                    self._fields_size[S_Size] = len(str(folder.size))

            if S_Extension in fields:
                if S_Extension in self._fields_size:
                    if self._fields_size.get(S_Extension) < len(folder.kind.name):
                        self._fields_size[S_Extension] = len(folder.kind.name)
                else:
                    self._fields_size[S_Extension] = len(folder.kind.name)
        else: #Print
            line = f'{(depth-folder.depth) * separator}{folder.name}'
            if len(fields) > 0:#There are more fields to print -> Add space to right
                line += separator * (self._fields_size[S_Sharp] - len(line) + len(separator))
            if S_Name in fields:
                line += folder.name
                line += separator * (self._fields_size[S_Name] - len(folder.name) + len(separator))
            if S_Path in fields:
                line += folder.path
                line += separator * (self._fields_size[S_Path] - len(folder.path) + len(separator))
            if S_Fullpath in fields:
                line += folder.full_path
                line += separator * (self._fields_size[S_Fullpath] - len(folder.full_path) + len(separator))
            if S_Type in fields:
                line += folder.kind.name
                line += separator * (self._fields_size[S_Type] - len(folder.kind.name) + len(separator))
            if S_Size in fields:
                line += str(folder.size)
                line += separator * (self._fields_size[S_Size] - len(str(folder.size)) + len(separator))
            if S_Extension in fields:
                line += folder.extension
            print(line)

        folder.status = True

        for child in folder.children:
            if child.kind == IOKind.DIR:
                self._printDirectory(child, depth, separator, fields, do_print)
            else:
                if not do_print: #Only evaluate
                    line = f'{(depth-child.depth) * separator}{child.name}'
                    if self._fields_size.get(S_Sharp) < len(line):
                        self._fields_size[S_Sharp] = len(line)
                    
                    if S_Name in fields:
                        if S_Name in self._fields_size:
                            if self._fields_size.get(S_Name) < len(child.name):
                                self._fields_size[S_Name] = len(child.name)
                        else:
                            self._fields_size[S_Name] = len(child.name)
                    
                    if S_Path in fields:
                        if S_Path in self._fields_size:
                            if self._fields_size.get(S_Path) < len(child.path):
                                self._fields_size[S_Path] = len(child.path)
                        else:
                            self._fields_size[S_Path] = len(child.path)
                    
                    if S_Fullpath in fields:
                        if S_Path in self._fields_size:
                            if self._fields_size.get(S_Fullpath) < len(child.full_path):
                                self._fields_size[S_Fullpath] = len(child.full_path)
                        else:
                            self._fields_size[S_Fullpath] = len(child.full_path)
                    
                    if S_Type in fields:
                        if S_Type in self._fields_size:
                            if self._fields_size.get(S_Type) < len(child.kind.name):
                                self._fields_size[S_Type] = len(child.kind.name)
                        else:
                            self._fields_size[S_Type] = len(child.kind.name)

                    if S_Size in fields:
                        if S_Size in self._fields_size:
                            if self._fields_size.get(S_Size) < len(str(child.size)):
                                self._fields_size[S_Size] = len(str(child.size))
                        else:
                            self._fields_size[S_Size] = len(str(child.size))

                    if S_Extension in fields:
                        if S_Extension in self._fields_size:
                            if self._fields_size.get(S_Extension) < len(child.kind.name):
                                self._fields_size[S_Extension] = len(child.kind.name)
                        else:
                            self._fields_size[S_Extension] = len(child.kind.name)
                else: #Print
                    line = f'{(depth-child.depth) * separator}{child.name}'
                    if len(fields) > 0:#There are more fields to print -> Add space to right
                        line += separator * (self._fields_size[S_Sharp] - len(line) + len(separator))
                    if S_Name in fields:
                        line += child.name
                        line += separator * (self._fields_size[S_Name] - len(child.name) + len(separator))
                    if S_Path in fields:
                        line += child.path
                        line += separator * (self._fields_size[S_Path] - len(child.path) + len(separator))
                    if S_Fullpath in fields:
                        line += child.full_path
                        line += separator * (self._fields_size[S_Fullpath] - len(child.full_path) + len(separator))
                    if S_Type in fields:
                        line += child.kind.name
                        line += separator * (self._fields_size[S_Type] - len(child.kind.name) + len(separator))
                    if S_Size in fields:
                        line += str(child.size)
                        line += separator * (self._fields_size[S_Size] - len(str(child.size)) + len(separator))
                    if S_Extension in fields:
                        line += child.extension
                    print(line)
                
                child.status = True
        

if __name__=="__main__":
    #Initialize supported commands
    commands: dict = {}

    initCmd: PrintCommand = PrintCommand()
    commands[initCmd.name] = initCmd

    help: str = ''
    for c in commands.values():
        help += f'\n  {c.name}:    {c.description}'
    parser: optparse.OptionParser = optparse.OptionParser(usage='%prog directory command [options]\n\nCommands:    '+ help)
    errno: int = 1

    args = sys.argv[1::]
    if len(args) <= 0:
        parser.print_help()
        exit(0)
    
    dir = args[0]
    if dir == '-h' or dir == '--help':
        parser.print_help()
        exit(0)

    temp: str = os.path.abspath(dir)
    if not os.path.isdir(temp):
        print(f'{dir} is not a directory')
        exit(1)
    else:
        dir = temp

    if len(args) < 2:
        print('Error: No command specified')
        parser.print_help()
        exit(0)
    cmd =  args[1]
    command: Command = None
    if cmd in commands:
        command = commands[cmd]
    
    if command is None:
        print(f'Unknown command "{cmd}"')
        errno+=1
        exit(errno)
    
    options = args[2::]
    if not command.parse_args(options):
        errno+=1
        exit(errno)
    
    if command.execute(dir):
        exit(0)
    else:
        exit(-1)