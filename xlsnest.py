from asyncio.windows_events import NULL
from pickle import TRUE
import xlrd
import xlutils.copy
import xlwt

import sys
import re
import os

from Common.log import logger


class EsMergedException(Exception):

    def __init__(self, msg=None):
        self.msg = msg


class Escell:
    ''' 关注的excel字段数据

    Attributes:
        pos:    cell的坐标, 例如: (x1,x2,y1,y2), 表示坐标(x1_x2,y1_y2)的(合并)单元格
        name:   关注字段名
        type:   模板类型,ver表示向下读取一个list,hor表示只提取当前位置的值
        max:    type=ver场景下,最多可以向下读取的元素长度
    '''

    def __init__(self, pos, name, type, max=-1):
        # 坐标
        self.x1 = pos[0]
        self.x2 = pos[1]
        self.y1 = pos[2]
        self.y2 = pos[3]
        # 标记名
        self.name = name
        # 类型：ver -> []类型, hor -> {}类型
        self.type = type
        # 类型为ver时,最大读取数量,默认值-1表示没有限制
        self.max = max

    def cell_width(self):
        return self.x2 - self.x1

    def cell_heigh(self):
        return self.y2 - self.y1


class Esmod:

    def __init__(self, book):
        self.__source__ = book

        self.style_list = []  # 格式列表
        logger.info('读取模板格式...')
        for rdxf in book.xf_list:
            wtxf = xlwt.Style.XFStyle()
            #
            # number format
            #
            wtxf.num_format_str = book.format_map[rdxf.format_key].format_str
            #
            # font
            #
            wtf = wtxf.font
            rdf = book.font_list[rdxf.font_index]
            wtf.height = rdf.height
            wtf.italic = rdf.italic
            wtf.struck_out = rdf.struck_out
            wtf.outline = rdf.outline
            wtf.shadow = rdf.outline
            wtf.colour_index = rdf.colour_index
            wtf.bold = rdf.bold  # This attribute is redundant, should be driven by weight
            wtf._weight = rdf.weight  # Why "private"?
            wtf.escapement = rdf.escapement
            wtf.underline = rdf.underline_type
            # wtf.???? = rdf.underline #### redundant attribute, set on the fly when writing
            wtf.family = rdf.family
            wtf.charset = rdf.character_set
            wtf.name = rdf.name
            #
            # protection
            #
            wtp = wtxf.protection
            rdp = rdxf.protection
            wtp.cell_locked = rdp.cell_locked
            wtp.formula_hidden = rdp.formula_hidden
            #
            # border(s) (rename ????)
            #
            wtb = wtxf.borders
            rdb = rdxf.border
            wtb.left = rdb.left_line_style
            wtb.right = rdb.right_line_style
            wtb.top = rdb.top_line_style
            wtb.bottom = rdb.bottom_line_style
            wtb.diag = rdb.diag_line_style
            wtb.left_colour = rdb.left_colour_index
            wtb.right_colour = rdb.right_colour_index
            wtb.top_colour = rdb.top_colour_index
            wtb.bottom_colour = rdb.bottom_colour_index
            wtb.diag_colour = rdb.diag_colour_index
            wtb.need_diag1 = rdb.diag_down
            wtb.need_diag2 = rdb.diag_up
            #
            # background / pattern (rename???)
            #
            wtpat = wtxf.pattern
            rdbg = rdxf.background
            wtpat.pattern = rdbg.fill_pattern
            wtpat.pattern_fore_colour = rdbg.pattern_colour_index
            wtpat.pattern_back_colour = rdbg.background_colour_index
            #
            # alignment
            #
            wta = wtxf.alignment
            rda = rdxf.alignment
            wta.horz = rda.hor_align
            wta.vert = rda.vert_align
            wta.dire = rda.text_direction
            # wta.orie # orientation doesn't occur in BIFF8! Superceded by rotation ("rota").
            wta.rota = rda.rotation
            wta.wrap = rda.text_wrapped
            wta.shri = rda.shrink_to_fit
            wta.inde = rda.indent_level
            # wta.merg = ????
            #
            self.style_list.append(wtxf)

        def read_xls_defines(book):
            ''' 加载excel模板

            Args:
                book:  excel(xls)文件
            Returns:
                返回一个Esmod
            '''
            # 获取所有sheet
            sheets = book.sheets()

            # 定义一个字典
            defines = {}

            def read_sheet(sheet):
                ''' 读取excel表单

                读取使用xlrd加载的excel表单
                Args:
                    sheet:  excel表(xlrd)
                Returns:
                    返回一个字典,格式如下:

                    {
                        '字段名': Escell(坐标,字段名,类型,列表类型下最大读取数量)
                        ...
                    }
                '''
                grp = {}
                pos_dict = {(x1, y1): (x1, x2, y1, y2)
                            for (x1, x2, y1, y2) in sheet.merged_cells}

                def get_pos(x, y):
                    if (x, y) in pos_dict:
                        return (x, pos_dict[(x, y)][1], y, pos_dict[(x, y)][3])
                    return (x, x+1, y, y+1)

                def check(dict, sheet, name, x, y):
                    if name in dict:
                        logger.warning('字段重复[{}],将使用\'{}\'({})'.format(
                            name, sheet.name, xlrd.cellname(x, y)))

                for x in range(sheet.nrows):
                    row = sheet.row_values(x)
                    for y in range(sheet.ncols):
                        cell = str(row[y])
                        key = cell[1:-1]

                        escell = None
                        if cell.startswith('{') and cell.endswith('}'):
                            check(grp, sheet, key, x, y)
                            escell = Escell(get_pos(x, y), key, 'hor')
                            grp[key] = escell
                        elif cell.startswith('[') and cell.endswith(']'):
                            try:
                                idx = key.index(':')
                            except ValueError:
                                idx = -1
                            max = -1
                            cell_type = 'ver'
                            if idx > 0:
                                obj = key[idx+1:]
                                if obj == '+':
                                    max = -1
                                    cell_type = 'ver-auto-add'
                                else:
                                    max = int(key[idx+1:])
                                key = key[0:idx]

                            check(grp, sheet, key, x, y)
                            escell = Escell(get_pos(x, y), key, cell_type, max)
                            grp[key] = escell
                        if escell:
                            logger.info("读取字段'{}'({}),类型'{}'".format(
                                escell.name, xlrd.cellname(escell.x1, escell.y1), escell.type))

                return grp

            # 遍历sheet
            for i, sheet in enumerate(sheets):
                # 读取sheet，并以(sheet名,sheet_id)作为key存到字典中
                defines[(sheet.name, i)] = read_sheet(sheet)

            # 返回Esmod
            return defines

        logger.info('读取字段定义...')
        self.__defines__ = read_xls_defines(book)

        logger.info('模板初始化完毕')

    def get_defines(self):
        return self.__defines__.copy()

    def load(self, path, inindex=None):
        logger.info('读取数据文档: {}...'.format(path))

        book = xlrd.open_workbook(path)

        def load_sheet(sheet, cell_dict):
            logger.info('读取表单({})数据...'.format(sheet.name))
            data = {}
            sheet_max_row = sheet.nrows

            def read_value(sheet, x, y):
                try:
                    return sheet.cell_value(x, y)
                except IndexError:
                    logger.error("读取 ({}) 时出错,超出单元格范围！！！".format(
                        xlrd.cellname(x, y)))
                    raise EsMergedException("读取 ({}) 时出错,超出单元格范围！！！".format(
                        xlrd.cellname(x, y)))

            for field in cell_dict:
                escell = cell_dict[field]
                if escell.type == 'hor':
                    data[field] = read_value(sheet, escell.x1, escell.y1)
                elif escell.type == 'ver':
                    count = int((sheet_max_row - escell.x1) /
                                escell.cell_heigh())
                    if not escell.max == -1 and count > escell.max:
                        count = escell.max

                    field_value = []
                    for i in range(count):
                        field_value.append(read_value(
                            sheet, escell.x1 + escell.cell_heigh() * i, escell.y1))

                    data[field] = field_value

                if escell.type == 'ver-auto-add':
                    logger.info('自增型字段：{}'.format(field))
                else:
                    logger.info('读取字段: {}, 结果: {}'.format(field, data[field]))

            logger.info('结束表单({})字段读取'.format(sheet.name))
            return data

        def merged_dict(src, tar):
            logger.info('合并表单数据...')

            if inindex:
                logger.info("根据索引'{}'合并数据...".format(inindex))
            else:
                logger.info('使用追加（默认）方式合并数据...')

            return merged_defines(src, tar, inindex)

        field_value = {}

        for (name, idx) in self.__defines__:
            sheet = None
            if name in book._sheet_names:
                sheet = book.sheet_by_name(name)
            else:
                if len(book.sheets()) > idx:
                    sheet = book.sheet_by_index(idx)

            if not sheet:
                logger.warning(
                    'not found sheet ({},index: {})'.format(name, idx))
                continue

            field_value = merged_dict(field_value, load_sheet(
                sheet, self.__defines__[(name, idx)]))

        return field_value

    def write(self, field_values, path):

        book = xlutils.copy.copy(self.__source__)

        work_sheets = book._Workbook__worksheets

        def fix_sheet(sheet, define, field_values):
            logger.info(
                "-" * 10 + '准备写入sheet表: {}'.format(sheet.name) + "-" * 10)

            def get_auto_range(escell):
                max = -1
                for cn in define:
                    cur_cell = define[cn]
                    if escell == cur_cell:
                        continue
                    if cur_cell.x1 == escell.x1 and cur_cell.x2 == escell.x2:
                        if cur_cell.name in field_values:
                            col_len = len(field_values[cur_cell.name])
                            if max < col_len:
                                max = col_len

                if max < 0:
                    return [1]
                return [i+1 for i in range(max)]

            for field in define:
                escell = define[field]

                if escell.type == 'ver-auto-add':
                    values = get_auto_range(escell)
                elif field not in field_values:
                    values = '' if escell.type == 'hor' else ['']
                else:
                    values = field_values[field]

                if not isinstance(values, list):
                    values = [values]

                logger.info("准备写入字段'{}'({}), 填入类型: {}".format(
                    field, xlrd.cellname(escell.x1, escell.y1), escell.type))

                for idx, val in enumerate(values):
                    if escell.max > -1 and idx > escell.max - 1:
                        logger.warning('超过最大写入长度,剩余数据将不会写入！！')
                        break
                    x1 = escell.x1 + idx * escell.cell_heigh()
                    x2 = escell.x2 + idx * escell.cell_heigh()
                    y1 = escell.y1
                    y2 = escell.y2

                    logger.info('写入({}): {}'.format(
                        xlrd.cellname(x1, y1), val))
                    sheet.write_merge(x1, x2 - 1, y1, y2 - 1, val)

                    if escell.type == 'hor':
                        break

            logger.info(
                "#" * 10 + "结束sheet表'{}'的写入操作".format(sheet.name) + "#" * 10)

        def fix_sheet_style(stname, out_sheet):
            sheet = self.__source__.sheet_by_name(stname)

            def get_out_cell(sheet, x, y):
                row = sheet._Worksheet__rows.get(x)
                if not row:
                    return None

                cell = row._Row__cells.get(y)
                if not cell:
                    cell = xlwt.Cell.MulBlankCell(x, y, y, 1)
                    row._Row__cells[y] = cell
                return cell

            for x in range(sheet.nrows):
                for y in range(sheet.ncols):
                    cell = sheet.cell(x, y)
                    if cell:
                        style = self.style_list[cell.xf_index]
                        out_xf_id = book.add_style(style)
                        out_cell = get_out_cell(out_sheet, x, y)
                        if out_cell:
                            out_cell.xf_idx = out_xf_id

        for stname, idx in self.__defines__:
            for sheet in work_sheets:
                if sheet.name.lower() == stname.lower():
                    fix_sheet(sheet, self.__defines__[
                              (stname, idx)], field_values)
                    fix_sheet_style(stname, sheet)

        logger.info("将结果保存为'{}'......".format(path))
        book.save(path)
        logger.info("文件保存完毕！")


def read_xls_mod(path):
    logger.info('读取模板: {}......'.format(path))
    book = xlrd.open_workbook(path, formatting_info=True)
    return Esmod(book)


def load_multi(mod, *data_file, index=None):
    datas = {}
    for file in data_file:
        datas = merged_defines(datas, mod.load(file), index)
    return datas


def get_file_info(path):
    def fn_info(fn):
        info = {}
        info['show_name'] = fn
        info['index'] = None
        info['inindex'] = None
        pat = re.match(r'(.*)\.xlsnest\.\$(.+)\.[A-Z]+', fn, re.I)
        if not pat:
            return info

        info['show_name'] = pat.group(1)
        params = pat.group(2)
        for val in params.split('$'):
            kv = val.split('-')
            if len(kv) == 1:
                info[kv[0]] = ''
            else:
                info[kv[0]] = kv[1]

        return info

    info = {}
    det = os.path.split(path)
    info['dir'] = det[0]
    info['name'] = det[1]

    info.update(fn_info(det[1]))

    return info

def get_details(dct,index):
    felst = dct[index]
    details = {}
    for idx, field in enumerate(felst):
        vals = {}
        details[field] = vals
        for key in dct:
            items = dct[key]
            vals[key] = items[idx]
    return details

def remove_blank_row(dct):
    def take_size(dct):
        size = -1
        for k in dct:
            if not isinstance(dct[k], list):
                dct[k] = [dct[k]]
            if size == -1:
                size = len(dct[k])
            elif size != len(dct[k]):
                raise EsMergedException('各列长度不一致，无法清除')
        
        return size
    
    size = take_size(dct)

    ret_rows = []

    for i in range(size):
        row_val = {}
        for k in dct:
            row_val[k] = dct[k][i]
        
        delete = True
        for key in row_val:
            if row_val[key] != '':
                delete = False
                break
        
        if not delete:
            ret_rows.append(row_val)
    
    ret = {}
    for row in ret_rows:
        for k in row:
            if k in ret:
                lst = ret[k]
            else:
                lst = []
                ret[k] = lst
            lst.append(row[k])
    return ret


def merged_defines(src, tar, index=None):
    if len(src) == 0:
        return tar
    if len(tar) == 0:
        return src

    if index and not index == '':
        append_field = []
        for k in tar:
            if not k in src:
                append_field.append(k)

        if index in src and index in tar:
            def format(dct):
                logger.info('{}'.format(dct))
                size = -1
                for k in dct:
                    if not isinstance(dct[k], list):
                        dct[k] = [dct[k]]
                    if size == -1:
                        size = len(dct[k])
                    elif size != len(dct[k]):
                        raise EsMergedException('各列长度不一致，不允许合并')

            format(src)
            format(tar)

            src_details = get_details(src,index)

            empty = list(src_details.values())[0].copy()
            for k in empty:
                empty[k] = ''

            for k1 in src_details:
                for field in append_field:
                    if not field in src_details[k1]:
                        src_details[k1][field] = ''

            tar_details = get_details(tar,index)
            for field in tar_details:
                define = src_details[field] if field in src_details else empty.copy()
                define.update(tar_details[field])
                src_details[field] = define
                # if field == '':
                #     delete = True
                #     for val in define:
                #         if define[val] != '':
                #             delete = False
                #             break

                #     if delete:
                #         del src_details[field]
                #     else:
                #         src_details[field] = define
            
            ret_rows = {}
            for k in src_details:
                dct = src_details[k]
                for field in dct:
                    if not field in ret_rows:
                        lst = []
                        ret_rows[field] = lst
                    else:
                        lst = ret_rows[field]
                    lst.append(dct[field])

            return ret_rows
        else:
            raise EsMergedException('找不到索引对应的字段')

    ret_rows = tar.copy()
    for k in src:
        src_val = src[k]
        if k in tar:
            tar_val = tar[k]
            if not isinstance(src_val, list):
                src_val = [src_val]
            if not isinstance(tar_val, list):
                tar_val = [tar_val]
            src_val.extend(tar_val)
        ret_rows[k] = src_val

    return ret_rows


if __name__ == "__main__":

    path = 'modules/xxx-moudles.xls'
    data_path = 'modules/xxx.xls'

    mod = read_xls_mod(path)
    field_values = load_multi(mod, data_path,data_path,index='name')
    field_values = remove_blank_row(field_values)
    print(field_values)

    # mod.write(field_values, '测试输出数据.xls')
