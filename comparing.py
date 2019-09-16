# coding=utf-8
import xlrd
import xlwt
import xlutils.copy
from datetime import datetime
from xlrd import xldate_as_tuple
import traceback
import re
import copy


def compare(oldFile, newFile, outputFile, cumulative_compare, comments=True, comment_pattern="DPL Comment",
            table_split="^Table\s*\d+$",
            max_columns=30):
    """compare based on the directions, outputD for out put"""
    # print "COMMENT PATTERN: ", comment_pattern
    oldwb = xlrd.open_workbook(oldFile)
    newwb = xlrd.open_workbook(newFile, formatting_info=True)

    yield "Pairing the files....\n"
    paired, OldExtra, NewExtra = pairFiles(oldwb, newwb)

    if (len(OldExtra) > 0):
        sheets = "\n".join(OldExtra)
        yield "Warning: Sheet \n {}\nare only in Old File!\n".format(sheets)
    if (len(NewExtra) > 0):
        sheets = "\n".join(NewExtra)
        yield "Warning: Sheet \n {}\nare only in New File!\n".format(sheets)

    yield "\nComparing the files....\n"

    outwb = xlutils.copy.copy(newwb)
    outwb.set_colour_RGB(0x2A, 100, 200, 100)  # set light green color for output

    def get_sheet_by_name(book, name):
        import itertools
        try:
            for idx in itertools.count():
                sheet = book.get_sheet(idx)
                if sheet.name == name:
                    return sheet
        except IndexError:
            return None

    for k, v in paired.iteritems():
        if v is not None:
            sheetOld = oldwb.sheet_by_name(v)
            sheetNew = newwb.sheet_by_name(k)
            sheetOut = get_sheet_by_name(outwb, k)
        else:
            sheetOld = None
            sheetNew = newwb.sheet_by_name(k)
            sheetOut = get_sheet_by_name(outwb, k)

        for txt in GDOcomMain(cumulative_compare, k, sheetOut, sheetNew, sheetOld, comments, comment_pattern,
                              table_split, max_columns):
            yield txt
    outwb.save(outputFile)


def GDOcomMain(cumulative_compare, sheetName, sheetOut, sheetNew, sheetOld, comments, comment_pattern, table_split,
               max_columns):
    """ Main Comparison
    """
    yield "Comparing Sheet {}.".format(sheetName)
    if sheetOld is None:  # For a New Only Sheet
        for txt in copyNew(sheetNew, sheetOut):
            yield txt
    else:
        has_comment, datapool = readOld(sheetOld, comment_pattern, table_split,
                                        max_columns)  # Read Data Pool from Old Sheet.
        nTable = len(datapool)
        if nTable > 1:
            nTable -= 1
        yield sheetName + " has " + str(nTable) + " table(s)."
        # Comparing file can read new sheet
        for txt in typeCompare(sheetName, sheetOut, sheetNew, sheetOld, datapool, has_comment, table_split,
                               comment_pattern,
                               comments, max_columns):
            yield txt

        if cumulative_compare == True:
            newhas_comment, newdatapool = readNew(sheetNew, comment_pattern, table_split,
                                                  max_columns)  # Read Data Pool from new Sheet.
            newTable = len(newdatapool)
            if newTable > 1:
                has_multitable = True
            else:
                has_multitable = False
            # Comparing file can read old sheet
            for txt in oldtypeCompare(sheetName, sheetOut, sheetNew, sheetOld, newdatapool, newhas_comment,
                                      has_multitable, table_split,
                                      comment_pattern, comments, max_columns):
                yield txt


def hash_key(thisrow, comment_column):
    """ define hash key
    """
    return u"|".join(thisrow[:comment_column + 1])


def readOld(sheet, comment_pattern, table_split, max_columns):
    """ readOld sheet
    """
    import re
    p = re.compile(table_split, re.UNICODE)  # split pattern

    datapool = []  # replied data pool
    bufs = []  # buffers to store rows

    has_comment = False

    buf = []
    thispool = {}
    commentTag = None
    newpart = False
    for row in xrange(sheet.nrows):
        thisrow = [u'' for i in range(max_columns)]  # fill row with "" according to max_columns
        for col in xrange(sheet.ncols):
            cell = sheet.cell(row, col)
            v = unicode(cell.value).strip()
            if p.match(v):
                newpart = True  # split flag
            if v.upper() == comment_pattern.upper():  # detect comment_pattern
                if commentTag is not None:
                    assert col == commentTag, "Comments should in the same column in old Sheet {}".format(sheet.name)
                commentTag = col
                has_comment = True
            thisrow[col] = v

        buf.append(thisrow)
        if newpart:
            datapool.append((thispool, commentTag))
            bufs.append(buf)
            thispool = {}
            commentTag = None
            buf = []
            newpart = False

    datapool.append((thispool, commentTag))  # The last part
    bufs.append(buf)

    for (pool, Tag), buf in zip(datapool, bufs):  # fill data pool
        for thisrow in buf:
            if Tag is None:  # if there is no comments
                key = hash_key(thisrow, len(thisrow) - 1)
                pool[key] = ""
            else:
                key = hash_key(thisrow, Tag - 1)  # if there is comments
                pool[key] = thisrow[Tag]
    return has_comment, datapool


def readNew(sheet, comment_pattern, table_split, max_columns):
    """ readNew sheet
    """
    import re
    p = re.compile(table_split, re.UNICODE)  # split pattern
    newdatapool = []  # replied data pool
    bufs = []  # buffers to store rows
    newhas_comment = False
    buf = []
    thispool = {}
    commentTag = None
    newpart = False
    for row in xrange(sheet.nrows):
        thisrow = [u'' for i in range(max_columns)]  # fill row with "" according to max_columns
        for col in xrange(sheet.ncols):
            cell = sheet.cell(row, col)
            v = unicode(cell.value).strip()
            if p.match(v):
                newpart = True  # split flag
            if v.upper() == comment_pattern.upper():  # detect comment_pattern
                if commentTag is not None:
                    assert col == commentTag, "Comments should in the same column in old Sheet {}".format(sheet.name)
                commentTag = col
                newhas_comment = True
            thisrow[col] = v

        buf.append(thisrow)
        if newpart:
            newdatapool.append((thispool, commentTag))
            bufs.append(buf)
            thispool = {}
            commentTag = None
            buf = []
            newpart = False

    newdatapool.append((thispool, commentTag))  # The last part
    bufs.append(buf)

    for (pool, Tag), buf in zip(newdatapool, bufs):  # fill data pool
        for thisrow in buf:
            if Tag is None:  # if there is no comments
                key = hash_key(thisrow, len(thisrow) - 1)
                pool[key] = ""
            else:
                key = hash_key(thisrow, Tag - 1)  # if there is comments
                pool[key] = thisrow[Tag]
    return newhas_comment, newdatapool


def copyNew(sheetNew, sheetOut, color='yellow'):
    """Annotate a new sheet """
    for row in range(sheetNew.nrows):
        for col in range(sheetNew.ncols):
            v = sheetNew.cell(row, col).value
            style = xlwt.easyxf('pattern: pattern solid, fore_colour {}'.format(color))
            if v == u"":  # only change the null cells
                sheetOut.write(row, col, v, style)
    yield "Sheet {} is new one and annotated in YELLOW.\n".format(sheetOut.name)


def typeCompare(sheetName, sheetOut, sheetNew, sheetOld, datapool, has_comment, table_split, comment_pattern, comments,
                max_columns, color='0x2A'):
    """ Comparing sheets and output
    """
    import re
    p = re.compile(table_split, re.UNICODE)  # table split pattern
    recordpool = []
    record = []
    newpart = False
    partid = 0
    commentTag = datapool[partid][1]

    # Check the tables in new sheet
    for row in xrange(sheetNew.nrows):
        for col in xrange(sheetNew.ncols):
            cell = sheetNew.cell(row, col)
            v = unicode(cell.value).strip()
            if p.match(v):
                newpart = True
            if v.upper() == comment_pattern.upper():
                assert col == commentTag, "Comments should in the same column in new Sheet {}".format(sheetNew.name)
        record.append(row)
        if newpart:
            recordpool.append((record, commentTag))
            record = []
            partid += 1
            try:
                commentTag = datapool[partid][1]
            except:
                raise ValueError("Sheet {} has different number of tables".format(sheetName))
            newpart = False

    recordpool.append((record, commentTag))
    assert len(recordpool) == len(datapool), "Sheet {} has different number of tables".format(sheetName)

    for partid, (record, Tag) in enumerate(recordpool):
        for row in record:
            thisrow = [u'' for i in range(max_columns)]  # fill the this row with null according to max_columns
            for col in xrange(sheetNew.ncols):
                cell = sheetNew.cell(row, col)
                v = unicode(cell.value).strip()
                thisrow[col] = v

            if Tag is None:  # define key by commentTag
                key = hash_key(thisrow, len(thisrow) - 1)
            else:
                key = hash_key(thisrow, Tag - 1)
            if key in datapool[partid][0]:
                # old records
                oldTag = datapool[partid][1]
                if oldTag is not None and comments:
                    sheetOut.write(row, oldTag, datapool[partid][0][key])
            else:
                # new records
                for col in xrange(sheetNew.ncols):
                    cell = sheetNew.cell(row, col)
                    v = cell.value
                    is_date = False

                    try:  # detect if the format is possible date field
                        date = float(v)
                        if date > 10000.0 and date < 99999.0:
                            is_date = True
                    except:
                        pass

                    if not is_date:
                        style = xlwt.easyxf(
                            'pattern: pattern solid, fore_colour {}; align: horiz center; border:top thin,bottom thin;'.format(
                                color))
                        sheetOut.write(row, col, v, style)

    yield "Sheet {} is OK. New items are annotated in GREEN\n".format(sheetOut.name)


def oldtypeCompare(sheetName, sheetOut, sheetNew, sheetOld, newdatapool, newhas_comment, has_multitable, table_split,
                   comment_pattern,
                   comments, max_columns, color='0x16'):
    import re
    p = re.compile(table_split, re.UNICODE)  # table split pattern
    add_row = sheetNew.nrows
    recordpool = []
    record = []
    newpart = False
    partid = 0
    commentTag = newdatapool[partid][1]
    comment_col = []  # 存储表的comment pattern所在的列数
    # Check the tables in new sheet
    for row in xrange(sheetOld.nrows):
        for col in xrange(sheetOld.ncols):
            cell = sheetOld.cell(row, col)
            v = unicode(cell.value).strip()
            if p.match(v):
                newpart = True
            if v.upper() == comment_pattern.upper():
                # print "old comment row: ", row
                comment_col.append(col)
        record.append(row)
        if newpart:
            recordpool.append((record, commentTag))
            record = []
            partid += 1
            try:
                commentTag = newdatapool[partid][1]
            except:
                raise ValueError("Sheet {} has different number of tables".format(sheetName))
            newpart = False
    recordpool.append((record, commentTag))
    assert len(recordpool) == len(newdatapool), "Sheet {} has different number of tables".format(sheetName)
    # print "comment_col: ", comment_col
    for partid, (record, Tag) in enumerate(recordpool):
        for row in record:
            thisrow = [u'' for i in range(max_columns)]  # fill this row with null according to max_columns
            for col in xrange(sheetOld.ncols):
                cell = sheetOld.cell(row, col)
                v = unicode(cell.value).strip()
                thisrow[col] = v
            if Tag is None:
                key = hash_key(thisrow, len(thisrow) - 1)
            else:
                key = hash_key(thisrow, Tag - 1)
            # print "处理前的key: ", key
            no_handle_key = copy.deepcopy(key)  # 处理前的key
            # 支持comment列填写任何备注
            is_comment = None
            is_comment_no = None  # 此flag主要作用就是当用户的旧表Comment列忘记填写Comment的备注，但是确实是有效数据
            key1 = key.split("|")
            key1 = list(reversed(key1))
            for i in range(len(key1)):
                if key1[i] != "":
                    key1[i] = ""
                    is_comment = max_columns - i - 1
                    is_comment_no = is_comment + 1  # Comment pattern 所在列的数据为空时就会导致向前推进一个单元格
                    break
            key = list(reversed(key1))
            key = "|".join(key)
            have_handle_key = copy.deepcopy(key)  # 处理后的key
            # print "处理后的key： ", key
            # print "is_comment: ", is_comment
            # print "is_comment_no: ", is_comment_no
            # 需要考虑Comment pattern列用户忘记填写时，如果旧表有但是新表没有时就写入新表

            if is_comment_no in comment_col:
                # 说明就是Comment Pattern列，只是用户忘记填写comment备注，这时候前面的步骤会将Comment 列前面有数据的
                # 单元格变为""，所以需要还原到处理前的数据
                key = no_handle_key
                is_comment = True
            elif is_comment in comment_col:  # 说明是在Comment Pattern列的数据
                key = have_handle_key
                is_comment = True
            else:
                # 其他情况也需要还原处理前的数据
                key = no_handle_key
                is_comment = True

            # print "最终判断的key: ", key
            # print "新表的数据: ", newdatapool[partid][0]
            if key in newdatapool[partid][0]:
                pass
            else:
                if is_comment:
                    try:
                        # print "旧表中存在新表中没有的数据: ", key
                        for col in xrange(sheetOld.ncols):
                            cell = sheetOld.cell(row, col).value
                            v = cell
                            is_date = False
                            # print "读出的数据: ", v
                            try:
                                date = str(v)
                                date1 = date.split(".")
                                if len(date1) > 1:
                                    if len(date1[1]) > 1 or len(date1[1]) == 1:
                                        is_date = True
                            except Exception, e:
                                print e
                                pass
                            if not is_date:
                                style = xlwt.easyxf(
                                    'pattern: pattern solid, fore_colour {}; align: horiz center; border:top thin,bottom thin;'.format(
                                        color))
                                sheetOut.row(add_row + 1).height = 300
                                try:
                                    sheetOut.write(add_row + 1, col, v, style)
                                except Exception, e:
                                    print "写入sheetOut表出错！", e
                            else:
                                try:
                                    cell = float(cell)
                                    date_time = datetime(*xldate_as_tuple(cell, 0))
                                    v = date_time.strftime('%d-%m-%Y')
                                    sheetOut.row(add_row + 1).height = 300
                                    sheetOut.write(add_row + 1, col, v)
                                except Exception, e:
                                    sheetOut.row(add_row + 1).height = 300
                                    sheetOut.write(add_row + 1, col, v)
                                    print e
                        add_row += 1
                    except Exception, e:
                        print traceback.format_exc(e)

    yield "Sheet {} is OK. Old sheet have but New sheet no and annotated in Grey\n".format(sheetOut.name)


def pairFiles(oldwb, newwb):
    """paired the sheets based on format"""
    newSheets = {k.strip(): k for k in newwb.sheet_names()}
    oldSheets = {k.strip(): k for k in oldwb.sheet_names()}

    oldExtra = set(oldSheets.keys()) - set(newSheets.keys())
    newExtra = set(newSheets.keys()) - set(oldSheets.keys())

    paired = {}
    for s, v in newSheets.iteritems():
        if s in oldSheets:
            paired[v] = oldSheets[s]
        else:
            paired[v] = None

    return paired, oldExtra, newExtra


# Consolidate Comment function(合并comment)
def consolidate_compare(con_oldD, consolidate_newD, outputCon, COMMENT_PATTERN, TABLE_SPLIT):
    yield "Consolidate the files....\n"
    consolidate_newD = xlrd.open_workbook(consolidate_newD, formatting_info=True)
    outwb = xlutils.copy.copy(consolidate_newD)
    oldD = []
    for i in range(len(con_oldD)):  # con_oldD[i]代表每一个工作簿
        oldD.append(xlrd.open_workbook(con_oldD[i]))
    con_paired = con_pairFiles(consolidate_newD)
    print "总的工作表: ", con_paired
    for i in range(len(con_paired)):
        try:
            sheet = consolidate_newD.sheet_by_name(con_paired[i])
            sheetOut = get_sheet_by_name(outwb, con_paired[i])
            valid_data, comment_col = read_sheet(sheet, COMMENT_PATTERN, TABLE_SPLIT)
            valid_datas = []  # 存储comment列没有数据的行
            if comment_col:
                for m in range(len(valid_data)):
                    row = valid_data[m][0]  # 行
                    row_data = valid_data[m][1]  # 行对应的数据
                    row_data_length = len(row_data)  # 行数据的长度
                    for n in range(len(comment_col)):
                        if comment_col[n] + 1 == row_data_length:
                            comment_data = row_data[comment_col[n]]
                            if comment_data == "":  # 评论列没有数据
                                valid_datas.append((row, row_data))
                Consolidate_main(sheet, sheetOut, valid_datas, oldD, TABLE_SPLIT, COMMENT_PATTERN)
            else:
                print "No match {} in sheet {} \n".format(COMMENT_PATTERN, con_paired[i])
                yield "No match {} in sheet {} \n".format(COMMENT_PATTERN, con_paired[i])
        except Exception, e:
            print traceback.format_exc(e)
    outwb.save(outputCon)


def get_sheet_by_name(book, name):
    import itertools
    try:
        for idx in itertools.count():
            sheet = book.get_sheet(idx)
            if sheet.name == name:
                return sheet
    except IndexError:
        return None


# 读取表的内容
def read_sheet(sheet, COMMENT_PATTERN, TABLE_SPLIT):
    comment_row = None
    comment_col = None
    valid_row = None  # 有效数据的最后一行
    newpart = False
    valid_data = []
    p = re.compile(TABLE_SPLIT, re.UNICODE)
    # 匹配Comment Pattern所在的行和列
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell = sheet.cell(row, col)
            v = unicode(cell.value).strip()
            if p.match(v):
                newpart = True
            if v.upper() == COMMENT_PATTERN.upper():
                comment_row = row
                comment_col = col
                break
    if comment_row:
        if newpart:
            print "多个table表{}".format(sheet.name)
            # 有多个table
            comment_rows = []  # 存储多个comment的行
            comment_cols = []  # 存储多个comment的列
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    cell = sheet.cell(row, col)
                    v = unicode(cell.value).strip()
                    if v.upper() == COMMENT_PATTERN.upper():
                        comment_rows.append(row)
                        comment_cols.append(col)
            # print "comment_rows: ", comment_rows
            # print "comment_cols: ", comment_cols
            # Detect valid data row
            valid_rows = []  # 存储有效结束行
            for row in range(len(comment_rows)):
                need_for = False
                for i in range(comment_rows[row] + 1, sheet.nrows):
                    if need_for:
                        valid_rows.append(i - 2)
                        break
                    for j in range(len(sheet.row_values(i))):
                        if j < 2:
                            if sheet.row_values(i)[j] == "":
                                need_for = True
                                break

            for row1 in range(len(comment_rows)):
                row2 = valid_rows[row1]
                comment_col = comment_cols[row1]
                for row in range(comment_rows[row1] + 1, row2 + 1):
                    thisrow = [u'' for i in range(comment_col + 1)]
                    for col in range(comment_col + 1):
                        cell = sheet.cell(row, col)
                        v = unicode(cell.value).strip()
                        thisrow[col] = v
                    row_data = (row, thisrow)
                    valid_data.append(row_data)
            return valid_data, comment_cols
        else:
            print "单个table表{}".format(sheet.name)
            comment_cols = []
            # 从comment row向下读如果遇到前两列中有空的单元格就认为有效数据结束
            need_for = False
            comment_cols.append(comment_col)
            for i in range(comment_row + 1, sheet.nrows):
                if need_for:
                    break
                for j in range(len(sheet.row_values(i))):
                    if j < 2:
                        # print sheet.row_values(i)[j]
                        if sheet.row_values(i)[j] == "":
                            valid_row = i - 1
                            need_for = True
                            break
            # 将有效数据放入字典
            # print "comment_row: ", comment_row
            # print "vaild_row: ", valid_row
            if valid_row:
                for row in range(comment_row + 1, valid_row + 1):
                    thisrow = [u'' for i in range(comment_col + 1)]
                    for col in range(comment_col + 1):
                        cell = sheet.cell(row, col)
                        v = unicode(cell.value).strip()
                        thisrow[col] = v
                    row_data = (row, thisrow)
                    valid_data.append(row_data)
                return valid_data, comment_cols
            else:
                # 当有效数据就是最后一行时.这时候vaild_row为None
                for row in range(comment_row + 1, sheet.nrows):
                    thisrow = [u'' for i in range(comment_col + 1)]
                    for col in range(comment_col + 1):
                        cell = sheet.cell(row, col)
                        v = unicode(cell.value).strip()
                        thisrow[col] = v
                    row_data = (row, thisrow)
                    valid_data.append(row_data)
                return valid_data, comment_cols
    else:
        print "没有comment_row {}".format(sheet.name)
        comment_cols = None
        return valid_data, comment_cols


def con_pairFiles(consolidate_file):
    con_paired = [k for k in consolidate_file.sheet_names()]
    return con_paired


def Consolidate_main(sheet, sheetOut, valid_datas, oldD, TABLE_SPLIT, COMMENT_PATTERN):
    try:
        for i in range(len(oldD)):  # len(olD)就是有几个工作簿
            con_paired = con_pairFiles(oldD[i])
            for j in range(len(con_paired)):  # len(con_paired)每个工作簿下有几张工作表
                file = oldD[i].sheet_by_name(con_paired[j])
                # 找到相同的sheet表再进行comment合成
                if file.name == sheet.name:
                    # 读取表的内容
                    file_data, file_col = read_sheet(file, COMMENT_PATTERN, TABLE_SPLIT)  # 读取该表的内容
                    file_datas = []
                    if file_col:
                        # 读取fila_data里面行的comment列有数据的行
                        for p in range(len(file_data)):
                            row_data = file_data[p][1]
                            row_data_length = len(row_data)
                            for n in range(len(file_col)):
                                if file_col[n] + 1 == row_data_length:
                                    comment_data = row_data[file_col[n]]
                                    if comment_data != "":  # 存储comment列有数据的行
                                        file_datas.append((file_data[p][0], row_data))
                        count = 0
                        for m in range(len(valid_datas)):
                            if count == len(valid_datas):  # 循环次数取决于主表缺少comment的行数
                                break
                            row = valid_datas[m][0]  # 行号
                            main_data = valid_datas[m][1]  # 行
                            for k in range(len(file_datas)):
                                file_row_data = file_datas[k][1]  # 行
                                temp_data = copy.deepcopy(file_row_data)
                                temp_data_length = len(temp_data)
                                # 将comment列对应的数据去掉
                                add_comment_col = None
                                comment = None
                                for x in range(len(file_col)):
                                    if file_col[x] + 1 == temp_data_length:
                                        add_comment_col = file_col[x]
                                        comment = temp_data[file_col[x]]
                                        temp_data[file_col[x]] = u""
                                if temp_data == main_data:
                                    sheetOut.write(row, add_comment_col, comment)
                            count += 1
                    else:
                        print "没有找到comment列不进行合并"
    except Exception, e:
        print traceback.format_exc(e)
