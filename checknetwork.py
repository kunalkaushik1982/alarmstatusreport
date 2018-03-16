import os
from os import listdir
from collections import OrderedDict
import xlwt
import time
import operator

#path = "C:\Users\ekunkau\Desktop\AlaramStatusReport\INPUT\L1900\NODE LOG"

StartTag = ['st fdd','alt'] #st FDD
Cleanlist = ["", "stopfile=/tmp/", "=====", '$', '>>>','Total: ', 'Node type: ', '---------'] #remove "."


# Create workbook
wb = xlwt.Workbook()

class CommonUtilityNodeDump:

    def __init__(self):
        print 'welcome to CommonUtility class'

    def GetTheSentences(self, starttag, cleanlist, path):
        # stag = starttag
        global SummaryListtup
        SummaryListtup = []
        abab = ''
        ababalt = ''
        abab_fs = ''
        d = {}
        dd = OrderedDict()
        dd_mstatus = OrderedDict()
        proxy = 2
        dictlist = []
        fdddictlist = []
        altdictlist = []
        dictlist_mstatus = []
        onlyfiles = [f for f in listdir(path) if f.endswith('.log')]
        for stag in starttag:
            for x in onlyfiles[:]:
                lines = {}
                lines[x] = []
                etag = x[:-4] + '>'
                sheetname = stag
                for ch in ['|', '^', '.', '>']:
                    if ch in sheetname:
                        sheetname = sheetname.replace(ch, '')
                stag1 = etag + ' ' + stag
                with open(x, 'r') as in_file:
                    for line in in_file:
                        if stag1.lower() in line.strip().lower():  # Test Start Tag
                            break
                    for line in in_file:  # This keeps reading the file
                        if etag.lower() in line.strip().lower():  # Test End Tag
                            break
                        else:
                            if cleanlist[7] in line.strip():
                                proxy = 2
                            elif cleanlist[2] in line.strip() and proxy == 2:
                                proxy = 0
                            elif cleanlist[2] in line.strip() and proxy == 0:
                                proxy = 1
                            if line.strip() != cleanlist[0] and cleanlist[1] not in line.strip() and cleanlist[
                                2] not in line.strip() and cleanlist[3] not in line.strip() and cleanlist[
                                4] not in line.strip() and cleanlist[5] not in line.strip() and cleanlist[
                                6] not in line.strip() and cleanlist[7] not in line.strip():
                                tmpLine = line.strip().split()
                                lines[x].append(tmpLine)
                                if proxy == 0:
                                    if "alt" in stag:
                                        del tmpLine[:]
                                        tmpLine.append('S')
                                        tmpLine.append('Specific Problem')
                                        tmpLine.append('MO (Cause/AdditionalInfo)')
                                    keylist = self.parselist(tmpLine)
                                    indexlist = []
                                    for ite in keylist:
                                        sindex = line.strip().find(ite)
                                        indexlist.append(sindex)
                                    parts = [line.strip()[i:j] for i, j in zip(indexlist, indexlist[1:] + [None])]
                                    parts = [items1.strip(' ') for items1 in parts]
                                    keylist = parts
                                    # proxy = 1
                                if proxy == 1:
                                    if 'cabx'.lower() in stag1.lower() or 'invrx'.lower() in stag1.lower() or 'st fdd'.lower() in stag1.lower():

                                        parts = [line[i:j] for i, j in zip(indexlist, indexlist[1:] + [None])]
                                        parts = [items1.strip(' ') for items1 in parts]

                                        if 'st fdd'.lower() in stag1.lower():
                                            for n, item in enumerate(parts):
                                                if item == '1 (UNLOCKED)':
                                                    parts[n] = 'U'
                                                    var_adm = 'U'
                                                elif item == '0 (LOCKED)':
                                                    parts[n] = 'L'
                                                    var_adm = 'L'
                                                elif item == '1 (ENABLED)':
                                                    parts[n] = 'E'
                                                    var_op = 'E'
                                                elif item == '0 (DISABLED)':
                                                    parts[n] = 'D'
                                                    var_op = 'D'
                                                elif item.startswith('ENodeBFunction=1,EUtranCellFDD='):#
                                                    itemaa = item.strip('\n').split(',')
                                                    itembb =itemaa[1].split('=')
                                                    itemcc = itembb[1]
                                                    var_car = itemcc[:1]
                                                    var_sec = itemcc[-2:]
                                                    parts[n] = var_car+','+var_sec

                                        var_status = var_car + ',' + var_sec + ',' + var_adm + ',' + var_op
                                        var_status1 = var_car + ',' + var_sec + ',' + var_adm + ',' + var_op + '#$'
                                        abab += var_status1
                                        if (var_op == 'D' and var_adm == 'U') and (var_op == 'E' and var_adm == 'L'):
                                            abab_fs += 'NOK' + '#$'
                                        elif(var_op == 'E' and var_adm == 'U') or (var_op == 'D' and var_adm == 'L'):
                                            abab_fs += 'OK'+ '#$'
                                        elif var_op == "" or var_adm == "":
                                            abab_fs += 'NOK'+ '#$'
                                    else:

                                        parts = [line.strip()[i:j] for i, j in zip(indexlist, indexlist[1:] + [None])]
                                        parts = [items1.strip(' ') for items1 in parts]

                                        if 'alt'.lower() in stag1.lower():
                                            var_altstatus = parts[1] + '#$'
                                        ababalt += var_altstatus
                                    tmpLine = parts
                                    l1 = len(keylist)
                                    l2 = len(tmpLine)
                                    for i in range(0, len(keylist)):
                                        if l1 == l2 and keylist[i] != tmpLine[i] and not keylist[i].strip(':').isdigit():
                                            dd[keylist[i]] = tmpLine[i].strip()
                    if 'st fdd'.lower() in stag1.lower() or 'alt'.lower() in stag1.lower():
                        if 'st fdd'.lower() in stag1.lower():
                            dd_mstatus['!!NodeName!!'] = x[:-4]
                            dd_mstatus['!!M_Status!!'] = abab
                            final_abab_fs = ''
                            if abab_fs == '':
                                final_abab_fs = 'NOK'
                            elif abab_fs != '':
                                abab_fs_list = abab_fs.split('#$')
                                del abab_fs_list[-1]
                                if 'NOK' in abab_fs_list:
                                    final_abab_fs = 'NOK'
                                else:
                                    final_abab_fs = 'OK'
                            dd_mstatus['!!FINAL_Status!!'] = final_abab_fs
                            dictlist_mstatus.append(dd_mstatus)
                            dd_mstatus = OrderedDict()
                            abab = ''
                            abab_fs = ''
                            dictlist = self.merge_lists(dictlist, dictlist_mstatus, '!!NodeName!!')
                        elif 'alt'.lower() in stag1.lower() and ababalt:
                            dd_mstatus['!!NodeName!!'] = x[:-4]
                            dd_mstatus['!!M_Status!!'] = ababalt
                            dd_mstatus['!!FINAL_Status!!'] = 'NOK'

                            dictlist_mstatus.append(dd_mstatus)
                            dd_mstatus = OrderedDict()
                            ababalt = ''
                            dictlist = self.merge_lists(dictlist, dictlist_mstatus, '!!NodeName!!')
                    proxy = 2
                etag = ""
            # FOR DELETING COLUMNS FROM FDD COMMANDS
            dictlist= [{k: v for k, v in d.iteritems() if k != 'MO'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != '!!Status!!'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != 'Adm State'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != 'Proxy'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != 'Op. State'} for d in dictlist]

            #FOR DELETING COLUMNS FROM ALT COMMANDS
            dictlist = [{k: v for k, v in d.iteritems() if k != 'MO (Cause/AdditionalInfo)'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != 'S'} for d in dictlist]
            dictlist = [{k: v for k, v in d.iteritems() if k != 'Specific Problem'} for d in dictlist]

            sorting_key = operator.itemgetter("!!NodeName!!")
            dictlist = sorted(dictlist, key=sorting_key)
            self.writetoexcel(dictlist, stag[-10:])

            if stag == 'st fdd':
                for item in dictlist:
                    del item['!!M_Status!!']
                    valuev = item['!!FINAL_Status!!']
                    del item['!!FINAL_Status!!']
                    item['FDD'] = valuev
                    fdddictlist.append(item)
            elif stag == 'alt':
                for item in dictlist:
                    del item['!!M_Status!!']
                    valuev = item['!!FINAL_Status!!']
                    del item['!!FINAL_Status!!']
                    item['ALT'] = valuev
                    altdictlist.append(item)

            dictlist = []
            dictlist_mstatus = []
        sorting_key = operator.itemgetter("!!NodeName!!")
        fdddictlist = sorted(fdddictlist, key=sorting_key)
        altdictlist = sorted(altdictlist, key=sorting_key)
        for i in fdddictlist:
            for j in altdictlist:
                if i['!!NodeName!!'] == j['!!NodeName!!']:
                    i.update(j)

        for i in fdddictlist:
            fddcheck = ''
            altcheck = ''
            finalcheck = ''
            if 'FDD' in i:
                aa = i['FDD']
                if aa == 'OK':
                    fddcheck = 'OK'
                elif aa == 'NOK':
                    fddcheck = 'NOK'
            if 'ALT' in i:
                bb = i['ALT']
                if bb == 'OK':
                    altcheck = 'OK'
                elif bb == 'NOK':
                    altcheck = 'NOK'

            if (fddcheck == 'OK' and altcheck == 'OK') or (fddcheck == 'OK' and altcheck == ''):
                finalcheck = 'OK'
            else:
                finalcheck = 'NOK'

            i['FINALCHECK'] = finalcheck

        fdddictlist = sorted(fdddictlist, key=sorting_key)
        self.writetoexcel(fdddictlist, 'Summry')
        print fdddictlist

    def merge_lists(self,l1, l2, key):
        merged = {}
        for item in l1 + l2:
            if item[key] in merged:
                merged[item[key]].update(item)
            else:
                merged[item[key]] = item
        return [val for (_, val) in merged.items()]

    def writetoexcel(self,data, sheetname):

        ws = wb.add_sheet(sheetname)
        all_keys = reduce(lambda x, y: x.union(y.keys()), data, set())
        headers = list(all_keys)

        if "MO" in headers and "!!NodeName!!" in headers:
            headers.remove("!!NodeName!!")
            headers.insert(0, "!!NodeName!!")
            headers.remove("MO")
            headers.insert(1, "MO")
        elif "!!FINAL_Status!!" in headers and "!!NodeName!!" in headers and "!!M_Status!!" in headers:
            headers.remove("!!NodeName!!")
            headers.remove("!!FINAL_Status!!")
            headers.remove("!!M_Status!!")
            headers.insert(0, "!!NodeName!!")
            headers.insert(1, "!!M_Status!!")
            headers.insert(2, "!!FINAL_Status!!")
        elif "ALT" in headers and "!!NodeName!!" in headers and "FDD" in headers and "FINALCHECK" in headers:
            headers.remove("!!NodeName!!")
            headers.remove("ALT")
            headers.remove("FDD")
            headers.remove("FINALCHECK")
            headers.insert(0, "!!NodeName!!")
            headers.insert(1, "ALT")
            headers.insert(2, "FDD")
            headers.insert(3, "FINALCHECK")
        else:
            pass

        # enumerate() function adds a counter to an iterable.
        for column, header in enumerate(headers):
            ws.write(0, column, header)
        # Write data #By default, enumerate() starts counting at 0 but if you give it a second integer argument, it'll start from that number instead:
        for row, row_data in enumerate(data, start=1):
            for column, key in enumerate(headers):
                value = row_data.get(key)
                if value is None:
                    ws.write(row, column, '')
                else:
                    ws.write(row, column, row_data[key])

        # Save file
        wb.save("test.xls")


    def parselist(self, tmplist):
        try:
            aa = []
            aa = tmplist
            for i in range(0, len(aa)):
                if aa[i] == '(UNLOCKED)' or aa[i] == '(ENABLED)' or aa[i] == '(NOT_BARRED)' or aa[i] == '(STATUS)' or \
                                aa[i] == '(SET)' or \
                                aa[i] == '(LOCKED)' or aa[i] == '(DISABLED)' or aa[i] == '(BARRED)' or aa[
                    i] == '(NO_STATUS)' or aa[i] == '(NOT_SET)' or \
                                aa[i] == '=' or aa[i] == 'NR' or aa[i] == '(cellId,PCI)' or aa[i] == '(RL1)' or aa[
                    i] == '(RL2)' or aa[i] == '(RL3)' or \
                                aa[i] == '(RL4)' or aa[i] == '(LNH)' or aa[i] == '(localCellIds/CellIds,PCIs)' or aa[
                    i] == '(RL)' or aa[i] == '(W/dBm)' or \
                                aa[i] == ' - MO2' or aa[i] == 'State':
                    bb = aa[i - 1] + ' ' + aa[i]
                    abx = i - 1
                    aa.pop(i)
                    aa.pop(i - 1)
                    aa.insert(abx, bb)

            return aa
        except:
            return aa



if __name__ == "__main__":
    import sys
    time1 = time.time()

    #owd = os.getcwd()
    #os.chdir(path)

    #cund = CommonUtilityNodeDump()
    #cund.GetTheSentences(StartTag, Cleanlist)

    #os.chdir(owd)

    TotalTime = time.time() - time1
    print 'Total Time of Execution is: ' + str(TotalTime)