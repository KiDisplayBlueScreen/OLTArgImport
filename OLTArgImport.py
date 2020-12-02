import xlwings as xw

ExcelPath = "D:\OLT参数导入步骤\已填写参数\播植新合HW-OLT02-MA58002槽(未填写).xls"

BASEIP = "119.146.99.73"
NAS_SLOT = 7
NAS_PORT = 4
NAS_SUBSLOT = 0


def WriteArg(Sheet, BaseRow, SVLAN, CVLAN_BEGIN, CVLAN_STEP, CVLAN_NUM):
    TempBaseRow = BaseRow  # BaseRow = 2
    LastRow = BaseRow + 112
    i = 0
    while TempBaseRow <= LastRow:
        SVLAN_Index = "D" + str(TempBaseRow)
        CVLAN_Index = "H" + str(TempBaseRow)
        CVLANSTEP_Index = "I" + str(TempBaseRow)
        CVLANNUM_Index = "J" + str(TempBaseRow)
        BASEIP_Index = "K" + str(TempBaseRow)
        NASSLOT_Index = "L" + str(TempBaseRow)
        NASPORT_Index = "M" + str(TempBaseRow)
        NASSUBPORT_Index = "N" + str(TempBaseRow)

        Sheet.range(CVLANSTEP_Index).value = CVLAN_STEP

        Sheet.range(BASEIP_Index).value = BASEIP
        Sheet.range(NASSLOT_Index).value = NAS_SLOT
        Sheet.range(NASPORT_Index).value = NAS_PORT
        Sheet.range(NASSUBPORT_Index).value = NAS_SUBSLOT
        if i <= 3:
            Sheet.range(SVLAN_Index).value = SVLAN[0]
            Sheet.range(CVLAN_Index).value = CVLAN_BEGIN[i]
            Sheet.range(CVLANNUM_Index).value = CVLAN_NUM[i]
        else:
            Sheet.range(SVLAN_Index).value = SVLAN[1]
            Sheet.range(CVLAN_Index).value = CVLAN_BEGIN[i - 4]
            Sheet.range(CVLANNUM_Index).value = CVLAN_NUM[i - 4]
        TempBaseRow = TempBaseRow + 16
        i = i + 1


if __name__ == '__main__':
    WorkBook = xw.Book(ExcelPath)
    Sheet = WorkBook.sheets["sheet1"]

    SVLAN = [2403, 2404]

    ADSLD_CVLAN = [101, 165, 229, 293]
    ADSLD_CVLAN_NUM = [64, 64, 64, 64]

    IPTV_CVLAN = [45, 45, 45, 45]
    IPTV_CVLAN_NUM = [1, 1, 1, 1]

    VPN_CVLAN = [485, 517, 549, 581]
    VPN_CVLAN_NUM = [32, 32, 32, 32]

    WIFI_CVLAN = [2601, 2665, 2729, 2793]
    WIFI_CVLAN_NUM = [64, 64, 64, 64]

    YZJBDJRSB_CVLAN = [46, 46, 46, 46]
    YZJBDJRSB_CVLAN_NUM = [1, 1, 1, 1]

    WriteArg(Sheet, 2, SVLAN, ADSLD_CVLAN, 1, ADSLD_CVLAN_NUM)
    WriteArg(Sheet, 7, SVLAN, IPTV_CVLAN, 1, IPTV_CVLAN_NUM)
    WriteArg(Sheet, 11, SVLAN, WIFI_CVLAN, 1, WIFI_CVLAN_NUM)
    WriteArg(Sheet, 12, SVLAN, WIFI_CVLAN, 1, WIFI_CVLAN_NUM)
    WriteArg(Sheet, 14, SVLAN, YZJBDJRSB_CVLAN, 1, YZJBDJRSB_CVLAN_NUM)
