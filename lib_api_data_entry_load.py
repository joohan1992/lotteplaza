from openpyxl import load_workbook, Workbook


def read_vendor_code(g_data):
    g_data['dict_vendor_item_code'] = dict()
    for item in g_data['conf_vendor_item_code']:
        vc = g_data['conf_vendor_item_code'][item]['vc']
        file_nm = g_data['conf_vendor_item_code'][item]['name']
        sheet_nm = g_data['conf_vendor_item_code'][item]['sheet']
        sc_type = g_data['conf_vendor_item_code'][item]['sc_type']
        sc = g_data['conf_vendor_item_code'][item]['sc']
        prefix = g_data['conf_vendor_item_code'][item]['prefix']
        pos_start, pos_ic, pos_upc, pos_desc, pos_csize = g_data['conf_vendor_item_code'][item]['pos']

        wb_r = load_workbook('./vendor_data/'+file_nm)
        ws_r = wb_r[sheet_nm]
        dict_IC = dict()
        dict_UPC = dict()
        idx1 = 0
        for row in ws_r.rows:
            # 컬럼 명 통과
            if idx1 >= pos_start:
                tmp_ic = str(row[pos_ic].value)
                if prefix:
                    tmp_ic = g_data['dict_vend'][vc]['prefix']+tmp_ic
                tmp_upc = str(row[pos_upc].value).replace('-', '').replace(' ', '')
                tmp_desc = str(row[pos_desc].value)
                tmp_csize = str(row[pos_csize].value)
                if row[pos_csize].value is None or row[pos_csize].value == '':
                    continue

                if not tmp_ic or tmp_ic == '':
                    continue

                if tmp_ic not in dict_IC:
                    dict_IC[tmp_ic] = dict()
                if tmp_upc not in dict_IC[tmp_ic]:
                    dict_IC[tmp_ic][tmp_upc] = dict()
                dict_IC[tmp_ic][tmp_upc]['desc'] = tmp_desc
                dict_IC[tmp_ic][tmp_upc]['csize'] = tmp_csize

                if tmp_upc not in dict_UPC:
                    dict_UPC[tmp_upc] = dict()
                if tmp_ic not in dict_UPC[tmp_upc]:
                    dict_UPC[tmp_upc][tmp_ic] = dict()
                dict_UPC[tmp_upc][tmp_ic]['desc'] = tmp_desc
                dict_UPC[tmp_upc][tmp_ic]['csize'] = tmp_csize

            idx1 += 1
        wb_r.close()
        dict_total = {
            'IC': dict_IC,
            'UPC': dict_UPC
        }

        if vc not in g_data['dict_vendor_item_code']:
            g_data['dict_vendor_item_code'][vc] = dict()
        for item2 in g_data['dict_store_code']:
            if sc_type == 'all':
                g_data['dict_vendor_item_code'][vc][item2] = dict_total
            elif sc_type == 'eq':
                if item2 in sc:
                    g_data['dict_vendor_item_code'][vc][item2] = dict_total
            else:
                if item2 not in sc:
                    g_data['dict_vendor_item_code'][vc][item2] = dict_total

    print('Vendor DB file is loaded')
