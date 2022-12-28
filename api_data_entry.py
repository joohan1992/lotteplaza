from flask import Flask, request, redirect, url_for, flash, render_template, session, send_file
from flask_cors import CORS, cross_origin
from flask_caching import Cache
from werkzeug.utils import secure_filename
from os.path import isdir, join, isfile
from os import mkdir, remove, listdir

# import lib_api_data_entry as de_lib
# import lib_api_data_entry_vendor_data as de_lib
# import lib_api_data_entry_case_each as de_lib
# import lib_api_data_entry_all_prefix as de_lib
# import lib_api_data_entry_shift_f26 as de_lib
import lib_api_data_entry_store_data_clean as de_lib
import lib_api_data_entry_load as de_lib_load


g_data = {'base_address': 'http://211.58.242.46:8082',
          'fname_dict_vendor': '',
          'status_dict_vendor': 0,
          'fname_dict_store': '',
          'fname_dict_vend': '',
          'status_step_two': 0,
          'dict_vendor': None,
          'dict_store': None,
          'dict_vend': None,
          'list_step_two': {},
          'list_step_four': {},
          'fname_base_data': '120922 INVOICE FORMAT.xlsx',
          # 업체 데이터 활용 시 기본 설정값
          # - name: 파일명
          # - sheet: 시트명
          # - vc: 업체코드
          # - sc_type: 점포코드 참조 방식(all=모든점포, eq=동일한코드, ne=아닌코드)
          # - sc: 점포코드(리스트)
          # - pos: 시트 내 시작 열, 제품코드, UPC, Description, Case Size의 위치 (모든 순서는 0부터 시작)
          # - prefix: 데이터 읽어들일 시 업체별 아이템 prefix 적용 여부 True/False(Boolean)
          'conf_vendor_item_code': {
              'SV_FL01': {'name': '112219 FL01 SV MP.xlsx', 'sheet': 'SV',
                          'vc': '1229', 'sc_type': 'eq', 'sc': ['011'], 'pos': (2, 0, 1, 3, 4), 'prefix': True},
              'SV_ETC': {'name': '102819 SV MP.xlsx', 'sheet': 'SV',
                         'vc': '1229', 'sc_type': 'ne', 'sc': ['011'], 'pos': (2, 0, 1, 3, 4), 'prefix': True},
              'JFC': {'name': '112919 JFC.xlsx', 'sheet': 'SQL Results',
                      'vc': '0202', 'sc_type': 'all', 'sc': [], 'pos': (1, 1, 2, 3, 4), 'prefix': True},
              'CJ': {'name': '113019 CJ.xlsx', 'sheet': 'CJ ORDER BOOK',
                     'vc': '0106', 'sc_type': 'all', 'sc': [], 'pos': (4, 0, 1, 3, 6), 'prefix': True},
              'ENI': {'name': '113019 ENI ITEM LIST.xlsx', 'sheet': 'Sheet1',
                      'vc': '0179', 'sc_type': 'all', 'sc': [], 'pos': (2, 0, 1, 2, 5), 'prefix': False},
              'PACIFIC BLUE': {'name': '113019 PACIFIC BLUE CORP ITEM LIST.xlsx', 'sheet': 'ITEM LIST',
                               'vc': '0806', 'sc_type': 'all', 'sc': [], 'pos': (1, 1, 2, 3, 5), 'prefix': True},
              'K&P': {'name': '120419 KPI ITEM LIST.xlsx', 'sheet': 'KPI',
                      'vc': '0118', 'sc_type': 'all', 'sc': [], 'pos': (10, 1, 2, 3, 5), 'prefix': False},
              'OTG': {'name': '120419 OTG ITEM LIST.xlsx', 'sheet': 'OTTOGI USA (11-26-2019)',
                      'vc': '0137', 'sc_type': 'all', 'sc': [], 'pos': (6, 0, 3, 2, 6), 'prefix': False},
              'WON': {'name': '120419 WON TRADING ITEM LIST.xlsx', 'sheet': 'PULMUONE',
                      'vc': '0151', 'sc_type': 'all', 'sc': [], 'pos': (9, 0, 1, 2, 6), 'prefix': True},
              'CLV': {'name': '121019 CLOVERLAND ITEM LIST.xlsx', 'sheet': 'BuyGrp_Item_Prices__20191208114',
                      'vc': '1205', 'sc_type': 'all', 'sc': [], 'pos': (5, 0, 1, 2, 4), 'prefix': False},
              'GOYA': {'name': '120519 GOYA ITEM LIST.xlsx', 'sheet': '120519 GOYA',
                      'vc': '1005', 'sc_type': 'all', 'sc': [], 'pos': (4, 1, 2, 3, 5), 'prefix': True}
          },
          'dict_vendor_item_code': None,
          'dict_store_code': ['001', '002', '003', '004', '005', '006', '007', '008', '009', '010', '011', '012', '013']
}


app = Flask(__name__)
CORS(app, support_credentials=True)
cache = Cache(app, config={'CACHE_TYPE': 'simple'})


@app.route('/', methods=['GET', 'POST'])
def index():
    global g_data
    return render_template('main.html', g_data=g_data)


@app.route('/change_base_data', methods=['GET', 'POST'])
def change_base_data():
    global g_data
    return render_template('change_base_data.html', g_data=g_data)


@app.route('/dict_vendor_upload', methods=['GET', 'POST'])
def dict_vendor_upload():
    global g_data
    if request.method == 'POST' and g_data['status_dict_vendor'] in [0, 9]:
        f = request.files['file_dict_vendor']
        final_fname = secure_filename(f.filename)
        g_data['fname_dict_vendor'] = final_fname
        g_data['status_dict_vendor'] = 1
        f.save('./org_store_data/'+final_fname)

        de_lib.convert_dict_vendor(g_data)

        return redirect('/change_base_data')
    else:
        flash('현재 다른 매장데이터 파일이 업로드 중입니다.')
        return redirect('/change_base_data')


@app.route('/step_two/<error>', methods=['GET', 'POST'])
def step_two(error):
    global g_data
    if error is None:
        error = ''
    list_file = listdir('./result_step_2')
    result_list = []
    for f in list_file:
        if f.split('.')[-1] == 'xlsx':
            err_flag = 0
            err_text = ''
            if isfile('./error/'+f.rsplit('.', 1)[0]+'.txt'):
                err_file = open('./error/'+f.rsplit('.', 1)[0]+'.txt', 'r')
                err_flag = 1
                err_text = ''
                list_error = err_file.readlines()
                for item_error in list_error:
                    err_text += item_error
                err_file.close()
                err_text = err_text.replace('\n', '\\n')

            list_input_file = listdir('./input_step_2/'+f.rsplit('.', 1)[0])

            result_list.append({
                'file_name': f,
                'input_cnt': len(list_input_file),
                'err_flag': err_flag,
                'err_text': err_text
            })
    return render_template('step_two.html', g_data=g_data, result_list=result_list, error=error)


@app.route('/step_four/<error>', methods=['GET', 'POST'])
def step_four(error):
    global g_data
    if error is None:
        error = ''
    list_file = listdir('./result_step_4')
    result_list = []
    for f in list_file:
        err_flag = 0
        err_text = ''
        if isfile('./error_step_4/'+f.rsplit('.', 1)[0]+'.txt'):
            err_file = open('./error_step_4/'+f.rsplit('.', 1)[0]+'.txt', 'r')
            err_flag = 1
            err_text = ''
            list_error = err_file.readlines()
            for item_error in list_error:
                err_text += item_error
            err_file.close()
            err_text = err_text.replace('\n', '\\n')
        result_list.append({
            'file_name': f,
            'err_flag': err_flag,
            'err_text': err_text
        })
    return render_template('step_four.html', g_data=g_data, result_list=result_list, error=error)


@app.route('/step_five/<error>', methods=['GET', 'POST'])
def step_five(error):
    global g_data
    if error is None:
        error = ''
    list_file = listdir('./result_step_5')
    result_list = []
    for f in list_file:
        err_flag = 0
        err_text = ''
        if isfile('./error_step_5/'+f.rsplit('.', 1)[0]+'.txt'):
            err_file = open('./error_step_5/'+f.rsplit('.', 1)[0]+'.txt', 'r')
            err_flag = 1
            err_text = ''
            list_error = err_file.readlines()
            for item_error in list_error:
                err_text += item_error
            err_file.close()
            err_text = err_text.replace('\n', '\\n')
        result_list.append({
            'file_name': f,
            'err_flag': err_flag,
            'err_text': err_text
        })
    return render_template('step_five.html', g_data=g_data, result_list=result_list, error=error)


@app.route('/reference', methods=['GET', 'POST'])
def reference():
    global g_data
    param = {
        ('입력 포맷', 'root', '첨부 1. 1단계 입력 폼 CS'),
        ('업체/점포 코드 테이블', 'root', '첨부 2, 3 통합본_200423')
    }

    return render_template('reference.html', g_data=g_data, param=param)


@app.route('/down_file/<data_type>/<file_name>', methods=['GET', 'POST'])
def down_file(data_type, file_name):
    base_dir = ''
    print(data_type)
    print(file_name)
    if data_type == 'result_step_2':
        base_dir = './result_step_2'
    elif data_type == 'result_step_4':
        base_dir = './result_step_4'
    elif data_type == 'result_step_5':
        base_dir = './result_step_5'
    elif data_type == 'result_step_2_error':
        base_dir = './error'
        file_name = file_name.rsplit('.', 1)[0]+'.txt'
    elif data_type == 'result_step_4_error':
        base_dir = './error_step_4'
        file_name = file_name.rsplit('.', 1)[0]+'.txt'
    elif data_type == 'result_step_5_error':
        base_dir = './error_step_5'
        file_name = file_name.rsplit('.', 1)[0]+'.txt'
    elif data_type == 'store_data':
        base_dir = './processed_store_data'
        file_name = file_name.rsplit('.', 1)[0]+'.xlsx'
    elif data_type == 'root':
        base_dir = './'
        file_name = file_name.rsplit('.', 1)[0]+'.xlsx'
        file_name = file_name.replace('%5BPOINT%5D', '.').replace('[POINT]', '.')
    response = send_file(join(base_dir, file_name), attachment_filename=file_name, as_attachment=True)
    return response


@app.route('/file_input_upload', methods=['GET', 'POST'])
def file_input_upload():
    global g_data
    if g_data['status_dict_vendor'] != 9:
        return redirect('/step_two/nodv')
    file_result_name = secure_filename(request.form['file_result_name'])
    files = request.files.getlist('file_input')
    if len(files) > 0:
        if not isdir(app.config['UPLOAD_FOLDER']+'/'+file_result_name):
            mkdir(app.config['UPLOAD_FOLDER']+'/'+file_result_name)
        elif len(listdir(app.config['UPLOAD_FOLDER']+'/'+file_result_name)):
            for file_to_rm in listdir(app.config['UPLOAD_FOLDER']+'/'+file_result_name):
                remove(join(app.config['UPLOAD_FOLDER']+'/'+file_result_name, file_to_rm))
        if isfile('./result_step_2/'+file_result_name+'.txt'):
            remove('./result_step_2/'+file_result_name+'.txt')
        for f in files:
            f.save(join(app.config['UPLOAD_FOLDER']+'/'+file_result_name, f.filename))
        g_data['list_step_two'][file_result_name] = 0
        print('including '+str(files[0].filename)+', '+str(len(files))+' files were uploaded!')
        return redirect('/step_two_process')
    return redirect('/step_two/no')


@app.route('/file_input_upload_step_4', methods=['GET', 'POST'])
def file_input_upload_step_4():
    global g_data
    if g_data['status_dict_vendor'] != 9:
        return redirect('/step_four/nodv')
    file_result_name = secure_filename(request.form['file_result_name'])
    files = request.files.getlist('file_input')
    if len(files) > 0:
        if not isdir(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name):
            mkdir(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name)
        elif len(listdir(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name)):
            for file_to_rm in listdir(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name):
                remove(join(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name, file_to_rm))
        for f in files:
            f.save(join(app.config['UPLOAD_FOLDER_STEP_FOUR']+'/'+file_result_name, f.filename))
        g_data['list_step_four'][file_result_name] = 0
        return redirect('/step_four_process')
    return redirect('/step_four/no')


@app.route('/step_two_process', methods=['GET', 'POST'])
def step_two_process():
    global g_data
    de_lib.scheduler_step_two(g_data)
    return redirect('/step_two/no')


@app.route('/step_four_process', methods=['GET', 'POST'])
def step_four_process():
    global g_data
    de_lib.scheduler_step_four(g_data)
    return redirect('/step_four/no')


@app.route('/practice', methods=['GET', 'POST'])
def practice():
    global g_data
    return render_template('practice.html', g_data=g_data)


if __name__ == '__main__':
    port_num = g_data['base_address'].rsplit(':', 1)[1]
    print(f'Loading MARGIN_DIFF_FLAG {de_lib.flag_margin_diff} on {port_num}')
    de_lib.read_config(g_data)
    de_lib_load.read_vendor_code(g_data)
    app.secret_key = "super secret key"
    app.config['SESSION_TYPE'] = 'filesystem'
    app.config['UPLOAD_FOLDER'] = 'C:/Users/user/Documents/GitHub/lotteplaza/input_step_2'
    app.config['UPLOAD_FOLDER_STEP_FOUR'] = 'C:/Users/user/Documents/GitHub/lotteplaza/input_step_4'
    app.config['UPLOAD_FOLDER_STEP_FIVE'] = 'C:/Users/user/Documents/GitHub/lotteplaza/input_step_5'
    cache.init_app(app)
    base_dir_list = ['org_store_data', 'processed_store_data', 'base_file', 'input_step_2', 'result_step_2', 'error', 'input_step_4', 'result_step_4', 'srp_result']
    for base_dir_unit in base_dir_list:
        if not isdir('./'+base_dir_unit):
            mkdir('./'+base_dir_unit)
    app.run(host='0.0.0.0', port=int(port_num))
