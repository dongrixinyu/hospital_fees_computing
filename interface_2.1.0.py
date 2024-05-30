# -*- coding=utf-8 -*-
# 更新内容
# 在 1.0.0 基础上，将所有的 “号”都算在内，但是多减一项卫生材料费。


import os
import traceback
import numpy as np
import pandas as pd
import tkinter as tk
import tkinter.filedialog


# 第一列过滤其它 骨科、眼科等 `门诊`表用
dermatology_medical_department_list = [
    '皮肤科专病诊室', '皮肤科诊室', '美容外科诊室', '美容皮肤科诊室']

# 第一列过滤其它 病区和科室 `住院`表用
dermatology_inpatient_ward_list = [
    '过敏性疾病科', '皮肤感染疾病科', '银屑病科', '大疱性疾病科']


def total_sum_judge_name(doctor_name):
    """ 判断病人医生是否计入 总额。共分几种情况
    `艾俊俊` `儿科专病李慧号` `过敏专病吴玲霞` `胡晓玲主任号` `激光便民号`

    与  judge_name 函数不同，该函数只要有医生名字，就算在内，所以，包含主任和副主任的都算。

    :param doctor_name:
    :return:
    """
    if '便民' in doctor_name:
        return False, doctor_name

    if '美容外科普通' in doctor_name:
        return False, doctor_name

    if '副主任' in doctor_name or '主任' in doctor_name:
        if '副主任' in doctor_name:
            return True, doctor_name.split('副主任')[0]
        elif '主任' in doctor_name:
            return True, doctor_name.split('主任')[0]

    if '号' not in doctor_name:  # and '病' not in doctor_name:
        if '专病' in doctor_name:
            return True, doctor_name.split('专病')[1]
        else:
            return True, doctor_name

    if '号' in doctor_name:
        if '专病' in doctor_name:
            doctor_name = doctor_name.split('专病')[1].split('号')[0]  # `专病`后，`号`前
            return True, doctor_name
            # return False, doctor_name
        else:
            doctor_name = doctor_name.split('号')[0]  # `专病`后，`号`前
            return True, doctor_name
            # return False, doctor_name


def judge_name(doctor_name):
    """ 判断病人医生是否计入 总额。共分几种情况
    `艾俊俊` `儿科专病李慧号` `过敏专病吴玲霞` `胡晓玲主任号` `激光便民号`

    :param doctor_name:
    :return:
    """
    if '未知' in doctor_name:
        return False, doctor_name

    if '科室小计' in doctor_name:
        return False, doctor_name

    if '0元惠民' in doctor_name:
        return False, doctor_name

    if '服务' in doctor_name:
        return False, doctor_name

    if '便民' in doctor_name:
        return False, doctor_name

    if '外聘专家' in doctor_name:
        return True, doctor_name.replace('外聘专家', '')

    if '副主任' in doctor_name or '主任' in doctor_name:
        return False, doctor_name

    if '号' not in doctor_name:  # and '病' not in doctor_name:
        if '专病' in doctor_name:
            return True, doctor_name.split('专病')[1]
        else:
            return True, doctor_name

    if '号' in doctor_name:
        if '专病' in doctor_name:
            doctor_name = doctor_name.split('专病')[1].split('号')[0]  # `专病`后，`号`前
            return True, doctor_name
            # return False, doctor_name
        else:
            doctor_name = doctor_name.split('号')[0]  # `专病`后，`号`前
            return True, doctor_name
            # return False, doctor_name


def combine_two_dict(dict_1, dict_2):
    new_dict = dict()
    for key_1, value_1 in dict_1.items():
        if key_1 in dict_2:
            new_dict.update({key_1: value_1 + dict_2[key_1]})
        else:
            new_dict.update({key_1: value_1})

    for key_2, value_2 in dict_2.items():
        if key_2 not in dict_1:
            new_dict.update({key_2: value_2})

    return new_dict


def find_all_receiving_department(xls_content, file_type):
    # file_type 指定文件属于 门诊还是住院
    receiving_department_index = find_key_item_index(xls_content, keywords='过滤执行科室')

    # 过滤科室，只保留皮肤大科的科室项目
    medical_department_index = find_key_item_index(xls_content, keywords='患者科室')

    receiving_department_list = []
    for line_idx in range(receiving_department_index[0] + 1, xls_content.values.shape[0]):
        res = xls_content.values[line_idx][receiving_department_index[1]]
        if type(res) is str:
            if file_type == '门诊':
                if xls_content.values[line_idx][medical_department_index[1]] in dermatology_medical_department_list:
                    # 检查科室
                    receiving_department_list.append(res)
            elif file_type == '住院':
                if xls_content.values[line_idx][medical_department_index[1]] in dermatology_inpatient_ward_list:
                    # 检查科室
                    receiving_department_list.append(res)

    return list(set(receiving_department_list))


def find_key_item_index(xls_content, keywords='病人医生'):
    xls_content_shape = xls_content.values.shape
    for i in range(xls_content_shape[0]):
        for j in range(xls_content_shape[1]):
            if xls_content.values[i][j] == keywords:
                return i, j

    raise ValueError('这个表格存在问题，无法找到 `{}` 单元。'.format(keywords))


def doctor_total_sum_statistics(
        xls_content, doctor_base_index, total_sum_index, file_type):
    """ 统计每个医生的所有总收入，一年统计一次

    :param xls_content:
    :param doctor_base_index: 医生索引号
    :param total_sum_index: `合计` 项索引号
    :param file_type: `住院` 或 `门诊`，用于判断哪些费用不减
    :return:
    """
    doctor_fee_info_dict = dict()
    # 根据检查检验项目，计算每个医生的总额
    tmp_doctor_fee_sum = 0
    cur_doctor_name = None

    for line_idx in range(doctor_base_index[0] + 1, xls_content.shape[0]):
        if xls_content.values[line_idx][0] in ['小计', '合计']:
            # 到了文件末尾，直接跳过
            continue

        doctor_name = xls_content.values[line_idx][doctor_base_index[1]]
        if type(doctor_name) is str:
            if line_idx == doctor_base_index[0] + 1:
                # 第一个医生，之前没有医生
                pass
            else:
                # 新一位医生计算开始，结束前一位医生总额。
                # 若前面有该医生，则叠加
                # 若前面没有出现过该医生，则直接更新

                # 判断该医生名是否应当计入
                is_a_doctor, cur_doctor_name = total_sum_judge_name(cur_doctor_name)
                if is_a_doctor:
                    if cur_doctor_name in doctor_fee_info_dict:
                        doctor_fee_info_dict[cur_doctor_name] += tmp_doctor_fee_sum
                    else:
                        doctor_fee_info_dict.update({cur_doctor_name: tmp_doctor_fee_sum})
                else:
                    print('###: ', cur_doctor_name)

                tmp_doctor_fee_sum = 0

            cur_doctor_name = doctor_name
            # print(cur_doctor_name)

            tmp_doctor_fee_sum = xls_content.values[line_idx][total_sum_index[1]]

        elif type(doctor_name) is float:  # 其含义为 nan 数据
            tmp_doctor_fee_sum += xls_content.values[line_idx][total_sum_index[1]]

        else:
            raise ValueError('表格中 `病人医生` 列存在问题。')

    # 将末尾的医生添加到信息中
    # 判断该医生名是否应当计入
    is_a_doctor, cur_doctor_name = total_sum_judge_name(cur_doctor_name)
    if is_a_doctor:
        if cur_doctor_name in doctor_fee_info_dict:
            doctor_fee_info_dict[cur_doctor_name] += tmp_doctor_fee_sum
        else:
            doctor_fee_info_dict.update({cur_doctor_name: tmp_doctor_fee_sum})
    else:
        print('###: ', cur_doctor_name)

    return dict(sorted(doctor_fee_info_dict.items(), key=lambda i: i[0]))


def class_statistics(
        xls_content, computing_class, medical_department_index,
        doctor_base_index, receiving_department,
        medical_income_total, sanitation_material_fee, consultation_fee,
        expensive_disposable_fee, file_type):
    """ 统计各个类型的计费

    :param xls_content:
    :param computing_class: 计费类型，如检查检验费，皮肤科门诊费等等。
    :param doctor_base_index:
    :param receiving_department:
    :param medical_income_total:
    :param sanitation_material_fee: 卫生检查费
    :param consultation_fee: 诊查费
    :param file_type: `住院` 或 `门诊`，用于判断哪些费用不减
    :return:
    """
    doctor_fee_info_dict = dict()
    # 根据检查检验项目，计算每个医生的总额
    tmp_doctor_fee_info = dict()
    cur_doctor_name = None
    for line_idx in range(doctor_base_index[0] + 1, xls_content.shape[0]):
        medical_department_name = xls_content.values[line_idx][medical_department_index[1]]
        # 检测大科室名字是否在皮肤科大类里
        if file_type == '门诊':
            if medical_department_name in dermatology_medical_department_list:
                pass
            else:
                continue
        elif file_type == '住院':
            if medical_department_name in dermatology_inpatient_ward_list:
                pass
            else:
                continue

        doctor_name = xls_content.values[line_idx][doctor_base_index[1]]
        if type(doctor_name) is str:
            is_a_doctor, cur_doctor_name = judge_name(doctor_name)
            if is_a_doctor:
                pass
                # print('医生姓名：{}'.format(cur_doctor_name))
            else:
                continue

            tmp_doctor_fee_info = dict()
            # 获取接收科室名，并判定是否在检查检验中
            receiving_department_name = xls_content.values[line_idx][receiving_department[1]]
            if receiving_department_name in computing_class:
                if file_type == '住院':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]] - \
                              xls_content.values[line_idx][expensive_disposable_fee[1]]
                    print('\t', receiving_department_name, line_idx, ':\t', tmp_fee)
                elif file_type == '门诊':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]] - \
                              xls_content.values[line_idx][expensive_disposable_fee[1]] - \
                              xls_content.values[line_idx][consultation_fee[1]]
                    print('\t', receiving_department_name, ':\t', tmp_fee)
                if tmp_fee != 0:
                    tmp_doctor_fee_info.update({receiving_department_name: tmp_fee})

            if cur_doctor_name in doctor_fee_info_dict:
                # print('医生姓名：', cur_doctor_name)
                doctor_fee_info_dict[cur_doctor_name].update(
                    combine_two_dict(doctor_fee_info_dict[cur_doctor_name],
                                     tmp_doctor_fee_info))
            else:
                # print('医生姓名：', cur_doctor_name)
                doctor_fee_info_dict.update({cur_doctor_name: tmp_doctor_fee_info})


    print()

    return doctor_fee_info_dict


class BasePage(object):
    def __init__(self, root):
        self.root = root
        # self.root.config()
        self.root.title('长治市第二人民医院皮肤科医生绩效统计')
        self.root.geometry('1000x618')

        # scrollbar = tk.Scrollbar(self.root)
        # scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.data_obj = DataStatistics()

        SelectFilePage(self.root, self.data_obj)


class DataStatistics(object):
    def __init__(self):
        self.file_type = None  # `住院` 或者 `门诊`
        self.xls_file_path = None
        self.xls_content = None
        self.receiving_department_list = []
        self.inspection_testing_list = [
            '超声医学科', '内镜室', '128层CT室', 'CT室', '1.5T磁共振室', 'DR室',
            '彩超室', '医学影像科', '心电图室', '磁共振室', '医学检验科']
        self.dermatology_clinic_list = [
            '皮肤科药浴室', '皮肤科治疗室', '皮肤白癜风治疗室', '皮肤痤疮治疗室',
            '皮肤科308照射室', '皮肤科二氧化碳激光室', '皮肤科光疗室',
            '皮肤科中医治疗室',
            '皮肤科果酸换肤室', '皮肤科He-Ni激光室', '皮肤科红蓝光治疗室',
            '皮肤科皮肤镜检查室', '皮肤科冷疗室', '皮肤科检查室', '皮肤科氦氖激光室',
            '皮肤科冷冻室', '皮肤中医治疗室', '皮肤科门诊', '皮肤科外擦药室']

        # self.recovery_clinic_list = ['康复医学科病区', '康复医学科']
        self.cosmetic_surgery_list = ['美容外科诊室']
        self.cosmetic_dermatology_list = ['美容皮肤科诊室']
        # self.special_need_list = ['特需病区']
        self.dermatology_ward_list = [
            '过敏性疾病科病区', '皮肤感染疾病科病区', '银屑病病区', '大疱性疾病病区',
            '特需病区', '五官科病区', '妇科病区', '康复医学科病区', '老年病病区', '儿科病区'
        ]

    def get_xls_content(self, xls_file_path):
        """ 获取 xls 文件的 pandas 实例

        :param xls_file_path:
        :return:
        """
        self.xls_content = pd.read_excel(io=xls_file_path)

        def compensate_first_column(xls_content):
            # 补全第一列的`患者科室`
            medical_department_index = find_key_item_index(xls_content, keywords='患者科室')
            temp_medical_department_name = '患者科室'
            for i in range(medical_department_index[0], xls_content.shape[0]):
                if xls_content.values[i][medical_department_index[1]] is np.nan:
                    xls_content.values[i][medical_department_index[1]] = temp_medical_department_name
                else:
                    temp_medical_department_name = xls_content.values[i][medical_department_index[1]]

        def compensate_second_column(xls_content):
            # 补全第二列 `患者医生`
            doctor_index = find_key_item_index(xls_content, keywords='患者医生')
            temp_doctor_name = '患者医生'
            for i in range(doctor_index[0], xls_content.shape[0]):
                if xls_content.values[i][doctor_index[1]] is np.nan:
                    xls_content.values[i][doctor_index[1]] = temp_doctor_name
                else:
                    temp_doctor_name = xls_content.values[i][doctor_index[1]]

        compensate_first_column(self.xls_content)
        compensate_second_column(self.xls_content)

    def get_receiving_department_list(self):
        self.receiving_department_list = sorted(
            find_all_receiving_department(self.xls_content, self.file_type))


class SelectFilePage(object):
    def __init__(self, root, data_obj):
        self.root = root
        self.data_obj = data_obj

        self.interface = tk.Frame(self.root)
        self.interface.pack()

        title_label = tk.Label(
            self.interface, text='1. 选择 xls 文件，并指定该文件是 住院|门诊',
            font='Helvetica 13 bold')
        title_label.grid(row=0, column=0, columnspan=3)

        label = tk.Label(self.interface, text='请打开xls文件：')
        label.grid(row=1, column=0)

        self.entry = tk.Entry(self.interface, bd=5, width=40)
        self.entry.grid(row=1, column=1)

        button = tk.Button(self.interface, text="打开", command=self.select_statistics_file)
        button.grid(row=1, column=2)

        label = tk.Label(self.interface, text='请指定文件类型：')
        label.grid(row=2, column=0)

        self.menu_button = tk.Menubutton(self.interface, text='住院/门诊', relief=tk.RAISED)
        self.menu_button.grid(row=2, column=1)

        self.menu_button.menu = tk.Menu(self.menu_button, tearoff=0)
        self.menu_button['menu'] = self.menu_button.menu

        self.inpatient_var = tk.IntVar()
        self.outpatient_var = tk.IntVar()

        self.menu_button.menu.add_checkbutton(label='住院', variable=self.inpatient_var)
        self.menu_button.menu.add_checkbutton(label='门诊', variable=self.outpatient_var)

        next_page_button = tk.Button(self.interface, text='下一步', command=self.next_page)
        next_page_button.grid(row=3, column=3)

    def next_page(self):
        if self.inpatient_var.get() == 1 and self.outpatient_var.get() == 1:
            return
        elif self.inpatient_var.get() == 1 and self.outpatient_var.get() == 0:
            self.data_obj.file_type = '住院'
        elif self.inpatient_var.get() == 0 and self.outpatient_var.get() == 1:
            self.data_obj.file_type = '门诊'
        elif self.inpatient_var.get() == 0 and self.outpatient_var.get() == 0:
            return

        if self.data_obj.xls_content is None:
            pass
        else:
            self.interface.destroy()
            DefinitionPage(self.root, self.data_obj)

    def select_statistics_file(self):
        statistics_file_name = tk.filedialog.askopenfilename()
        print(statistics_file_name)
        if statistics_file_name != '':

            try:
                self.data_obj.get_xls_content(statistics_file_name)
                self.data_obj.xls_file_path = statistics_file_name
                self.entry.insert(0, statistics_file_name)
            except:
                print(traceback.print_exc())
                print('输入的文件不是 xls 格式。')
                self.entry.insert(0, '输入的文件不是 xls 格式')

        else:
            self.entry.insert(0, '您没有选择任何文件')

    def select_file_type_inpatient(self):
        self.data_obj.file_type = '住院'
        self.menu_button.config(text='住院')

    def select_file_type_outpatient(self):
        self.data_obj.file_type = '门诊'
        self.menu_button.config(text='门诊')


class DefinitionPage(object):
    """ 定义各种统计项

    """
    def __init__(self, root, data_obj):
        self.root = root
        self.data_obj = data_obj
        # self.root.config()

        self.scrollbar = tk.Scrollbar(self.root)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=tk.FALSE)

        self.interface = tk.Canvas(self.root, yscrollcommand=self.scrollbar.set)
        self.interface.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE)

        self.scrollbar.config(command=self.interface.yview)

        # reset the view
        self.interface.xview_moveto(0)
        self.interface.yview_moveto(0)

        # create a frame inside the canvas which will be scrolled with it
        self.interior = tk.Frame(self.interface)
        self.interior_id = self.interface.create_window(
            0, 0, window=self.interior, anchor=tk.NW)
        self.interior.bind('<Configure>', self._configure_interior)
        self.interface.bind('<Configure>', self._configure_canvas)

        horizontal, vertical = 0, 0
        title_label = tk.Label(
            self.interior, text='二. 确定类别、子项', font='Helvetica 13 bold')
        title_label.grid(row=horizontal, column=vertical)

        self.inspection_testing_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.inspection_testing_list, horizontal, vertical,
            sub_title='1. 检查检验费合计', bg_color='light blue')
        horizontal += 1

        self.dermatology_clinic_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_clinic_list, horizontal, vertical,
            sub_title='2. 皮肤科门诊费合计', bg_color='red')
        horizontal += 1

        # self.recovery_clinic_check_buttons, horizontal = self._get_sub_item_info(
        #     self.data_obj.recovery_clinic_list, horizontal, vertical,
        #     sub_title='3. 康复类合计', bg_color='light yellow')
        # horizontal += 1

        self.cosmetic_surgery_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.cosmetic_surgery_list, horizontal, vertical,
            sub_title='3. 美容外科费用', bg_color='light green')
        horizontal += 1

        self.cosmetic_dermatology_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.cosmetic_dermatology_list, horizontal, vertical,
            sub_title='4. 美容皮肤科费用', bg_color='gray')
        horizontal += 1

        # self.special_need_check_buttons, horizontal = self._get_sub_item_info(
        #     self.data_obj.special_need_list, horizontal, vertical,
        #     sub_title='6. 特需费用', bg_color='pink')
        # horizontal += 1

        self.dermatology_ward_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_ward_list, horizontal, vertical,
            sub_title='5. 皮肤科病区', bg_color='light blue')
        horizontal += 1

        # self.dermatology_ward_2_check_buttons, horizontal = self._get_sub_item_info(
        #     self.data_obj.dermatology_ward_2_list, horizontal, vertical,
        #     sub_title='8. 皮肤科（二病区）', bg_color='green')
        # horizontal += 1
        #
        # self.dermatology_ward_3_check_buttons, horizontal = self._get_sub_item_info(
        #     self.data_obj.dermatology_ward_3_list, horizontal, vertical,
        #     sub_title='9. 皮肤科（三病区）', bg_color='red')
        # horizontal += 1
        #
        # self.dermatology_ward_4_check_buttons, horizontal = self._get_sub_item_info(
        #     self.data_obj.dermatology_ward_4_list, horizontal, vertical,
        #     sub_title='10. 皮肤科（四病区）', bg_color='yellow')
        # horizontal += 1

        next_page_button = tk.Button(self.interior, text='下一步', command=self.next_page)
        next_page_button.grid(row=horizontal, column=0)

    # track changes to the canvas and frame width and sync them,
    # also updating the scrollbar
    def _configure_interior(self, event):
        # update the scrollbars to match the size of the inner frame
        size = (self.interior.winfo_reqwidth(), self.interior.winfo_reqheight())
        self.interface.config(scrollregion="0 0 %s %s" % size)
        if self.interior.winfo_reqwidth() != self.interface.winfo_width():
            # update the canvas's width to fit the inner frame
            self.interface.config(width=self.interior.winfo_reqwidth())

    def _configure_canvas(self, event):
        if self.interior.winfo_reqwidth() != self.interface.winfo_width():
            # update the inner frame's width to fill the canvas
            self.interface.itemconfigure(self.interior_id, width=self.interface.winfo_width())

    def _get_sub_item_info(self, class_items, horizontal, vertical,
                           sub_title='1. 检查检验费合计', bg_color='light blue'):
        """ 获取每个大类的所有可选子项

        :param class_items: 大类别默认子选项
        :param horizontal:
        :param vertical:
        :return:
        """
        horizontal += 1
        subtitle_label = tk.Label(
            self.interior, text=sub_title, font='Helvetica 11 bold')
        subtitle_label.grid(row=horizontal, column=vertical)

        self.data_obj.get_receiving_department_list()

        check_button_list = list()
        horizontal += 1
        for item in self.data_obj.receiving_department_list:
            check_var = tk.IntVar()
            if item in class_items:
                cb = tk.Checkbutton(self.interior, text=item, variable=check_var,
                                    bg=bg_color, height=1, width=15)
            else:
                cb = tk.Checkbutton(self.interior, text=item, variable=check_var,
                                    height=1, width=15)

            cb.grid(row=horizontal, column=vertical)
            check_button_list.append(check_var)
            if vertical == 5:
                horizontal += 1
                vertical = 0
            else:
                vertical += 1

        return check_button_list, horizontal

    def next_page(self):
        # 获取所有的子项目
        self.data_obj.get_receiving_department_list()
        self.data_obj.inspection_testing_list = [
            text for button, text in zip(
                self.inspection_testing_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]
        print(self.data_obj.inspection_testing_list)

        self.data_obj.dermatology_clinic_list = [
            text for button, text in zip(
                self.dermatology_clinic_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        # self.data_obj.recovery_clinic_list = [
        #     text for button, text in zip(
        #         self.recovery_clinic_check_buttons,
        #         self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.cosmetic_surgery_list = [
            text for button, text in zip(
                self.cosmetic_surgery_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.cosmetic_dermatology_list = [
            text for button, text in zip(
                self.cosmetic_dermatology_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        # self.data_obj.special_need_list = [
        #     text for button, text in zip(
        #         self.special_need_check_buttons,
        #         self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.dermatology_ward_list = [
            text for button, text in zip(
                self.dermatology_ward_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        # self.data_obj.dermatology_ward_2_list = [
        #     text for button, text in zip(
        #         self.dermatology_ward_2_check_buttons,
        #         self.data_obj.receiving_department_list) if button.get() == 1]
        #
        # self.data_obj.dermatology_ward_3_list = [
        #     text for button, text in zip(
        #         self.dermatology_ward_3_check_buttons,
        #         self.data_obj.receiving_department_list) if button.get() == 1]
        #
        # self.data_obj.dermatology_ward_4_list = [
        #     text for button, text in zip(
        #         self.dermatology_ward_4_check_buttons,
        #         self.data_obj.receiving_department_list) if button.get() == 1]

        self.interface.destroy()
        DetailPage(self.root, self.data_obj)


class DetailPage(object):
    def __init__(self, root, data_obj):
        self.root = root
        self.root.config()

        self.data_obj = data_obj

        self.interface = tk.Frame(self.root, )
        self.interface.pack()

        title_label = tk.Label(self.interface, text='三. 计算结果如下', font='Helvetica 13 bold')
        title_label.grid(row=0, column=0)

        # 用于计算合计
        total_sum_index = find_key_item_index(self.data_obj.xls_content, keywords='合计')

        medical_department_index = find_key_item_index(self.data_obj.xls_content, keywords='患者科室')
        doctor_base_index = find_key_item_index(self.data_obj.xls_content, keywords='患者医生')
        receiving_department = find_key_item_index(self.data_obj.xls_content, keywords='过滤执行科室')
        medical_income_total = find_key_item_index(self.data_obj.xls_content, keywords='合计')
        sanitation_material_fee = find_key_item_index(self.data_obj.xls_content, keywords='卫生材料费')
        consultation_fee = find_key_item_index(self.data_obj.xls_content, keywords='诊查费')
        expensive_disposable_fee = find_key_item_index(self.data_obj.xls_content, keywords='高值耗材费')

        # self.docker_total_sum_dict = doctor_total_sum_statistics(
        #     self.data_obj.xls_content, doctor_base_index, total_sum_index, self.data_obj.file_type)

        self.inspection_testing_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.inspection_testing_list,
            medical_department_index,
            doctor_base_index, receiving_department,
            medical_income_total, sanitation_material_fee, consultation_fee, expensive_disposable_fee,
            self.data_obj.file_type)

        self.dermatology_clinic_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_clinic_list,
            medical_department_index,
            doctor_base_index, receiving_department,
            medical_income_total, sanitation_material_fee, consultation_fee, expensive_disposable_fee,
            self.data_obj.file_type)

        self.cosmetic_surgery_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.cosmetic_surgery_list,
            medical_department_index, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee, expensive_disposable_fee,
            self.data_obj.file_type)

        self.cosmetic_dermatology_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.cosmetic_dermatology_list,
            medical_department_index, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee, expensive_disposable_fee,
            self.data_obj.file_type)

        self.dermatology_ward_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_ward_list,
            medical_department_index, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee, expensive_disposable_fee,
            self.data_obj.file_type)

        text_canvas = tk.Text(self.interface, height=38, width=120)

        # csv_total_sum_fee_path = self.print_and_write_total_sum_file()
        csv_total_sum_fee_path = ''

        print_detail_text, csv_file_path = self.print_and_write_file()
        text_canvas.grid(row=1, column=0, columnspan=4)
        text_canvas.insert(tk.END, print_detail_text)

        text_path_canvas_1 = tk.Text(self.interface, height=2, width=120)
        text_path_canvas_1.grid(row=2, column=0, columnspan=4)
        text_path_canvas_1.insert(tk.END, csv_file_path)

        text_path_canvas_2 = tk.Text(self.interface, height=2, width=120)
        text_path_canvas_2.grid(row=3, column=0, columnspan=4)
        text_path_canvas_2.insert(tk.END, csv_total_sum_fee_path)

    def print_and_write_total_sum_file(self):

        doctor_total_sum_fee = []
        for doctor_name, doctor_sum in self.docker_total_sum_dict.items():
            doctor_total_sum_fee.append('{},{}'.format(doctor_name, doctor_sum))

        doctor_fee_txt = '\n'.join(doctor_total_sum_fee)
        with open(os.path.join(DIR_PATH, self.data_obj.file_type + '_医生月度总计.csv'), 'w', encoding='utf-8') as fw:
            fw.write(doctor_fee_txt)

        return os.path.join(DIR_PATH, self.data_obj.file_type + '_医生月度总计.csv')

    def print_and_write_file(self):
        doctor_fees_dict = dict()
        item_list = ['检查检验合计', '皮肤科门诊合计', '美容外科诊室', '美容皮肤科诊室', '皮肤科病区']
        doctor_fees_txt = '医生姓名,' + ','.join(item_list) + '\n'

        num_format = '{:.2f}'
        print_text = [self.data_obj.file_type]
        for doctor_name in self.inspection_testing_info_dict:
            print_text.append(doctor_name + ' ' + self.data_obj.file_type)
            tmp_dict = {
                '检查检验合计': sum(list(self.inspection_testing_info_dict[doctor_name].values())),
                '皮肤科门诊合计': sum(list(self.dermatology_clinic_info_dict[doctor_name].values())),
                '美容外科诊室': sum(list(self.cosmetic_surgery_info_dict[doctor_name].values())),
                '美容皮肤科诊室': sum(list(self.cosmetic_dermatology_info_dict[doctor_name].values())),
                '皮肤科病区': sum(list(self.dermatology_ward_info_dict[doctor_name].values()))
            }

            if len(self.inspection_testing_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '检查检验合计: ' + num_format.format(
                    sum(list(self.inspection_testing_info_dict[doctor_name].values()))))
                for key, val in self.inspection_testing_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.dermatology_clinic_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科门诊合计: ' + num_format.format(
                    sum(list(self.dermatology_clinic_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_clinic_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))

            if len(self.cosmetic_surgery_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '美容外科诊室: ' + num_format.format(
                    sum(list(self.cosmetic_surgery_info_dict[doctor_name].values()))))
                for key, val in self.cosmetic_surgery_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.cosmetic_dermatology_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '美容皮肤科诊室: ' + num_format.format(
                    sum(list(self.cosmetic_dermatology_info_dict[doctor_name].values()))))
                for key, val in self.cosmetic_dermatology_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))

            if len(self.dermatology_ward_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科病区: ' + num_format.format(
                    sum(list(self.dermatology_ward_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_ward_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))

            print_text.append('\n')
            doctor_fees_txt += doctor_name + ',' + ','.join(map(str, list(tmp_dict.values()))) + '\n'
            doctor_fees_dict.update({doctor_name: tmp_dict})

        with open(os.path.join(DIR_PATH, self.data_obj.file_type + '_医生各项明细.csv'), 'w', encoding='utf-8') as fw:
            fw.write(doctor_fees_txt)

        return '\n'.join(print_text), os.path.join(DIR_PATH, self.data_obj.file_type + '_医生各项明细.csv')


if __name__ == '__main__':
    DIR_PATH = os.path.dirname(os.path.abspath(__file__))

    root = tk.Tk()

    BasePage(root)
    # 进入消息循环
    root.mainloop()


