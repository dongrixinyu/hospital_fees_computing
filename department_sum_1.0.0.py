﻿# -*- coding=utf-8 -*-
# 计算各个科室的金额总和
# 1.0.0


import os
import pandas as pd
import tkinter as tk
import tkinter.filedialog


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


def find_all_receiving_department(xls_content):
    receiving_department_index = find_key_item_index(xls_content, keywords='接收科室')

    receiving_department_list = list()
    for line_idx in range(receiving_department_index[0] + 1, xls_content.values.shape[0]):
        res = xls_content.values[line_idx][receiving_department_index[1]]
        if type(res) is str:
            receiving_department_list.append(res)

    return list(set(receiving_department_list))


def find_key_item_index(xls_content, keywords='病人医生'):
    xls_content_shape = xls_content.values.shape
    for i in range(xls_content_shape[0]):
        for j in range(xls_content_shape[1]):
            if xls_content.values[i][j] == keywords:
                return i, j
    raise ValueError('这个表格存在问题，无法找到 `{}` 单元。'.format(keywords))


def class_statistics(xls_content, computing_class, doctor_base_index, receiving_department,
                     medical_income_total, sanitation_material_fee, consultation_fee,
                     file_type):
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
        doctor_name = xls_content.values[line_idx][doctor_base_index[1]]
        if type(doctor_name) is str:
            if line_idx == doctor_base_index[0] + 1:
                pass
            else:
                # 新一位医生计算开始，结束前一位医生总额。
                # 若前面有该医生，则叠加
                # 若前面没有该医生，则直接更新

                # 判断该医生名是否应当计入
                is_a_doctor, cur_doctor_name = judge_name(cur_doctor_name)
                if is_a_doctor:
                    if cur_doctor_name in doctor_fee_info_dict:
                        doctor_fee_info_dict[cur_doctor_name].update(
                            combine_two_dict(doctor_fee_info_dict[cur_doctor_name],
                                             tmp_doctor_fee_info))
                    else:
                        doctor_fee_info_dict.update({cur_doctor_name: tmp_doctor_fee_info})
                else:
                    print('###: ', cur_doctor_name)
                tmp_doctor_fee_info = dict()

            cur_doctor_name = doctor_name
            print(cur_doctor_name)

            # 获取接收科室名，并判定是否在检查检验中
            receiving_department_name = xls_content.values[line_idx][receiving_department[1]]
            if receiving_department_name in computing_class:
                if file_type == '住院':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]]
                elif file_type == '门诊':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]] - \
                              xls_content.values[line_idx][consultation_fee[1]]
                print('\t', receiving_department_name, ':\t', tmp_fee)
                tmp_doctor_fee_info.update({receiving_department_name: tmp_fee})

        elif type(doctor_name) is float:
            # 获取接收科室名，并判定是否在检查检验中
            receiving_department_name = xls_content.values[line_idx][receiving_department[1]]
            if receiving_department_name in computing_class:
                if file_type == '住院':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]]
                elif file_type == '门诊':
                    tmp_fee = xls_content.values[line_idx][medical_income_total[1]] - \
                              xls_content.values[line_idx][sanitation_material_fee[1]] - \
                              xls_content.values[line_idx][consultation_fee[1]]
                print('\t', receiving_department_name, ':\t', tmp_fee)
                tmp_doctor_fee_info.update({receiving_department_name: tmp_fee})

        else:
            raise ValueError('表格中 `病人医生` 列存在问题。')

    # 将末尾的医生添加到信息中
    # 判断该医生名是否应当计入
    is_a_doctor, cur_doctor_name = judge_name(cur_doctor_name)
    if is_a_doctor:
        if cur_doctor_name in doctor_fee_info_dict:
            doctor_fee_info_dict[cur_doctor_name].update(
                combine_two_dict(doctor_fee_info_dict[cur_doctor_name], tmp_doctor_fee_info))
        else:
            doctor_fee_info_dict.update({cur_doctor_name: tmp_doctor_fee_info})
    else:
        print('###: ', cur_doctor_name)

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
        self.receiving_department_list = list()
        self.inspection_testing_list = ['128层CT室', 'CT室', '1.5T磁共振室', 'DR室',
                                        '彩超室', '心电图室', '内镜室', '检验科']
        self.dermatology_clinic_list = [
            '皮肤科308照射室', '皮肤科二氧化碳激光室', '皮肤科光疗室',
            '皮肤科果酸换肤室', '皮肤科He-Ni激光室', '皮肤科红蓝光治疗室',
            '皮肤科皮肤镜检查室', '皮肤科冷疗室', '皮肤科检查室', '皮肤科氦氖激光室',
            '皮肤科冷冻室', '皮肤中医治疗室', '皮肤科门诊', '皮肤科外擦药室']

        self.recovery_clinic_list = ['康复医学科病区', '康复医学科']
        self.cosmetic_surgery_list = ['美容外科诊室']
        self.cosmetic_dermatology_list = ['美容皮肤科诊室']
        self.special_need_list = ['特需病区']
        self.dermatology_ward_1_list = ['皮肤科（一病区）']
        self.dermatology_ward_2_list = ['皮肤科（二病区）']
        self.dermatology_ward_3_list = ['皮肤科（三病区）']
        self.dermatology_ward_4_list = ['皮肤科（四病区）']

    def get_xls_content(self, xls_file_path):
        """ 获取 xls 文件的 pandas 实例

        :param xls_file_path:
        :return:
        """
        self.xls_content = pd.read_excel(io=xls_file_path)

    def get_receiving_department_list(self):
        self.receiving_department_list = sorted(find_all_receiving_department(self.xls_content))


class SelectFilePage(object):
    def __init__(self, root, data_obj):
        self.root = root
        self.data_obj = data_obj

        self.interface = tk.Frame(self.root)
        self.interface.pack()

        title_label = tk.Label(self.interface, text='1. 选择 xls 文件，并指定该文件是 住院|门诊',
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
        title_label = tk.Label(self.interior, text='二. 确定类别、子项', font='Helvetica 13 bold')
        title_label.grid(row=horizontal, column=vertical)

        self.inspection_testing_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.inspection_testing_list, horizontal, vertical,
            sub_title='1. 检查检验费合计', bg_color='light blue')
        horizontal += 1

        self.dermatology_clinic_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_clinic_list, horizontal, vertical,
            sub_title='2. 皮肤科门诊费合计', bg_color='red')
        horizontal += 1

        self.recovery_clinic_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.recovery_clinic_list, horizontal, vertical,
            sub_title='3. 康复类合计', bg_color='light yellow')
        horizontal += 1

        self.cosmetic_surgery_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.cosmetic_surgery_list, horizontal, vertical,
            sub_title='4. 美容外科费用', bg_color='light green')
        horizontal += 1

        self.cosmetic_dermatology_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.cosmetic_dermatology_list, horizontal, vertical,
            sub_title='5. 美容皮肤科费用', bg_color='gray')
        horizontal += 1

        self.special_need_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.special_need_list, horizontal, vertical,
            sub_title='6. 特需费用', bg_color='pink')
        horizontal += 1

        self.dermatology_ward_1_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_ward_1_list, horizontal, vertical,
            sub_title='7. 皮肤科（一病区）', bg_color='blue')
        horizontal += 1

        self.dermatology_ward_2_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_ward_2_list, horizontal, vertical,
            sub_title='8. 皮肤科（二病区）', bg_color='green')
        horizontal += 1

        self.dermatology_ward_3_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_ward_3_list, horizontal, vertical,
            sub_title='9. 皮肤科（三病区）', bg_color='red')
        horizontal += 1

        self.dermatology_ward_4_check_buttons, horizontal = self._get_sub_item_info(
            self.data_obj.dermatology_ward_4_list, horizontal, vertical,
            sub_title='10. 皮肤科（四病区）', bg_color='yellow')
        horizontal += 1

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
        subtitle_label = tk.Label(self.interior, text=sub_title, font='Helvetica 11 bold')
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

        self.data_obj.recovery_clinic_list = [
            text for button, text in zip(
                self.recovery_clinic_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.cosmetic_surgery_list = [
            text for button, text in zip(
                self.cosmetic_surgery_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.cosmetic_dermatology_list = [
            text for button, text in zip(
                self.cosmetic_dermatology_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.special_need_list = [
            text for button, text in zip(
                self.special_need_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.dermatology_ward_1_list = [
            text for button, text in zip(
                self.dermatology_ward_1_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.dermatology_ward_2_list = [
            text for button, text in zip(
                self.dermatology_ward_2_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.dermatology_ward_3_list = [
            text for button, text in zip(
                self.dermatology_ward_3_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

        self.data_obj.dermatology_ward_4_list = [
            text for button, text in zip(
                self.dermatology_ward_4_check_buttons,
                self.data_obj.receiving_department_list) if button.get() == 1]

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

        doctor_base_index = find_key_item_index(self.data_obj.xls_content, keywords='病人医生')
        receiving_department = find_key_item_index(self.data_obj.xls_content, keywords='接收科室')
        medical_income_total = find_key_item_index(self.data_obj.xls_content, keywords='医疗收入小计')
        sanitation_material_fee = find_key_item_index(self.data_obj.xls_content, keywords='卫生材料费')
        consultation_fee = find_key_item_index(self.data_obj.xls_content, keywords='诊查费')

        self.inspection_testing_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.inspection_testing_list,
            doctor_base_index, receiving_department,
            medical_income_total, sanitation_material_fee, consultation_fee,
            self.data_obj.file_type)

        self.dermatology_clinic_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_clinic_list,
            doctor_base_index, receiving_department,
            medical_income_total, sanitation_material_fee, consultation_fee,
            self.data_obj.file_type)

        self.recovery_clinic_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.recovery_clinic_list,
            doctor_base_index, receiving_department,
            medical_income_total, sanitation_material_fee, consultation_fee,
            self.data_obj.file_type)

        self.cosmetic_surgery_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.cosmetic_surgery_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.cosmetic_dermatology_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.cosmetic_dermatology_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.special_need_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.special_need_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.dermatology_ward_1_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_ward_1_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.dermatology_ward_2_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_ward_2_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.dermatology_ward_3_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_ward_3_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        self.dermatology_ward_4_info_dict = class_statistics(
            self.data_obj.xls_content, self.data_obj.dermatology_ward_4_list, doctor_base_index,
            receiving_department, medical_income_total, sanitation_material_fee,
            consultation_fee,
            self.data_obj.file_type)

        text_canvas = tk.Text(self.interface, height=42, width=120)
        print_detail_text, csv_file_path = self.print_and_write_file()
        text_canvas.grid(row=1, column=0, columnspan=4)
        text_canvas.insert(tk.END, print_detail_text)

        text_path_canvas = tk.Text(self.interface, height=2, width=120)
        text_path_canvas.grid(row=2, column=0, columnspan=4)
        text_path_canvas.insert(tk.END, csv_file_path)

    def print_and_write_file(self):
        doctor_fees_dict = dict()
        item_list = ['检查检验合计', '皮肤科门诊合计', '康复合计', '美容外科诊室', '美容皮肤科诊室',
                     '特需病区', '皮肤科（一病区）', '皮肤科（二病区）', '皮肤科（三病区）', '皮肤科（四病区）']
        doctor_fees_txt = '医生姓名,' + ','.join(item_list) + '\n'

        num_format = '{:.2f}'
        print_text = [self.data_obj.file_type]
        for doctor_name in self.inspection_testing_info_dict:
            print_text.append(doctor_name + ' ' + self.data_obj.file_type)
            tmp_dict = {'检查检验合计': sum(list(self.inspection_testing_info_dict[doctor_name].values())),
                        '皮肤科门诊合计': sum(list(self.dermatology_clinic_info_dict[doctor_name].values())),
                        '康复合计': sum(list(self.recovery_clinic_info_dict[doctor_name].values())),
                        '美容外科诊室': sum(list(self.cosmetic_surgery_info_dict[doctor_name].values())),
                        '美容皮肤科诊室': sum(list(self.cosmetic_dermatology_info_dict[doctor_name].values())),
                        '特需病区': sum(list(self.special_need_info_dict[doctor_name].values())),
                        '皮肤科（一病区）': sum(list(self.dermatology_ward_1_info_dict[doctor_name].values())),
                        '皮肤科（二病区）': sum(list(self.dermatology_ward_2_info_dict[doctor_name].values())),
                        '皮肤科（三病区）': sum(list(self.dermatology_ward_3_info_dict[doctor_name].values())),
                        '皮肤科（四病区）': sum(list(self.dermatology_ward_4_info_dict[doctor_name].values()))}

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
            if len(self.recovery_clinic_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '康复合计: ' + num_format.format(
                    sum(list(self.recovery_clinic_info_dict[doctor_name].values()))))
                for key, val in self.recovery_clinic_info_dict[doctor_name].items():
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
            if len(self.special_need_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '特需病区: ' + num_format.format(
                    sum(list(self.special_need_info_dict[doctor_name].values()))))
                for key, val in self.special_need_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.dermatology_ward_1_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科（一病区）: ' + num_format.format(
                    sum(list(self.dermatology_ward_1_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_ward_1_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.dermatology_ward_2_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科（二病区）: ' + num_format.format(
                    sum(list(self.dermatology_ward_2_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_ward_2_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.dermatology_ward_3_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科（三病区）: ' + num_format.format(
                    sum(list(self.dermatology_ward_3_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_ward_3_info_dict[doctor_name].items():
                    print_text.append('\t\t' + key + ': ' + num_format.format(val))
            if len(self.dermatology_ward_4_info_dict[doctor_name].items()) > 0:
                print_text.append('\t' + '皮肤科（四病区）: ' + num_format.format(
                    sum(list(self.dermatology_ward_4_info_dict[doctor_name].values()))))
                for key, val in self.dermatology_ward_4_info_dict[doctor_name].items():
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


