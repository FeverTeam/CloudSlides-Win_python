# -*- coding: UTF-8 -*-

def gen_base_path():
    return u"C:\\ppt"

def gen_ppt_path(ppt_id):
    base_dir_path = gen_base_path()
    return base_dir_path + u"\\"+unicode(ppt_id);

def gen_save_dir_path(ppt_id):
    base_dir_path = gen_base_path()
    return base_dir_path+u"\\converted_"+unicode(ppt_id);

def gen_single_png_path(ppt_id, index):
    save_dir_path = gen_save_dir_path(ppt_id);
    return save_dir_path+u"\\幻灯片"+unicode(str(index))+u".JPG";