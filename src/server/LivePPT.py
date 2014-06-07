# -*- coding: UTF-8 -*-

'''
Created on 2013-4-8

@author: simon_000
'''
from __future__ import with_statement

#导入
import config as conf
import log  
import path_utils as p
import redis
import mgfs

import os
import sys
import logging
import json
import traceback
import time

import win32com.client
import win32com.gen_py.MSO as MSO
import win32com.gen_py.MSPPT as PO


def main():
    try:
        g = globals()
        for c in dir(MSO.constants): g[c] = getattr(MSO.constants, c) # globally define these
        for c in dir(PO.constants): g[c] = getattr(PO.constants, c)

        #准备Logger
        logger = log.gen_logger()
        logger.info('logger ready')

        #准备GridFS
        gfs = mgfs.MGFS()

        #准备RedisMQ
        redis_client = redis.Redis(conf.REDIS_SVRIP, conf.REDIS_SVRPORT)

        # 准备PPT应用
        Application = win32com.client.Dispatch(conf.POWERPOINT_APPLICATION_NAME)
        Application.Visible = True


        reload(sys)
        sys.setdefaultencoding(conf.UTF8_ENCODING)
        print sys.getdefaultencoding()
        logger.debug('LivePPT-PPT-Converter is now launched.')

        while True:
        # 阻塞等待消息  
            logger.info('Waiting for jobs......')
            msg = redis_client.brpop(conf.PPT_CONVERT_REQMQ)
            if msg <> None:
                continue
            ppt_id = msg[1]
            logger.info(ppt_id)
                
            #准备路径参数
            ppt_path = p.gen_ppt_path(ppt_id) #PPT存放位置
            save_dir_path = p.gen_save_dir_path(ppt_id) #保存转换后图片的文件夹路径
                
            #从Gridfs获取ppt文件 
            with open(ppt_path, "wb") as ppt_file:
                data = gfs.getPPT(ppt_id)
                ppt_file.write(data.getvalue())
            #使用PowerPoint打开本地PPT，并进行转换
            try:
                myPresentation = Application.Presentations.Open(ppt_path)
                myPresentation.SaveAs(save_dir_path, ppSaveAsJPG)
            finally:
                 myPresentation.Close()
                    
            png_file_name_list = os.listdir(save_dir_path)
            ppt_count = len(png_file_name_list)
            print 'ppt_page_count'+str(ppt_count)
            for index in range(1, ppt_count+1):
                png_path = p.gen_single_png_path(ppt_id, index) #单个PNG文件路径
                png_key = ppt_id + "p"+ str(index)

                #上传单个文件
                with open(png_path, "rb") as img_file:
                    mgfs.putImg(image_fd = img_file, img_id = png_key, img_format = 'jpg')
                    logger.info( "uploaded " + str(index))
                
            #组装准备发到RedisMQ
            mqmsg = {}
            mqmsg['isSuccess'] = True
            mqmsg['storeKey'] = ppt_id
            mqmsg['pageCount'] = ppt_count
            redis_client.lpush(conf.PPT_CONVERT_RESPMQ, mqmsg)
                
    except Exception as e :
        logger.debug('Unexpected ERROR!')
        logger.debug(e)
        traceback.print_exc()
    finally:
        time.sleep(5)
        logger.debug('Try to relaunch LivePPT-Python!')
        main()
    return

if __name__ == '__main__':
    main()
