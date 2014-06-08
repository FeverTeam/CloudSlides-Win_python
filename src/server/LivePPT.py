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
    
    logger = log.gen_logger()
    logger.info('logger ready')
    
    try:
        g = globals()
        for c in dir(MSO.constants): g[c] = getattr(MSO.constants, c) # globally define these
        for c in dir(PO.constants): g[c] = getattr(PO.constants, c)

        #检测文件的根目录是否存在
        base_path = p.gen_base_path()
        if os.path.isdir(base_path): 
            pass 
        else: 
            os.mkdir(base_path)

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
           
            msg = redis_client.brpop(conf.PPT_CONVERT_REQMQ, timeout=30)
            if msg == None:
                logger.info('No message in 30 sec......')
                continue
            
            ppt_id = msg[1]
            logger.info('Get convert job pptid = %s' %(ppt_id))
                
            #准备路径参数
            ppt_path = p.gen_ppt_path(ppt_id) #PPT存放位置
            save_dir_path = p.gen_save_dir_path(ppt_id) #保存转换后图片的文件夹路径
                
            #从Gridfs获取ppt文件
            try:
                with open(ppt_path, "wb") as ppt_file:
                    data = gfs.getPPT(ppt_id)
                    ppt_file.write(data)
            except Exception , e :
                print e

            #使用PowerPoint打开本地PPT，并进行转换
            try:
                myPresentation = Application.Presentations.Open(ppt_path)
                myPresentation.SaveAs(save_dir_path, ppSaveAsJPG)
            finally:
                myPresentation.Close()
                #顺手清理PPT文件
                os.remove(ppt_path);

                
            img_file_name_list = os.listdir(save_dir_path)
            ppt_count = len(img_file_name_list)
            print 'ppt_page_count'+str(ppt_count)
            for index in range(1, ppt_count+1):
                img_path = p.gen_single_png_path(ppt_id, index) #单个PNG文件路径
                img_key = ppt_id + "p"+ str(index)

                #上传单个文件
                try:
                    with open(img_path, "rb") as img_file:
                        gfs.putImg( image_fd = img_file, img_id = img_key, img_format = 'jpg' )
                        logger.info( "uploaded image[%s] complete" %( img_path ) )
                finally:
                    #顺手清理图片文件
                    os.remove(img_path);
                
            #组装准备发到RedisMQ
            mqmsg = {}
            mqmsg['isSuccess'] = True
            mqmsg['storeKey'] = ppt_id
            mqmsg['pageCount'] = ppt_count
            redis_client.lpush(conf.PPT_CONVERT_RESPMQ, mqmsg)
                
    except Exception as e :
        logger.debug('Unexpected ERROR!')
        traceback.print_exc()
    finally:
        time.sleep(3)
        logger.debug('Try to relaunch LivePPT-Python!')
        main()
    return

if __name__ == '__main__':
    main()
