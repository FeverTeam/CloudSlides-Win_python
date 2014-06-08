# -*- coding: UTF-8 -*-5
import pymongo
import gridfs
import redis
import os
import sys

sys.path.append('..')
import server.config as conf


#对GridFS存入ppt
def uploadPPT(ppt_id, ppt_type, file):
        #建立Gridfs链接 
        con = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT)
        gfs = gridfs.GridFS(con[conf.MONGO_PPTDB])
        
        result = gfs.put(file.read(), filename = ppt_id, contentType = ppt_type)
        print result

def getImage(image_id):
        con = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT)
        gfs = gridfs.GridFS(con[conf.MONGO_IMGDB])
	data = gfs.get(image_id):
        return data

#对 RedisMQ发送任务
def sendJob(ppt_id):
        client = redis.Redis(conf.REDIS_SVRIP, conf.REDIS_SVRPORT)
	msg = {}
	msg['pptId'] = ppt_id
        result = client.lpush(conf.PPT_CONVERT_REQMQ, msg)
        print result

def waitForComplete():
        client = redis.Redis(conf.REDIS_SVRIP, conf.REDIS_SVRPORT)
        b = True
        while b:
                result = client.brpop(conf.PPT_CONVERT_RESPMQ)
                if result == None:
                        continue
                print "msg = ", result
                return result[1]

def main():
        
        #1.上传PPT并且递交任务, 以上传在GridFS时获得的_id为准
        ppt_name = '1.pptx'
        ppt_id = None
	ppt_name = os.path.splitext(ppt_name)[0]
        ppt_type = os.path.splitext(ppt_name)[1]
        with open(ppt_name, "rb") as ppt_file:
                ppt_id = uploadPPT(ppt_name, ppt_type, ppt_file)

	#2.发送任务 
        sendJob(ppt_id)
        
        
        #3.获取完成的消息
        complete_message = waitForComplete()
	print "Got message: " complete_message
        
        #4.把图片抓回到本地
        img_id_list = eval(complete_message)['imgIdList']
        for img_id in img_id_list:
		img_file_cache = getImage(img_id)
                with open("%s.jpg" %(img_id), "wb") as img_file:
                        img_file.write( img_file_cache.read() )
                        print "Got Image %s" %(img_id)

if __name__ == '__main__':
        main()
