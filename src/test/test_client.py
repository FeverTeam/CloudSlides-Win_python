import pymongo
import gridfs
import redis
import os

import src.server.config as conf


#对GridFS存入ppt
def uploadPPT(ppt_id, ppt_tpye, file):
	#建立Gridfs链接 
	con = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT)
	gfs = gridfs.GridFS(con[conf.MONGO_PPTDB])
	result = gfs.put(file.read(). filename = ppt_id, contentType = ppt_type)
	print result

def getImage(image_id)
	con = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT)
	gfs = gridfs.GridFS(con[conf.MONGO_IMGDB])
	grid_out = gfs.find( {"filename": image_id}, timeout=False)
	return grid_out[0]



#对 RedisMQ发送任务
def sendJob(ppt_id):
	client = redis.Redis(conf.REDIS_SVRIP, conf.REDIS_SVRPORT)
	result = client.lpush(conf.PPT_CONVERT_REQMQ, pptid)
	print result

def waitForComplete()
	client = redis.Redis(conf.REDIS_SVRIP, conf.REDIS_SVRPORT)
	result = client.brpop(conf.PPT_CONVERT_RESPMQ)
	return result

def main():
	ppt_name = 'test.pptx'
	ppt_id = 'abc123'
	ppt_type = os.path.splitext(ppt_name)[1]
	with open(ppt_name, "rb") as ppt_file:
		uploadPPT(ppt_id, ppt_type, ppt_file)
	sendJob(ppt_id)

	#获取完成的消息
	complete_message = waitForComplete()
	print complete_message

	#把图片抓回到本地 
	ppt_count = complete_message['pageCount']
	for index in range (1, ppt_count + 1):
		image_name = "%sp%d" %(ppt_id, index)
		data = getImage(image_name)
		with open("%s.jpg" %(image_name), "wb") as img_file:
			img_file.write( data.read() )
			print "Got Image %s" %(image_name)



		


if __name__ == '__main__'
	main()
