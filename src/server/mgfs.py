# -*- coding: UTF-8 -*-

import pymongo
import gridfs

import config as conf
import log
import threading

class MGFS:
	mongo_connect = None
	mongo_pptfs = None
	mongo_imgfs = None
	logger = None
	
	def __init__(self):
		MGFS.logger = log.gen_logger()
		MGFS._connect(self)
		
	
	def _connect(self):
		if not MGFS.mongo_connect:
			MGFS.mongo_connect = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT)
			MGFS.mongo_pptfs = gridfs.GridFS(MGFS.mongo_connect[conf.MONGO_PPTDB])
			MGFS.mongo_imgfs = gridfs.GridFS(MGFS.mongo_connect[conf.MONGO_IMGDB])
		

	def getPPT(self, grid_id):
		data = None
		try:
			ppt_file = MGFS.mongo_pptfs.get(grid_id)
			return file
			#for grid_out in MGFS.mongo_pptfs.find({"filename": ppt_id}, timeout=False).sort("uploadDate", -1).limit(1):
                        #    data = grid_out.read()
                        #    return data		
		except Exception, e:
			MGFS.logger.warning("fail to get ppt [%s] from GridFS .Reason:\n %s" %(grid_id, e))

	def putImg(self, image_fd, img_name):
		try:
			result = MGFS.mongo_imgfs.put(image_fd.read(), filename = img_name)
			MGFS.logger.info(result)
			return result
		except:
			MGFS.logger.warning("fail to put image [%s] to GridFS" %(img_name))

	
