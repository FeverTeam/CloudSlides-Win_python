import pymongo
import gridfs

import config as conf
import log

class MGFS:
	mongo_connect = None
	mongo_pptfs = None
	mongo_imgfs = None
	logger = None

	def __init__(self):
		MGFS._connect()
		logger = log.gen_logger()

	def _connect(self):
		if not MGFS.mongo_connect:
			try:
				MGFS.mongo_connect = pymongo.Connection(conf.MONGO_SVRIP, conf.MONGO_SVRPORT);
				MGFS.mongo_pptfs = MGFS.mongo_connect[conf.MONGO_PPTDB].GridFS();
				MGFS.mongo_imgfs = MGFS.mongo_connect[conf.MONGO_IMGDB].GridFS();
			except:
				logger.warning("Connection to mongoDB [%s:%s] fail" %(conf.MONGO_SVRIP, MONGO_SVRPORT))

	def getPPT(self, ppt_id):
		data = None
		try:
			grid_out = MGFS.mongo_pptfs.find({"filename": ppt_id}, timeout=False)
			if( len(grid_out) <> 1 ) 
				logger,warning("%d file named [%s] in GridFS" %(len(grid_out, pptid)))
			data = grid_out[0].read();
			return data
		except:
			logger.warning("fail to get ppt [%s] from GridFS" %(pptid)))

	def putImg(self, image_fd, img_id, img_format = 'jpg'):
		try:
			result = MGFS.mongo_imgfs.put(image_fd.read(), filename = img_id);
			logger.info(result)
		except:
			logger.warning("fail to put image [%s] to GridFS" %(img_id)))

	
