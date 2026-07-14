from motor.motor_asyncio import AsyncIOMotorClient

from . import config

_client = AsyncIOMotorClient(config.MONGO_URI)
database = _client[config.MONGO_DB_NAME]

workbooks_collection = database["workbooks"]
extractions_collection = database["extractions"]
structured_collection = database["structured"]
