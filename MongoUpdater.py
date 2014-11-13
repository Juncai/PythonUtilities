__author__ = 'Jon'

from pymongo import MongoClient
import pymongo
import datetime

client = MongoClient('mongodb://Admin:Plait2014Lab@blazing.ccs.neu.edu:27017')
par_id = "OA23001"
# par_ids = ["OA23003", "OA23004", "OA23005", "OA23006", "OA23007",
#           "OA23008", "OA23010", "OA23012", "OA23013",
#           "MA23021", "MA23002", "MA23003", "MA23024", "MA23005",
#           "MA23013", "MA23014",
#           "YA23001", "YA23003", "YA23007", "YA23010", "YA23011",
#           "YA23040"]

par_ids = ["MA23021", "MA23002", "MA23003", "MA23024", "MA23005",
           "MA23013", "MA23014"]

# error when OA23013
# db  = client['videotracker_Production']
# new_participants = db['newParticipants']

db  = client['videotracker_Production']
new_participants = db['newParticipants']
intervals = (17, 16, 17)


for pid in par_ids:
    start_gp = new_participants.find({"Participant ID" : pid}, {"System Time" : 1}).sort("System Time", 1).limit(1)

    # print start_gp.next()

    i = 1
    prev_ts = datetime.datetime.strptime(start_gp.next()['System Time'], "%H:%M:%S.%f")
    # print prev_ts.microsecond
    # prev_ts += datetime.timedelta(milliseconds=intervals[0])
    # print prev_ts.microsecond
    # ts_str = datetime.datetime.strftime(prev_ts, "%H:%M:%S.%f")
    # print ts_str[:-3]
    print datetime.datetime.strftime(prev_ts, "%H:%M:%S.%f")

    gps = new_participants.find({"Participant ID" : pid}).sort("System Time", 1)
    for gp in gps:
        if i > 1:
            # cur_ts = datetime.datetime.strptime(gp['System Time'], "%H:%M:%S.%f")
            cur_ts = prev_ts + datetime.timedelta(milliseconds=intervals[(i-2)%3])
            gp['System Time'] = datetime.datetime.strftime(cur_ts, "%H:%M:%S.%f")[:-3]
            prev_ts = cur_ts
            print gp['System Time']
            new_participants.save(gp)
        i += 1



