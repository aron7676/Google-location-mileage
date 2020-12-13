import json, datetime, googlemaps,xlsxwriter, os


def get_drive_details (orig_lat,orig_lng,dest_lat, dest_lng):
    API_key = 'YOUR_API_KEY'  # enter the key you got from Google. I removed mine here
    gmaps = googlemaps.Client(key=API_key)
    origin = int(orig_lat)/1e7, int(orig_lng)/1e7
    destination = int(dest_lat)/1e7, int(dest_lng)/1e7
    return gmaps.distance_matrix(origin, destination, mode='driving')


def td_format(td_object):
    hours, remainder = divmod(td_object.total_seconds(), 3600)
    minutes, seconds = divmod(remainder, 60)
    hours = int(hours)
    minutes = int(minutes)
    seconds = int(seconds)
    return '%s:%s:%s' % (hours, minutes,seconds)


def get_tripDetails(activity):
        print(activity['activitySegment']["activityType"])
        transport = (activity['activitySegment']["activityType"])
        start = (datetime.datetime.fromtimestamp(int(activity['activitySegment']["duration"]['startTimestampMs'])/1000.0))
        end = (datetime.datetime.fromtimestamp(int(activity['activitySegment']["duration"]['endTimestampMs']) / 1000.0))
        difference = end - start
        travelTime = td_format(difference)
        trip = get_drive_details(activity['activitySegment']['startLocation']['latitudeE7'],activity['activitySegment']['startLocation']['longitudeE7'],activity['activitySegment']['endLocation']['latitudeE7'],activity['activitySegment']['endLocation']['longitudeE7'])
        calcDistance = int(trip["rows"][0]["elements"][0]["distance"]["value"]) * 0.000621371
        origin = trip["origin_addresses"][0]
        dest = trip["destination_addresses"][0]
        print(str(travelTime))
        print(calcDistance)
        details = [start.strftime("%m/%d/%Y %H:%M:%S"),end.strftime("%m/%d/%Y %H:%M:%S"),transport,str(difference),round(calcDistance,2),origin,dest]
        return details

def find_filenames( path_to_dir, suffix=".json" ):
    filenames = os.listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) ]
    
    

#Start of program

if os.path.exists("Mileage.xlsx"):
  os.remove("Mileage.xlsx")

files = find_filenames(os.getcwd())
counter = 0
wb = xlsxwriter.Workbook('Mileage.xlsx')
ws = wb.add_worksheet()
header = ['start','end','transport','travelTime','distance','origin','destination']
for col_num, data in enumerate(header):
    ws.write(0, col_num, data)
row = 1


for file in files:
    with open(file) as f:
        # Load in Google Location History data.
        data = json.load(f)

    activity = ''
    for i in data["timelineObjects"]:
        try:
            i['activitySegment']
            activity = i
            counter += 1
            detail = get_tripDetails(activity)
            ws.write_row(row, 0, detail)
            row += 1
        except:
            pass

workbook.close()





print("{} Locations visited!".format(counter))



