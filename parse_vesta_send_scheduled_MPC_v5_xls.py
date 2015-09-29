#!/env/python
#import daemon # install python-daemon from pypi
import sys, sched
from datetime import datetime as dt
import datetime, time
import codecs, csv, re, xlrd, xlwt
import threading, Queue,  ephem, serial, numpy

import ephem

from skyfield.api import earth, mars, jupiter, moon,  now
from skyfield.units import Angle

from astropy import units as u
from astropy.coordinates import SkyCoord


def dms2dec(dms_str):
    dms_str = re.sub(r'\s', '', dms_str)
    dms_str = re.sub(r'\-', '', dms_str)
    dms_str = re.sub(r'\+', '', dms_str)
    if re.match('[swSW]', dms_str):
        sign = -1
    else:
        sign = 1
    (degree, minute, second, frac_seconds, junk) = re.split('\D+', dms_str, maxsplit=4)
    frac_seconds_len = len(frac_seconds)
    frac_seconds = float(frac_seconds)
    for i in xrange(frac_seconds_len):
        frac_seconds = frac_seconds / 10.0
    print(degree,minute,second,frac_seconds)
    return sign * (int(degree) + float(minute) / 60 + float(second) / 3600 + float(frac_seconds) / 3600)



def hms2dec(dms_str):
    """
    Return decimal representation of HMS

    """
    dms_str = re.sub(r'\s', '', dms_str)
    dms_str = re.sub(r'-', '', dms_str)
    dms_str = re.sub(r'\+', '', dms_str)
    if re.match('[swSW]', dms_str):
        sign = -1
    else:
        sign = 1
    (hours, minute, second, frac_seconds, junk) = re.split('\D+', dms_str, maxsplit=4)
    frac_seconds_len = len(frac_seconds)
    frac_seconds = float(frac_seconds)

    for i in xrange(frac_seconds_len):
        frac_seconds = frac_seconds / 10.0

    degree = 15 * int(hours)
    print(degree,minute,second,frac_seconds)
    return sign * (int(degree) + float(minute) / 60 + float(second) / 3600 + float(frac_seconds) / 3600)


def radec2decstring(ra,dec):
    #c = SkyCoord(ra=10.68458*u.degree, dec=41.26917*u.degree)
    c = SkyCoord(ra=ra*u.degree, dec=dec*u.degree)
    c.to_string('decimal')
    return c


def altaz2decstring(alt,az):
    #c = SkyCoord(ra=10.68458*u.degree, dec=41.26917*u.degree)
    c = SkyCoord(ra=alt*u.degree, dec=az*u.degree)
    c.to_string('decimal')
    return c



def decdeg2dms(dd):
    # convert decimal degrees to deg, min, sec
    is_positive = dd >= 0
    dd = abs(dd)
    minutes,seconds = divmod(dd*3600,60)
    degrees,minutes = divmod(minutes,60)
    degrees = degrees if is_positive else -degrees
    return (degrees,minutes,seconds)



def radec2altaz(ra,dec):
    gbt = ephem.Observer()
    gbt.long = '-04:00:00 W'
    gbt.lat = '36:00:00.00'
    gbt.pressure = 0 # no refraction correction.
    gbt.epoch = ephem.J2000
    # Set the date to the epoch so there is nothing changing.
    gbt.date = '2015/10/02 01:30:00'
    target = ephem.FixedBody()
    target._ra = ra
    target._dec = dec

    target._epoch = ephem.J2000
    target.compute(gbt)
    alt = target.alt
    az = target.az
    return (alt,az)



def u2ascii(ustring):
    udata=ustring.decode("utf-8","ignore")
    asciidata=udata.encode("ascii","ignore")
    return asciidata
    #return udata



def send2Arduino(alt, az):
    # for live script take off quotes
    #ser = serial.Serial("COM1", 9600)   # open serial port that Arduino
    #print ser #print ser  connection info
    print 'send alt, az > mount '

    # send alt/az data command to serial
    # refer to Leo/Jordan to command format in front of ra dec vars
    #ser.write(ra,dec)                  # print serial config
    print (alt, az)


def now_str():
    """Return hh:mm:ss string representation of the current time."""
    t = dt.now().time()
    return t.strftime("%H:%M:%S")



def main():
    def push_mount_messages_queue(message, alt, az):
        print 'RUNNING:', now_str(), message
        # Do whatever you need to do here
        ## then re-register task for same time tomorrow
        send2Arduino(alt, az)
        t = dt.combine(dt.now() + datetime.timedelta(days=1), daily_time)
        scheduler.enterabs(time.mktime(t.timetuple()), 1, push_mount_messages_queue, ('Running again',alt,az))

    # Build a scheduler object that will look at absolute times
    scheduler = sched.scheduler(time.time, time.sleep)
    print 'START:', now_str()

    # get data to build task queue
    #filePath = 'Vesta-position-data.xls'
    positions = xlrd.open_workbook('Vesta-Position-Data-MPC_V5.xls', on_demand = True)
    sheet = positions.sheet_by_name('Sheet1')

    rows= range(3,300)   #!!!! change this range to fit the numbe rof columns in the final M
                        # PC data file with 1s intervals of the final MPC dataset
                        # only make good data rows visible
    rowid= 0

    for row in rows:
        # each row is am ra dec data point we push out
        rowid += 1
        #print 'row ->' + str(rowid)t

        # col format in MPC V5 data file ......
        # YEAR	MON	DAY	HR	MIN	SEC	H	M	S	DEG	M	S	AZI	ALT
        (year, month, day) = [sheet.cell(row, 0).value, sheet.cell(row, 1).value, sheet.cell(row, 2).value]
        (hour, minute, second) = [sheet.cell(row, 3).value, sheet.cell(row, 4).value, sheet.cell(row, 5).value]

        # leave ra and dec for now, not using it
        (ra_h, ra_m, ra_s) =[0,0,0]
        (dec_deg, dec_m, dec_s)=[0,0,0]
        #(ra,dec) = [sheet.cell(row, 2).value, sheet.cell(row, 3).value]
        #print 'ra dec before decimal conversion, after nomralising'
        #print (ra,dec)

        (az, alt) = [sheet.cell(row, 12).value, sheet.cell(row, 13).value]

        # Put new ra dec movement task for today at derived time 7am on queue.
        # Executes immediately if past 7am
        daily_time = datetime.time(int(hour), int(minute), int(second))
        time_values = daily_time
        first_time = dt.combine(dt.now(), daily_time)
        print '___________________________________________________'
        print 'time now is ' + now_str()
        print 'target-time is '  + str(time_values)
        print 'passed target time' +  str(first_time)
        print str(first_time)
        # time, priority, callable, *args

        # make decimal ra dec
        #(ra,dec)= radec2decstring(ra,dec)
        #print 'ra dec '
        #print (ra, dec)

        print (alt,az)
        print 'decimal alt az'
        print (alt,az)
        scheduler.enterabs(time.mktime(first_time.timetuple()), 1, push_mount_messages_queue, ('move message ',alt,az))

    scheduler.run()

if __name__ == '__main__':
    if "-f" in sys.argv:
        main()
    else:
        #with daemon.DaemonContext():
        main()
