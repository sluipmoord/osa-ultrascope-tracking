from uex import UEX
import datetime


def main():

    uex = UEX('random_data.xls',
              ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5', 'Sheet6'])

    user_input = datetime.datetime(2015, 10, 1, 5, 0, 0)
    # number of photos times exposure time in seconds
    duration = 4 * 3
    # duration = (datetime.datetime(2015, 10, 5, 5, 20, 0) - user_input).total_seconds()

    uex.track_object(user_input, 1, duration, send2Arduino, 1)


def send2Arduino(data):
    print data
    az = data['object_location'][0]
    alt = data['object_location'][1]

    # for live script take off quotes
    # ser = serial.Serial("COM1", 9600)   # open serial port that Arduino
    # print ser #print ser  connection info
    print 'send alt, az > mount '

    # send alt/az data command to serial
    # refer to Leo/Jordan to command format in front of ra dec vars
    # ser.write(ra,dec)                  # print serial config
    print (alt, az)

if __name__ == '__main__':
    main()
