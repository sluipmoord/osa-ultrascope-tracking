from uex import UEX
import datetime


def main():

    uex = UEX('Vesta-Position-Data-MPC_v5.xls', ['Sheet1'])

    user_input = datetime.datetime(2015, 10, 5, 4, 10, 0)

    # number of photos times exposure time in seconds
    duration = 4 * 3 * 60

    uex.track_object(user_input, 1, duration, send2Arduino)


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
