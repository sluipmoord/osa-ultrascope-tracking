import datetime
import time
import xlrd
import bisect


class UEX(object):
    def __init__(self, file_path, sheets_names):
        self.book = xlrd.open_workbook(file_path,
                                       on_demand=True)
        self.LUT = {}
        self.sorted_ts = []

        try:
            self._create_lookup_table(self.book, sheets_names)
        except Exception as e:
            self.book.release_resources()
            raise e

    def _create_lookup_table(self, book, sheets_names):
        for sheet_name in sheets_names:
            sheet = book.sheet_by_name(sheet_name)

            for row in range(0, sheet.nrows):
                # YEAR MON DAY HR MIN SEC
                ncol = len(sheet.row(row))
                date_tuple = [sheet.cell(row, c).value for c in range(0, 6)]

                try:
                    # YEAR MON DAY HR MIN SEC
                    date = datetime.datetime(
                        int(date_tuple[0]), int(date_tuple[1]),
                        int(date_tuple[2]), int(date_tuple[3]),
                        int(date_tuple[4]), int(date_tuple[5]))

                    time_stamp = time.mktime(date.timetuple())

                    data = [sheet.cell(row, c).value for c in range(6, ncol)]
                    self.LUT[time_stamp] = data

                except Exception as e:
                    # assume it is a invalid row
                    print e, date_tuple
                    pass

        book.release_resources()
        self.sorted_ts = sorted(self.LUT)

    def get_data(self, date):
        """
        return (ascension, declination, object_location)
        ascension {H, M, S}
        declination {deg, M, S}
        object_location {Azimuth, Altitude}

        """

        time_stamp = time.mktime(date.timetuple())
        s = self.sorted_ts
        i = bisect.bisect_left(self.sorted_ts, time_stamp)

        # get closet matching time stamp
        ts = min(s[max(0, i-1): i+2], key=lambda t: abs(time_stamp - t))
        data_set = self.LUT[ts]

        ascension = [data_set[i] for i in range(0, 3)]
        declination = [data_set[i] for i in range(3, 6)]
        object_location = [data_set[i] for i in range(6, 8)]

        return {
            "ascension": ascension,
            "declination": declination,
            "object_location": object_location,
            "date": date
        }

    def track_object(self, start_date, interval, duration, track_callback):
        """
        `start_date` datetime to start tracking from
        `interval` give new tracking data every # of seconds
        `duration` track for how long in seconds
        `track_callback` passes result of get_data(date) to function
        """

        start_time = start_date
        end_time = start_time + datetime.timedelta(seconds=duration)

        while start_time <= end_time:
            data = self.get_data(start_time)
            if track_callback:
                track_callback(data)

            start_time = start_time + datetime.timedelta(seconds=interval)
            time.sleep(interval)





