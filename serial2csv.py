'''
Update an XLS graph with a stream of new data from multiple Microbits
THIS code integrates the serial sensor data into a consolidated CSV
See README for bigger picture

Design:
    Use Threading and a Queue to separate the serial receiver from file updates
    Regenerate the file every NN seconds with latest info from Queue
    https://www.troyfawkes.com/learn-python-multithreading-queues-basics/
'''
import serial
from serial.tools.list_ports import comports
import time
import datetime
import openpyxl
import re
from queue import Queue, Empty
from threading import Thread
import os
from pathlib import Path
import logging
from logging.handlers import RotatingFileHandler
import shutil


class Serial2CSV():

    qu = Queue()

    def config1(self):
        self.serial_path_ux = '/dev/ttyACM0'
        self.os_name_win = 'nt'
        self.os_name_ux = 'posix'

    def config2(self):
        # TODO: work out the Windows path
        self.csv_path = Path('./ubit_live_data.csv')
        self.new_ubit_dataD = {}
        self.csvLinesL = ['Time']
        self.heading_colsL = []
        self.sampling_period_secs = 3.1
        self.sample_count = int(24.0*3600/self.sampling_period_secs)
        self.continue_serial_read_thread = True
        # self.backup_interval_secs = 300 # 5 minutes
        self.backup_interval_secs = 60 # TODO DEV ONLY
        # self.backup_interval_secs = 5 # TODO DEV ONLY
        self.backup_folder_str = './backups/'
        self.backup_file_prefix = 'ubit_live_data_'
        self.backup_last_saved = 0
        self.backup_file_suffix = '%Y%m%d_%H%M%S'
        # self.logfile_path  defined in logger_setup
        self.new_csv_linesL = []

        '''
        Fake Data for System Testing
        
        Devices are used to pad out the data set
        to evaluate display issues in XLS
        These are needed outside pytest, for system testing
        '''
        self.fake_data_enabled = False
        self.fake_ubit_data_str = 'gbye:-20\nqueet:24\nllkji:23\nfasei:22\nasdic:21\nlkjls:20\naserc:17\nasecg:18\nasxei:19\n'
        # self.fake_devicesL = fake_devs_str.split(' ')

    def csv_update(self):
        '''
        Continuous loop
        Gather serial data for NN seconds
        Cleanup data
        Update headings row with new columns from data
        Insert new data row at top of file

        read-mod-write csv file
        readlines => List
            first line is headings =>List of headings
        Process serial data stream 
            => Hash of Col_heading = Latest Value
            append new headings to Headings Line
            Create a row of values
                Missing values => blank
        '''
        
        permission_error = False
        for i in range(self.sample_count):
            if permission_error is True:
                # Skip Readlines
                # readlines just replaces the existing data
                # it's fine to continue with data from last cycle
                pass
            else:
                try:
                    self.csv_readlines()
                    log.debug(f'read csv')
                except:
                    log.warning('readline error ignored')
                    pass
            try:
                self.csv_save_backup()
            except PermissionError:
                log.warning('error copying to backup')
            permission_Error = False
            self.mk_headings_list()
            log.debug(f'headings made')
            self.dequeue_into_dict()
            self.cleanup_new_data()
            self.csv_add_new_headings()
            self.build_next_data_row()
            self.csv_rebuild()
            try:
                self.csv_writelines()
            except PermissionError:
                permission_error = True
                log.warning('error writing to CSV file')

        self.continue_serial_read_thread = False
    
    def cleanup_new_data(self):
        '''
        Characters get dropped from serial data
        - can't fix missing digits
        - can't fix a missing colon separator
        - CAN fix missing ID characters
        '''
        remove_keysL = []
        fix_keysL = []
        for id in self.new_ubit_dataD.keys():
            if len(id) > 5:
                pass
                # 5 or more characters is a valid ID
            elif len(id) == 5:
                # valid ID, do nothing
                pass
            elif len(id) < 3:
                # too few characters, little hope
                remove_keysL.append(id)
                log.warning(f"ID too short, dropping: '{id}'")
            else:
                # less than 5 characters
                # some characters got dropped
                # try to find the existing id
                log.warning(f"Fixing ID: '{id}'")
                fix_keysL.append(id)
        
        for id in remove_keysL:
            self.new_ubit_dataD.pop(id, None)

        for id in fix_keysL:
            self.new_ubit_dataD.pop(id, None)
            # TODO self.fix_bad_ubit_id(id)


    def fix_bad_ubit_id(self, id):
        # search for a similar ID in the
        # existing list of headings
        # resave the value with the repaired ID

        # Mechanism = regex
        # allow extra characters between any input character
        # .*A.*B.*C.*D.*
        srch_str = ".*" + ".*".join([char for char in id]) + ".*"
        srch_re = re.compile(srch_str)
        log.debug(f"Regex Str='{srch_str}'")
        found_ids = []
        log.info(f'headings={self.heading_colsL}')
        for good_id in self.csvLinesL[0]:
            log.info(f'For bad id={id}, checking={good_id}')
            if re.match(srch_re, good_id):
                found_ids.append(id)
                log.info(f'For bad id={id}, found matches={found_ids}')
        if len(found_ids) != 1:
            # discard the data as
            # it could be more than one valid uBit ID
            # or didn't find any matches
            self.new_ubit_dataD.pop(id)
            log.info(f'deleted1')
        elif found_ids[0] in self.new_ubit_dataD:
            self.new_ubit_dataD.pop(id)
            log.info(f'deleted2')
        else:
            value = self.new_ubit_dataD[id]
            self.new_ubit_dataD[good_id] = value
            self.new_ubit_dataD.pop(id, None)
            log.info(f'replaced')

    def csv_rebuild(self):
        '''
        reconstruct the CSV with
        new Headings Row
        new data row = row 1
        remainder of original data appended
        '''
        heading_line = ','.join(self.heading_colsL) +'\n'
        new_first_line = ','.join(map(str,self.next_data_row)) +'\n'
        self.new_csv_linesL = [heading_line, new_first_line]
        self.new_csv_linesL.extend(self.csvLinesL[1:])

    def detect_invalid_time_column_debug_code(self):
        '''
        Rescan the whole csv for errors
        '''
        for linenum, row in enumerate(self.new_csv_linesL):
            if re.match('Time', row):
                log.info('Time heading found')
                return()
            elif isinstance(row[0], int):
                # valid Time value found
                return()

            # ERROR in new_csv_linesL - Time value is not a number
            # DUMP EVERYTHING AND HALT
            print(f'Invalid value in Time Column! Please send s2csv.log to developers.')
            print(f'To recover with minimal data loss,\n   - copy latest file in backups folder\n    - to ubit_live_data.csv')
            log.error(f'Invalid Time column detected in ROW {linenum}, CONTENT: {row}')
            self.dump_everything()
            log.error(f'Invalid Time column detected in ROW {linenum}, CONTENT: {row}')
            exit(-1)
        exit(f'new_csv_linesL is empty: {self.new_csv_linesL}')

    def dump_everything(self):
            log.error("Start of DATA DUMP - cause should be on the line above this")
            log.warning("Dump of new_csv_linesL")
            log.warning(f'new_csv_linesL: {self.new_csv_linesL}')
            log.warning(f'new_ubit_dataD: {self.new_ubit_dataD}')
            log.warning(f'csvLinesL: {self.csvLinesL}')
            log.warning(f'heading_colsL: {self.heading_colsL}')
            log.warning(f'next_data_row: {self.next_data_row}')
            # log.warning(f'keys: {self.__dict__.keys()}')
            log.warning("END of DATA DUMP - cause should be reported on the line below this")

    def csv_writelines(self, mypath=None):
        self.detect_invalid_time_column_debug_code()

        if mypath is None:
            mypath = self.csv_path
        keep_trying = 3
        while( keep_trying > 0 ):
            '''
            Writes can get blocked temporarily when Excel locks the file
            Try a few times before giving up
            '''
            try:
                with open(mypath, mode='w') as csvF:
                    keep_trying = 0
                    for row in self.new_csv_linesL:
                        csvF.write( row )
                mypath.chmod(0o0666)
            except Exception as e:
                keep_trying -= 1
                log.warning(f'failed write {keep_trying}')

    def csv_save_backup(self):
        '''
        Make a backup copy of CSV every NNN seconds
        Since the code could wipe out the source data
        Use file copy to create the backup
        '''
        if( time.time() > self.backup_last_saved + self.backup_interval_secs ):
            try:
                self.backup_last_saved = time.time()
                # Build path
                dt = datetime.datetime.fromtimestamp(time.time())
                suffix = dt.strftime(self.backup_file_suffix)
                backup_file_path = self.backup_folder_str + self.backup_file_prefix + suffix + '.csv'
                Path( self.backup_folder_str ).mkdir( exist_ok=True )
                shutil.copy2( self.csv_path, backup_file_path )
            except Exception as e:
                log.warning(f'CSV backup failed because: {e}')

    def csv_readlines(self, mypath=None):
        if mypath is None:
            mypath = self.csv_path
        if mypath.is_file():
            with open(mypath, mode='r') as csvF:
                self.csvLinesL = csvF.readlines()
        if (len(self.csvLinesL) == 0):
            self.csvLinesL = ['Time']

    def mk_headings_list(self, headingstr=None):
        '''
        turn first line of csv file
        into a List of headings

        args are only used for testing
        '''
        if headingstr is None:
            headingstr = self.csvLinesL[0]

        self.heading_colsL = headingstr.strip().split(',')

        if (self.heading_colsL[0] != 'Time'):
            self.heading_colsL.insert(0,'Time')

        
    def build_next_data_row(self, heading_colsL=None, new_ubit_dataD=None ):
        '''
        From heading_colsL and new_dataD
        Populate the next_data_row
        '''
        if (heading_colsL is None):
            heading_colsL = self.heading_colsL
        if (new_ubit_dataD is None):
            new_ubit_dataD = self.new_ubit_dataD

        self.next_data_row = []

        # Set the Time column value to current time in Excel Epochs
        dt = datetime.datetime.fromtimestamp(time.time())
        dt_excel = openpyxl.utils.datetime.to_excel(dt)
        self.next_data_row.append(dt_excel)

        for c in heading_colsL[1:]:
            if c in new_ubit_dataD:
                self.next_data_row.append(new_ubit_dataD[c])
            else:
                self.next_data_row.append('')

        log.info(f'heading_colsL={heading_colsL}')
        log.info(f'next_row={self.next_data_row}')
        # Build a dict for console output
        new_data = {k: v for k, v in zip(heading_colsL, self.next_data_row)}
        new_data.pop('Time', None)
        new_data.pop('C-DRIVE-DATA', None)
        print(f'Temperatures: {new_data}')
        
    def csv_add_new_headings(self, headingsL=None, dataD=None):
        '''
        Preflight: check minimal format of csv headings row
            First column = Time
            Don't care after that
        append to the csv headings row
            new uBit IDs (column names) from new_ubit_dataD
        '''    
        if dataD is None:
            dataD = self.new_ubit_dataD

        if headingsL is not None:
            self.heading_colsL = headingsL

        for k in dataD:
            if k not in self.heading_colsL:
                self.heading_colsL.append(k)

    def dequeue_into_dict(self, duration=0):
        '''
        For NN seconds (sampling interval for making a single CSV row)
            Pull microbit messages from queue 
            Return results in a dictionary of {microbit_name : value}
        duration param is only used for testing
        '''
        if duration == 0:
            duration = self.sampling_period_secs

        self.new_ubit_dataD = {}

        count = 0
        invalid_strings = []
        self.set_timer_start()
        while self.is_timeout(duration) is False:
            try:
                inline = ""
                inline = self.qu.get(timeout=0.1).strip()
                self.qu.task_done
                count = count+1
            except Empty:
                pass
                # ignore Empty Queue
                # keep waiting on the refresh_cycle_timeout

            try:
                (ubit_name, valuestr) = inline.split(':')
                value = float(valuestr)
                self.new_ubit_dataD[ubit_name] = value
            except Exception as e:
                # invalid data format (did not find a colon)
                invalid_strings.append(inline)
                continue
        log.info(f'Invalid strings from microbit: {invalid_strings}')
        log.info(f'qu elements processed={count}')
        log.info(f'Dump of dataD:\n{self.new_ubit_dataD}')

    def set_timer_start(self):
        self.prev_time = time.time()

    def is_timeout(self, delta_timeout):
        if time.time() >= self.prev_time + delta_timeout:
            return True
        else:
            return False

    def test_serial(self):
        ser = serial.Serial(self.serial_path, 115200, timeout = 1)
        count = 0
        self.set_timer_start()
        while True:
            data = ser.readline().decode('utf-8').strip()
            d1 = re.sub(' ','.',data)
            d2 = data.maketrans(" ",".")
            count += 1
            if self.is_timeout(2):
                self.set_timer_start()
                log.info(f'test_serial: {count} {data}')

    def test_serial_qu(self):
        self.infeed_mk_thread(self.infeed_serial)
        count = 0
        self.set_timer_start()
        while True:
            data = self.qu.get()
            self.qu.task_done()

            d1 = re.sub(' ','.',data)
            d2 = data.maketrans(" ",".")
            count += 1
            if self.is_timeout(2):
                self.set_timer_start()
                log.info(f'test_serial_qu: {count} {self.qu.qsize()} |{data}|')

    def infeed_mk_thread(self, infeed_method):
        worker = Thread(target=infeed_method)
        worker.setDaemon = True
        worker.start()
        return worker

    def infeed_serial(self):
        '''
        Thread for pulling data from serial port
        Includes a preamble of faked data
        '''
        if self.fake_data_enabled:
            for line in self.fake_ubit_data_str.splitlines():
                self.qu.put(line)
        try:
            self.ser = serial.Serial(self.serial_path, 115200, timeout=1)
            count = 0
            while self.continue_serial_read_thread is True:
                data = self.ser.readline().decode('utf-8').strip()
                self.qu.put(data)
        except Exception as e:
            log.warning(f'Serial Error ignored:\n      {e}\nSerial Port not available. Are you root? Is Microbit plugged in?\n')
            print(f'Serial Error ignored:\n      {e}\nSerial Port not available. Are you root? Is Microbit plugged in?\n')
            # raise(e)
            
    def env_detect(self):
        self.comports = comports()

        if os.name == self.os_name_ux:
            if os.geteuid() != 0:
                exit('On Linux systems you must run this script as root. Eg. sudo python ...')
            # Make a list of comport paths
            comport_paths = []
            for com in self.comports:
                comport_paths.append(com[0])

            # On Linux the Microbit comport has a very specific name
            if self.serial_path_ux not in comport_paths:
                exit("ux: Connect Microbit to USB before running this script.")
            self.serial_path = self.serial_path_ux
        
        elif os.name == self.os_name_win:
            '''
            find the desired comport
            '''
            if len(self.comports) == 0:
                exit("Win: Connect Microbit to USB before running this script.")
            elif len(self.comports) != 1:
                self.select_win_comport(self.comports)
            else:
                self.serial_path = self.comports[0][0]
        else:
            exit('TODO: Unrecognized OS "{os.name}"')

    def select_win_comport(self, comportsL):
        log.info('Multiple COM ports detected - prompting user')

        print("Which Port is the Microbit connected to?")
        print("It's safe to try any port until find the correct one")

        # Make a menu of COM IDs
        com_choices = {}
        for idx, port in enumerate(comportsL):
            com_choices[idx] = port
            print(f"{idx} - {port[0]} {port[1]}")
        log.info(f'com_choices={com_choices}')
        
        # User Chooses - use zero based choice list
        # NOTE: menu selections are integers NOT strings!
        user_choice = int(input("Enter your choice: "))
        log.info(f"User Chose |{user_choice}|")
        
        # The port number we need is field 0 in the selected Port record
        self.serial_path = com_choices[user_choice][0]
        log.info(f'User chosen comport: {user_choice} - {self.serial_path}')
        print(f'{self.serial_path} selected')

    def logger_setup(self):
        max_file_size = 1024 * 1024  # 1 MB
        max_backup_files = 2
        self.logfile_path = 's2csv.log'

        file_handler = RotatingFileHandler(
            self.logfile_path, maxBytes=max_file_size, backupCount=max_backup_files)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        global log
        log = logging.getLogger(__name__)
        log.setLevel(logging.INFO)
        log.addHandler(file_handler)
        return log

if __name__ == "__main__":
    ser2 = Serial2CSV()
    log = ser2.logger_setup()
    ser2.config1()
    ser2.env_detect()
    ser2.config2()
    log.info('STARTUP INFO for serial2csv')
    log.info(f'serialpath={ser2.serial_path}')
    log.info(f'sampling_period={ser2.sampling_period_secs} max_samples={ser2.sample_count}')
    serial_thread = ser2.infeed_mk_thread(ser2.infeed_serial)
    try:
        ser2.csv_update()
    except KeyboardInterrupt:
        ser2.continue_serial_read_thread = False
        serial_thread.join()



