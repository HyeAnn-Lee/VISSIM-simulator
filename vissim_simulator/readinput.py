# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import json
import logging
import logging.config
import sys
from collections import namedtuple

config = json.load(open("./logger.json"))
logging.config.dictConfig(config)
logger = logging.getLogger(__name__)

class Init:
    def __init__(self):
        self.Signal = ""        # string: absolute path of signal Excel file.
        self.VehicleInput = ""  # string: abs path of vehicle input Excel file.
        self.VissimInput = ""   # string: absolute path of Vissim inpx file.
        self.RandomSeed = -1    # int
        self.TimeInterval = 900     # int
        self.Comment = ""       # string


class SigControl:
    def __init__(self, name):
        self.Name = name    # string
        self.SigInd = []    # 2D-list of characters 'R', 'G' or 'Y'.
        self.BreakAt = []   # 1D-list of int.
        self.total_simulation = 0    # int


class VehInput:
    def __init__(self, timeint):
        self.TimeInt = timeint  # positive int
        self.VehInfo = []       # elements will be namedtuple 'LinkInfo'.


NUM_DISCRIPTION_LINE = 3


def read_json(filename):
    # Input
    # > 'filename' : Absolute path of json file. <class 'pathlib.WindowsPath'>.
    #
    # Output
    # > 'data' : Init().

    with filename.open('r', encoding='UTF8') as init_json:
        init_python = json.load(init_json)

    data = Init()
    data.Signal = init_python['TargetFile']['Signal']
    data.VehicleInput = init_python['TargetFile']['VehicleInput']
    data.VissimInput = init_python['TargetFile']['VissimInput']

    data.RandomSeed = init_python['Settings']['RandomSeed']
    data.TimeInterval = init_python['Settings']['TimeInterval of VehicleInput']
    data.Comment = init_python['Settings']['Comment']

    # Validation check
    if not isinstance(data.RandomSeed, int):
        logger.error("read_json() : RandomSeed should be an integer... Check json file again.")
    if not isinstance(data.TimeInterval, int):
        logger.error("read_json() : TimeInterval should be an positive integer... Check json file again.")

    return data


def read_signal(wb, Signal):
    # Input
    # > 'wb' : Excel file with contents of signal information.
    # > 'Signal' : Empty list.

    def _read_signal_seq(sigcon):
        # Input
        #   'sigcon' : SigControl() with self.Name.
        #
        # Read signal information of '현시' table.

        sg_nums = []

        # Read signal group No.
        row = NUM_DISCRIPTION_LINE + 1
        while isinstance(ws.Cells(row, 3).Value, str):
            sg_no = ws.Cells(row, 2).Value
            sg_nums.append(int(sg_no))
            row += 1
        sg_nums = tuple(sg_nums)

        # Read and store signal information.
        column = 3
        while ws.Cells(NUM_DISCRIPTION_LINE + 1, column).Value:
            # Each element of 'SigInd' will contain signal information
            # ('R', 'G', 'Y') from all "signal group"s in one signal step.

            sigind = [None] * len(sg_nums)

            for row in range(1, len(sg_nums)+1):
                value = ws.Cells(NUM_DISCRIPTION_LINE + row, column).Value
                if value not in ['R', 'G', 'Y']:
                    logger.error("_read_signal_seq() : Invalid signal from xlsx... You must use either 'R', 'G' or 'Y'.")
                sigind[sg_nums[row-1]-1] = value
            sigcon.SigInd.append(sigind)
            column += 1

        return

    def _read_signal_time(sigcon, offset):
        # Input
        #   'sigcon' : SigControl() with self.Name and self.SigInd.
        #   'offset' : offset of each intersection
        #
        # Read signal information of '현시 시간 배분' table.

        row = NUM_DISCRIPTION_LINE + len(sigcon.SigInd[0]) + 1
        column = 3

        period = 0
        accTime = offset
        sigcon.BreakAt.append(accTime)

        while ws.Cells(row, column).Value:
            time = ws.Cells(row, column).Value
            if not isinstance(time, float):
                logger.error("_read_signal_time() : You should use an integer for signal time...")
            if (int(time) != time) or (time < 1):
                logger.error("_read_signal_time() : Signal time should be a positive integer...")

            period += time
            accTime += time
            sigcon.BreakAt.append(accTime)
            column += 1
            if column - 3 == len(sigcon.SigInd):
                if offset > period:
                    logger.error(
                        "at _read_signal_time():\t"
                        + f"Signal offset ({offset} sec) is larger than "
                        + f"signal period ({int(period)} sec). "
                        + f"Check '{sigcon.Name}' sheet again.")
                period = 0
                row += 1
                column = 3

        sigcon.total_simulation = int(accTime)

        return

    try:
        num_worksheets = wb.Worksheets.Count
        num_intersections = num_worksheets - 1

        offset_info = dict()    # signal offset of each intersection
        ws = wb.Worksheets(1)
        for col in range(2, num_intersections+2):
            name = ws.Cells(NUM_DISCRIPTION_LINE + 1, col).Value
            offset = int(ws.Cells(NUM_DISCRIPTION_LINE + 2, col).Value)
            offset_info[name] = offset

        for i in range(1, num_intersections+1):
            ws = wb.Worksheets(i+1)

            # SigControl.Name
            sigcontrol = SigControl(ws.name)
            # SigControl.SigInd
            _read_signal_seq(sigcontrol)
            # SigControl.BreakAt
            _read_signal_time(sigcontrol, offset_info[ws.name])

            Signal.append(sigcontrol)

    except Exception as e:
        print('\t', sys.exc_info()[0], e)

    finally:
        ws = None

    # 'Signal' becomes a 1D-list of SigControl().
    if not Signal:
        logger.error("read_signal() : Signal file is empty.... Check json file again.")

    sim_len = Signal[0].total_simulation
    for sigcontrol in Signal:
        if sigcontrol.total_simulation != sim_len:
            logger.error("read_signal() : Simulation time of each sheet of excel should be same...")

    return


def read_vehicleinput(wb, VehicleInput):
    # Input
    # > 'wb' : Excel file with contents of vehicle input information.
    # > 'VehicleInput' : Empty list.

    def _set_vehinfo(ws, vehin):
        # Input
        # > 'ws' : Excel worksheet.
        # > 'vehin' : VehInput() with self.TimeInt.

        num_vehcomp = 0
        while ws.Cells(NUM_DISCRIPTION_LINE + 2, num_vehcomp + 2).Value:
            num_vehcomp += 1

        row = NUM_DISCRIPTION_LINE + 3
        while ws.Cells(row, 1).Value:
            # linkinfo.LinkNo   : int
            LinkNo = int(ws.Cells(row, 1).Value)

            # linkinfo.VehComp  : 1D-tuple of positive floats
            temp_list = []
            for column in range(2, num_vehcomp + 2):
                volume = ws.Cells(row, column).Value
                if not isinstance(volume, (int, float)):
                    volume = 0
                temp_list.append(volume)
            VehComp = tuple(temp_list)

            # add linkInfo to VehInfo
            vehin.VehInfo.append(LinkInfo(LinkNo, VehComp))
            row += 1

        return

    LinkInfo = namedtuple('namedtuple_linkinfo', ['LinkNo', 'VehComp'])

    try:
        for i in range(wb.Worksheets.Count):
            # set VehInput.TimeInt
            vehinput = VehInput(i + 1)

            # set VehInput.VehInfo
            ws = wb.WorkSheets(i + 1)
            _set_vehinfo(ws, vehinput)

            # add vehinput
            VehicleInput.append(vehinput)

    except Exception as e:
        print(e)

    finally:
        ws = None

    # 'VehicleInput' becomes a 1D-list of VehInput().
    if not VehicleInput:
        logger.error("read_vehicleinput() : VehicleInput file is empty.... Check json file again.")
    num_link = len(VehicleInput[0].VehInfo)
    for vehinput in VehicleInput:
        if len(vehinput.VehInfo) != num_link:
            logger.error("read_vehicleinput() : The number of links in VehicleInput Excel file is different in some sheets... Check the file again.")

    return


def set_accum_break(list_of_SigControl):
    # Input
    # > 'list_of_SigControl' : 1D-list of SigControl().
    #
    # Output
    # > 'accum_break' : 1D-list of int.
    #
    # Aggregate breakpoints of all SigControls.

    accum_break = []
    for sigcon in list_of_SigControl:
        accum_break += sigcon.BreakAt
    temp_set = set(accum_break)
    accum_break = list(temp_set)
    accum_break.sort()
    accum_break.remove(0)
    if not accum_break:
        logger.error("set_accum_break() : Simulation time is zero...\
                                        Check signal Excel file again.")

    return accum_break
