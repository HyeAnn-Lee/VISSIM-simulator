# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import json
import logging
import logging.config
import sys
from collections import namedtuple

config = json.load(open("resources/logger.json"))
logging.config.dictConfig(config)
logger = logging.getLogger(__name__)


class SigControl:
    def __init__(self, name, offset_info):
        self.Name = name        # string
        self.offset_info = offset_info  # (Int, Int). (offset[sec], main 현시)
        self.SigInd = []        # 2D-list of characters 'R', 'G' or 'Y'.
        self.signal_time = []   # 1D-list of int.
        self.BreakAt = []       # 1D-list of int.


class VehInput:
    def __init__(self, timeint):
        self.TimeInt = timeint  # positive int
        self.VehInfo = []       # elements will be namedtuple 'LinkInfo'.


NUM_DISCRIPTION_LINE = 3


def read_json(datainfo, filename):
    # Input
    # > 'datainfo' : dict.
    # > 'filename' : Absolute path of json file. <class 'pathlib.WindowsPath'>.
    #
    # Output
    # > 'data' : Init().

    with filename.open('r', encoding='UTF8') as init_json:
        data_dict = json.load(init_json)

    comp1 = data_dict['TargetFile']
    comp2 = data_dict['Settings']

    datainfo['signal_xlsx'] = comp1['Signal']
    datainfo['vehicle_input_xlsx'] = comp1['VehicleInput']
    datainfo['vissim_inpx'] = comp1['VissimInput']
    datainfo['vehicle_routes_xlsx'] = comp1['Static Vehicle Routes']

    datainfo['random_seed'] = comp2['RandomSeed']
    datainfo['quick_mode'] = comp2['Quick Mode']
    datainfo['simulation_time'] = comp2['Simulation period [sec]']
    datainfo['vehicle_input_period'] = comp2['TimeInterval of VehicleInput']
    datainfo['comment'] = comp2['Comment']

    if not isinstance(datainfo['random_seed'], int):
        logger.error(
            "RandomSeed should be an integer. Check json file again.")
    if not isinstance(datainfo['simulation_time'], int):
        logger.error(
            "Simulation period should be a positive integer.",
            "Check json file again.")
    if not isinstance(datainfo['vehicle_input_period'], int):
        logger.error(
            "TimeInterval of VehicleInput should be a positive integer.",
            "Check json file again.")

    return


def read_signal_xlsx(wb, Signal):
    # Input
    # > 'wb' : Excel file with contents of signal information.
    # > 'Signal' : Empty list.

    def _read_signal_seq(sigcon):
        # Input
        #   'sigcon' : SigControl() with self.Name and self.offset_info.
        #
        # Read signal information of '현시' table.

        sg_nums = []    # Signal group No's.

        # Read signal group No.
        # Column B
        row = NUM_DISCRIPTION_LINE + 1
        while isinstance(ws.Cells(row, 3).Value, str):  # G, Y, R
            sg_no = ws.Cells(row, 2).Value  # Signal group No.
            sg_nums.append(int(sg_no))
            row += 1

        # ex1) sg_nums = [7, 1, 4, 8, 2, 3, 5, 6]
        # ex2) sg_nums = [6, 1, 2, 4, 7, 8]

        # Read and store signal information.
        # Column C ~
        column = 3
        while ws.Cells(NUM_DISCRIPTION_LINE + 1, column).Value:
            # Each element of 'SigInd' will contain signal information
            # ('R', 'G', 'Y') from all "signal group"s in one signal step.

            sigind = [None] * max(sg_nums)

            for row in range(len(sg_nums)):
                value = ws.Cells(NUM_DISCRIPTION_LINE + row + 1, column).Value

                # Break if signal time met.
                if not isinstance(value, str):
                    break

                if value not in ['R', 'G', 'Y']:
                    logger.error("_read_signal_seq():\t"
                                 + "Invalid signal from xlsx. "
                                 + "You must use either 'R', 'G' or 'Y'.")
                sigind[sg_nums[row]-1] = value

            # ex1) sigind = ['R', 'R', 'R', 'R', 'R', 'R', 'G', 'R']
            # ex2) sigind = ['R', 'G', None, 'R', None, 'R', 'G', 'R']

            sigcon.SigInd.append(sigind)
            column += 1

        return

    def _read_signal_time(sigcon):
        # Input
        #   'sigcon' : SigControl() with self.Name and self.SigInd.
        #
        # Read signal information of '현시 시간 배분' table.

        actual_sg = len(sigcon.SigInd[0]) - sigcon.SigInd.count(None)
        row = NUM_DISCRIPTION_LINE + actual_sg + 1

        column = 3

        while ws.Cells(row, column).Value:
            time = ws.Cells(row, column).Value
            if (not isinstance(time, float)) or (not time.is_integer()):
                logger.error("_read_signal_time():\t"
                             + "Signal time should be non-negative integers.")

            sigcon.signal_time.append(time)
            column += 1
            if column - 3 == len(sigcon.SigInd):
                row += 1
                column = 3

        offset = sigcon.offset_info[0]
        period = sum(sigcon.signal_time[:len(sigcon.SigInd)])
        if offset > period:
            logger.error("at _read_signal_time():\t"
                         + f"Signal offset ({offset} sec) is larger than "
                         + f"signal period ({int(period)} sec). "
                         + f"Check '{sigcon.Name}' sheet again.")

        return

    try:
        num_worksheets = wb.Worksheets.Count
        num_intersections = num_worksheets - 1

        # signal offset of each intersection
        offset_info = dict()

        # Sheet1
        ws = wb.Worksheets(1)
        # Column B ~
        for col in range(2, num_intersections+2):
            name = ws.Cells(NUM_DISCRIPTION_LINE + 1, col).Value
            offset = int(ws.Cells(NUM_DISCRIPTION_LINE + 2, col).Value)
            main_signal = int(ws.Cells(NUM_DISCRIPTION_LINE + 3, col).Value)
            offset_info[name] = (offset, main_signal)  # SigControl.offset_info

        # Sheet2 ~
        for i in range(2, num_worksheets+1):
            ws = wb.Worksheets(i)

            # SigControl.Name & .offset_info
            sigcontrol = SigControl(ws.name, offset_info[ws.name])

            # SigControl.SigInd
            _read_signal_seq(sigcontrol)

            # SigControl.signal_time
            _read_signal_time(sigcontrol)

            Signal.append(sigcontrol)

    except Exception as e:
        print('\t', sys.exc_info()[0], e)

    finally:
        ws = None

    # 'Signal' becomes a 1D-list of SigControl().
    if not Signal:
        logger.error("read_signal():\t"
                     + "Signal file is empty. Check json file again.")

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
        logger.error("read_vehicleinput():\t"
                     + "VehicleInput file is empty. Check json file again.")
    num_link = len(VehicleInput[0].VehInfo)
    for vehinput in VehicleInput:
        if len(vehinput.VehInfo) != num_link:
            logger.error("read_vehicleinput():\t"
                         + "The number of links in VehicleInput Excel file is "
                         + "different in some sheets. Check the file again.")

    return


def read_static_vehicle_routes(wb, Static_Vehicle_Routes):
    try:
        ws = wb.Worksheets(1)

        # cell A1: "$VISION"
        # cell A2: "* File: ..."
        row = 2
        while ws.Cells(row, 1).Value.startswith('*'):
            row += 1

        # cell A{row}: "$VEHICLEROUTESTATIC ..."
        col = 1
        column_names = []
        while True:
            value = ws.Cells(row, col).Value
            if not value:
                break
            column_names.append(value.split(':')[-1])
            col += 1

        Static_Vehicle_Routes.append(tuple(column_names))
        row += 1

        # cell A{row}: "6", B{row}: "1", ...
        data = []
        while ws.Cells(row, 1).Value:
            single_route = []
            for col in range(1, len(column_names)+1):
                value = ws.Cells(row, col).Value
                single_route.append(str(value))
            data.append(tuple(single_route))
            row += 1

        Static_Vehicle_Routes.append(data)

    except Exception as e:
        print('\t', sys.exc_info()[0], e)

    finally:
        ws = None

    # Static_Vehicle_Routes ==
    # [
    #  ('VEHROUTDEC', 'NO', 'NAME', 'FORMULA', 'DESTLINK', 'DESTPOS', 'RELFLOW(1)'),
    #  [('6.0', '1.0', 'None', 'None', '10028.0', '9.649', '3.0'),
    #   ('6.0', '2.0', 'None', 'None', '10027.0', '9.763', '6.0'),
    #   ('6.0', '3.0', 'None', 'None', '10029.0', '10.284', '1.0'),
    #   ('7.0', '1.0', 'None', 'None', '10025.0', '11.319', '3.0'),
    #   ('7.0', '2.0', 'None', 'None', '10024.0', '12.447', '6.0'),
    #   ('7.0', '3.0', 'None', 'None', '10026.0', '14.177', '1.0'),
    #   ('9.0', '1.0', 'None', 'None', '10042.0', '15.812', '6.0'),
    #   ...]
    # ]

    return


def rearrange_Signal(Signal):
    # Input
    # > 'Signal' : 1D-list of SigControl().
    #
    # Rearrange 'Signal' with signal offset.

    for sigcontrol in Signal:
        offset, main_signal = sigcontrol.offset_info
        if offset == 0:
            continue

        period1 = sigcontrol.signal_time[:len(sigcontrol.SigInd)]
        i = main_signal*2-2

        i -= 1
        while offset > period1[i]:
            offset -= period1[i]
            i -= 1

        period1[i] = offset
        sigcontrol.signal_time[:len(sigcontrol.SigInd)] = period1[i:]

        sigcontrol.SigInd = sigcontrol.SigInd[i:] + sigcontrol.SigInd[:i]

    return


def calculate_breakpoint(Signal, simulation_len):
    # Input
    # > 'Signal' : 1D-list of SigControl().
    # > 'simulation_len' : simulation period in seconds.
    #
    # Output
    # > 1D-list of int.
    #
    # Calculate and aggregate breakpoints of SigControls.

    accum_break = []

    for sigcontrol in Signal:
        sum = 0
        sigcontrol.BreakAt = [sum]
        for sec in sigcontrol.signal_time:
            sum += sec
            sigcontrol.BreakAt.append(sum)
        accum_break += sigcontrol.BreakAt

    accum_break = list(set(accum_break))
    accum_break.sort()
    accum_break.remove(0)
    if not accum_break:
        logger.error("calculate_breakpoint():\t"
                     + "Simulation time is 0. Check signal Excel file again.")

    return [i for i in accum_break if i < simulation_len]
