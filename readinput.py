# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import json
from collections    import namedtuple
from warning     import *

class Init:
    def __init__(self):
        self.Signal = ""        # string : absolute path of signal Excel file.
        self.VehicleInput = ""  # string : absolute path of vehicle input Excel file.
        self.VissimInput = ""   # string : absolute path of Vissim inpx file.
        self.RandomSeed = -1    # int
        self.TimeInterval = 900 # int
        self.Comment = ""       # string

class SigControl:
    def __init__(self, name):
        self.Name = name    # string
        self.SigInd = []    # 2D-list of characters 'R', 'G' or 'Y'.
        self.BreakAt = []   # 1D-list of int.

class VehInput:
    def __init__(self, timeint):
        self.TimeInt = timeint  # positive int
        self.VehInfo = []       # elements will be namedtuple 'LinkInfo'.


def read_json(filename):
    # Input
    # > 'filename' : Absolute path of json file.
    #
    # Output
    # > 'data' : Init().

    with open(filename, "r", encoding = 'UTF8') as init_json:
        init_python = json.load(init_json)

    data = Init()
    data.Signal         = init_python['TargetFile']['Signal']
    data.VehicleInput   = init_python['TargetFile']['VehicleInput']
    data.VissimInput    = init_python['TargetFile']['VissimInput']

    data.RandomSeed     = init_python['Settings']['RandomSeed']
    data.TimeInterval   = init_python['Settings']['TimeInterval of VehicleInput']
    data.Comment        = init_python['Settings']['Comment']

    # Validation check
    if not isinstance(data.RandomSeed, int):
        warning(" ERROR : RandomSeed should be an integer... Check json file again.")
    if not isinstance(data.TimeInterval, int):
        warning(" ERROR : TimeInterval should be an positive integer... Check json file again.")

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

        column = 3

        while ws.Cells(1, column).Value:
            # Each element of 'SigInd' will contain signal information ('R', 'G', 'Y')
            # from all "signal group"s in one signal step.

            sigcon.SigInd.append([])

            row = 1
            while isinstance(ws.Cells(row, column).Value, str):
                value = ws.Cells(row, column).Value
                if value not in ['R', 'G', 'Y']:
                    warning(" ERROR : Invalid signal from xlsx... You must use either 'R', 'G' or 'Y'.")
                sigcon.SigInd[-1].append(value)
                row += 1
            column += 1

        return

    def _read_signal_time(sigcon):
        # Input
        #   'sigcon' : SigControl() with self.Name and self.SigInd.
        #
        # Read signal information of '현시 시간 배분' table.

        row = len(sigcon.SigInd[0]) + 1
        column = 3

        accTime = 0
        sigcon.BreakAt.append(accTime)

        while ws.Cells(row, column).Value:
            value = ws.Cells(row, column).Value
            if not isinstance(value, float):
                warning(" ERROR : You should use an integer for signal time...")
            if (int(value) != value) or (value < 1):
                warning(" ERROR : Signal time should be a positive integer...")

            accTime += value
            sigcon.BreakAt.append(accTime)
            column += 1
            if column - 3 == len(sigcon.SigInd):
                row += 1
                column = 3

        return

    for i in range(wb.Worksheets.Count):
        ws = wb.Worksheets(i+1)
        sigcontrol = SigControl(ws.name)    # SigControl.Name
        _read_signal_seq(sigcontrol)        # SigControl.SigInd
        _read_signal_time(sigcontrol)       # SigControl.BreakAt
        Signal.append(sigcontrol)

    # 'Signal' becomes a 1D-list of SigControl().
    if len(Signal) == 0:
        warning(" ERROR : Signal file is empty.... Check json file again.")

    sim_len = Signal[0].BreakAt[-1]
    for sigcontrol in Signal:
        if sigcontrol.BreakAt[-1] != sim_len:
            warning(" ERROR : Simulation time of each sheet of excel should be same...")

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
        while ws.Cells(2, num_vehcomp + 2).Value:
            num_vehcomp += 1

        row = 3
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

    for i in range(wb.Worksheets.Count):
        vehinput = VehInput(i + 1)                      # set VehInput.TimeInt
        _set_vehinfo(wb.WorkSheets(i + 1), vehinput)    # set VehInput.VehInfo
        VehicleInput.append(vehinput)                   # add vehinput

    # 'VehicleInput' becomes a 1D-list of VehInput().
    if len(VehicleInput) == 0:
        warning(" ERROR : VehicleInput file is empty.... Check json file again.")
    num_link = len(VehicleInput[0].VehInfo)
    for vehinput in VehicleInput:
        if len(vehinput.VehInfo) != num_link:
            warning(" ERROR : The number of links in VehicleInput Excel file is different in some sheets... Check the file again.")

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
        warning(" ERROR : Simulation time is zero... Check signal Excel file again.")

    return accum_break