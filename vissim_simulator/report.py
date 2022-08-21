# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import copy
import datetime
import json
import logging
import logging.config
from enum import Enum

config = json.load(open("resources/logger.json"))
logging.config.dictConfig(config)
logger = logging.getLogger(__name__)

row = 1


class Metric(Enum):
    Lane = 1
    Link = 2
    TT = 3
    Node = 4


def _print_text(ws, context):
    # Input
    # > 'ws'        : Excel worksheet.
    # > 'context'   : str.
    #
    # Set value of cell('row', 1) to 'context' and increases 'row'.

    global row
    ws.Cells(row, 1).Value = context
    row += 1

    return


def _fill_color(ws, color, from_row, from_col, to_row=None, to_col=None):
    # Input
    # > 'ws'        : Excel worksheet.
    # > 'color'     : int
    # > 'from_row'  : int
    # > 'from_col'  : int
    # > 'to_row'    : int
    # > 'to_col'    : int

    # Set color of cells from cell('from_row', 'from_col') to
    # cell('to_row', 'to_col') as 'color'.
    # 'color' = 19 : Ivory
    #           36 : Light Yellow
    #           38 : Rose

    if to_row is None:
        to_row = from_row
    if to_col is None:
        to_col = from_col
    ws.Range(ws.Cells(from_row, from_col), ws.Cells(to_row, to_col))\
      .Interior.ColorIndex = color

    return


def _align_row(ws):
    # Input
    # > 'ws'        : Excel worksheet.

    # Align contexts in 'row' to be horizontally center.
    ws.Rows(row).HorizontalAlignment = 3  # Center

    return


def _merge_for_link(ws, cur_row, col, num_SH):
    # Input
    # > 'ws'        : Excel worksheet.
    # > 'cur_row'   : int
    # > 'col'       : int
    # > 'num_SH'    : int

    # Output
    # new column number after merging.

    # Merge cells in case of link metric.
    # Caller function should update its col to return value.
    # This function is supposed to be called by _print_column_name() and
    # _print_row_item().

    fromCell = ws.Cells(cur_row, col)
    col += num_SH - 1
    toCell = ws.Cells(cur_row, col)
    ws.Range(fromCell, toCell).Merge()

    return col


def _print_column_name(ws, metric, column_name):
    # Input
    # > 'ws'            : Excel worksheet.
    # > 'metric'        : enum 'Metric'.
    # > 'column_name'   : 1D list

    # Print column names of result table and increase 'row'.
    # This function is supposed to be called from print_overall() and
    # _print_Metric().

    global row
    _align_row(ws)
    col = 2
    ws.Cells(row, col).Value = metric.name

    if metric == Metric.Lane:
        # 'column_name' : 1D list of (int(LinkNo), int(LaneNo), double(PosSH),
        #                             double(LinkLen))
        for linkNo, laneNo, *_ in column_name:
            col += 1
            ws.Cells(row, col).Value = f"'{linkNo} - {laneNo}"

    elif metric == Metric.Link:
        # 'column_name' : 1D list of (int(LinkNo), int(NumSH))
        for linkNo, num_SH in column_name:
            col += 1
            ws.Cells(row, col).Value = linkNo
            col = _merge_for_link(ws, row, col, num_SH)

    elif metric == Metric.TT:
        # 'column_name' : 1D list of (str(StartLink), str(EndLink))
        ws.Cells(row, col).Value = "Section"    # Overwrite.
        for startlink, endlink in column_name:
            col += 1
            ws.Cells(row, col).Value = f'{startlink} to {endlink}'

    elif metric == Metric.Node:
        # 'column_name' : 1D list of int(NodeNo)
        for NodeNo in column_name:
            col += 1
            ws.Cells(row, col).Value = NodeNo

    else:
        logger.error("_print_column_name() : Invalid [metric].")

    _fill_color(ws, 19, row, 2, row, col)   # Color table.

    # Draw border with solid line.
    ws.Range(ws.Cells(row, 2), ws.Cells(row, col)).Borders.LineStyle = 1

    row += 1

    return


def _print_row_item(ws, row_name, metric, list_1D, SH_per_link=None,
                    display_min=False):
    # Input
    # > 'ws'            : Excel worksheet.
    # > 'row_name'      : str.
    # > 'metric'        : enum 'Metric'.
    # > 'list_1D'       : 1D list.
    # > 'SH_per_link'   : 1D list of (int(LinkNo), int(NumSH)).
    #                       For link metric, 'SH_per_link' should be given.
    #                       Otherwise, 'SH_per_link' has no meaning.
    # > 'display_min'   : boolean.

    # This function is supposed to be called from _print_Metric() and
    # print_overall().

    def _display_minmax_value(list_of_tuple, number, find_max):
        # Input
        # > 'list_of_tuple' : 1D-list of tuple(number or str, int).
        # > 'number'        : int.
        # > 'find_max'      : boolean.
        #
        # Cells with the 'number' largest/smallest values will be colored by
        # _fill_color().

        # If the number of elements in 'list_of_tuple' is less then 'number',
        # color all cells.
        if number > len(list_of_tuple):
            for _, column in list_of_tuple:
                _fill_color(ws, 38, row, column)
            return

        list_of_tuple.sort(reverse=find_max)
        thsh = list_of_tuple[number - 1][0]     # list_of_tuple[number-1].item
        for value, column in list_of_tuple:
            if (find_max and value < thsh) or (not find_max and value > thsh):
                return
            _fill_color(ws, 38, row, column)

    global row
    _align_row(ws)

    # Insert 'row_name'.
    ws.Cells(row, 2).Value = row_name

    # Create 1D-list 'target_list'.
    if metric == Metric.Link:   # Modify list_1D in case of link metric.
        if SH_per_link is None:
            logger.error("_print_row_item(): "
                         + "SH_per_link must be given in case of link metric.")
        target_list = copy.deepcopy(list_1D)
        for index in range(len(SH_per_link) - 1, -1, -1):
            for _ in range(SH_per_link[index][1] - 1):
                target_list.insert(index + 1,
                                   99999999 if display_min
                                   else -99999999)  # Insert dummy value.
    else:   # lane, TT, Node
        target_list = list_1D

    # Create 1D-list 'to_find_max' of tuple (item, col)
    col = 2
    to_find_max = []
    for item in target_list:    # item : str or number
        col += 1
        if item == -1:
            ws.Cells(row, col).Value = 'None'
        else:
            ws.Cells(row, col).Value = item
            to_find_max.append((item, col))

    # Find minimum / maximum values and color them.
    _display_minmax_value(to_find_max, 3,
                          False if (metric == Metric.TT) or display_min
                          else True)

    # Merge cells in case of link metric.
    if metric == Metric.Link:
        col = 3
        for _, num_SH in SH_per_link:
            col = _merge_for_link(ws, row, col, num_SH) + 1

    # Increment 'row'.
    row += 1

    return


def print_simul_info(ws, data):
    # Input
    # > 'ws'        : Excel worksheet.
    # > 'data'      : Init().

    start_row = row
    _print_text(ws, "$ Simulation Info")

    _print_text(ws, "* Network File : ")
    ws.Cells(row - 1, 3).Value = data.VissimInput

    _print_text(ws, "* Signal : ")
    ws.Cells(row - 1, 3).Value = data.Signal

    _print_text(ws, "* Vehicle Input : ")
    ws.Cells(row - 1, 3).Value = data.VehicleInput

    _print_text(ws, "* Static Vehicle Routes : ")
    ws.Cells(row - 1, 3).Value = data.static_vehicle_routes

    _print_text(ws, "* Date : ")
    ws.Cells(row - 1, 3).Value = datetime.datetime.now().strftime("%c")

    _print_text(ws, "* Random Seed : ")
    ws.Cells(row - 1, 3).Value = data.RandomSeed

    _print_text(ws, "* Quick Mode : ")
    ws.Cells(row - 1, 3).Value = data.quick_mode

    _print_text(ws, "* Simulation time (sec) : ")
    ws.Cells(row - 1, 3).Value = data.simulation_time

    _print_text(ws, "* Comment : ")
    ws.Cells(row - 1, 3).Value = data.Comment

    _print_text(ws, "*")

    _fill_color(ws, 36, start_row, 1, row - 2, 1)

    return


def print_explanation(ws):
    # Input
    # > 'ws'        : Excel worksheet.

    start_row = row
    _print_text(ws, "$ Measurements")
    _print_text(ws, "* Delay : "
                    + "Total delay divided by total travel time of all "
                    + "vehicles in this link segment [%]")
    _print_text(ws, "* Density : Vehicle density [/km]")
    _print_text(ws, "* Emissions CO : Quantity of carbon monoxide [g]")
    _print_text(ws, "* Emissions VOC : "
                    + "Quantity of volatile organic compounds [g]")
    _print_text(ws, "* LOS : Level of service (A ~ F).")
    _print_text(ws, "* OccupRate : "
                    + "Share of time [0% ~ 100%] of the last simulation step, "
                    + "in which at least one data collection point of this "
                    + "data collection measurement was occupied.")
    _print_text(ws, "* QueueStop : "
                    + "The number of queue stops per meter. A queue stop "
                    + "counts when a vehicle that is directly upstream or "
                    + "within the queue length falls below the speed of the "
                    + "Begin attribute defined for the queue condition. [/m]")
    _print_text(ws, "* Speed : "
                    + "Average speed of vehicles passing through the section "
                    + "[km/h]")
    _print_text(ws, "*")

    _fill_color(ws, 36, start_row, 1, row - 2, 1)

    return


def print_overall(ws, lanes_with_SH, SH_per_link, node_nums, DelayRel_overall,
                  Density_overall, AvgSpeed_overall, QStop_overall,
                  OccupRate_overall, EmissionCO, EmissionVOC):
    # Input
    # > 'ws'            : Excel worksheet.
    # > 'lanes_with_SH' : 1D list of (int(LinkNo), int(LaneNo), double(PosSH)).
    # > 'SH_per_link'   : 1D list of (int(LinkNo), int(NumSH)).
    # > 'node_nums'       : 1D list of int(NodeNo) or empty list.
    # > 'DelayRel_overall'  : 1D list of floats.
    # > 'Density_overall'   : 1D list of floats.
    # > 'AvgSpeed_overall'  : 1D list of floats.
    # > 'QStop_overall'     : 1D list of non-negative numbers.
    # > 'OccupRate_overall' : 1D list of non-negative numbers.
    # > 'EmissionCO'        : 1D list of floats.
    # > 'EmissionVOC'       : 1D list of floats.

    start_row = row
    _print_text(ws, "$ Overall Results")

    _print_column_name(ws, Metric.Lane, lanes_with_SH)

    _print_row_item(ws, "Delay", Metric.Link, DelayRel_overall, SH_per_link)
    _print_row_item(ws, "Density", Metric.Link, Density_overall, SH_per_link)
    _print_row_item(ws, "Speed", Metric.Link, AvgSpeed_overall, SH_per_link,
                    display_min=True)
    _print_row_item(ws, "QueueStop", Metric.Link, QStop_overall, SH_per_link)
    _print_row_item(ws, "OccupRate", Metric.Lane, OccupRate_overall)

    _fill_color(ws, 19, start_row + 2, 2, row - 1, 2)
    _print_text(ws, "*")

    if node_nums:
        mid_row = row
        _print_column_name(ws, Metric.Node, node_nums)

        _print_row_item(ws, "Emissions CO",     Metric.Node, EmissionCO)
        _print_row_item(ws, "Emissions VOC",    Metric.Node, EmissionVOC)

        _fill_color(ws, 19, mid_row, 2, row - 1, 2)
        _print_text(ws, "*")

    _fill_color(ws, 36, start_row, 1, row - 2, 1)

    return


def print_hour(ws, lanes_with_SH, SH_per_link, Link_TT, node_nums, VehNum_hour,
               QStop_hour, OccupRate_hour, AvgSpeed_hour, LOS_hour,
               EmissionCO_hour, EmissionVOC_hour):
    # Input
    # > 'ws'            : Excel worksheet.
    # > 'lanes_with_SH' : 1D list of (int(LinkNo), int(LaneNo), double(PosSH),
    #                                 double(LinkLen)).
    # > 'SH_per_link'   : 1D list of (int(LinkNo), int(NumSH)).
    # > 'Link_TT'       : 1D list of (str(StartLink), str(EndLink)) or empty
    #                     list.
    # > 'node_nums'       : 1D list of int(NodeNo) or empty list.
    # > 'VehNum_hour'       : 2D-list of non-negative numbers.
    # > 'QStop_hour'        : 2D list of non-negative numbers.
    # > 'OccupRate_hour'    : 2D-list of non-negative numbers.
    # > 'AvgSpeed_hour'     : 2D-list of numbers.
    # > 'LOS_hour'          : 2D-list of str or empty list.
    # > 'EmissionCO_hour'   : 2D-list of float or empty list.
    # > 'EmissionVOC_hour'  : 2D-list of float or empty list.

    def _print_Metric(metric_name, metric, column_name, list_2D):
        # Input
        # > 'metric_name'   : str.
        # > 'metric'        : enum 'Metric'.
        # > 'column_name'   : 1D list.
        # > 'list_2D'       : 2D-list.

        _print_text(ws, metric_name)    # Print metric's name.
        _print_column_name(ws, metric, column_name)  # Print column name.

        # Print table contents.
        inter_row = row
        hour = 0
        for hour, list_1d in enumerate(list_2D):
            row_name = f'{hour}~{hour+1} hour'
            _print_row_item(ws, row_name, metric, list_1d, column_name)

        # The last one has to be overwritten.
        ws.Cells(row - 1, 2).Value = f'{hour}~END'

        _fill_color(ws, 19, inter_row, 2, row - 1, 2)   # Color table.
        _print_text(ws, "*")    # Print new line.

        return

    start_row = row
    _print_text(ws, "$ Per Hour Results")

    _print_Metric("* The Number of Vehicles", Metric.Lane, lanes_with_SH,
                  VehNum_hour)
    _print_Metric("* OccupRate", Metric.Lane, lanes_with_SH, OccupRate_hour)
    _print_Metric("* QueueStop", Metric.Link, SH_per_link, QStop_hour)

    if Link_TT:
        _print_Metric("* Speed", Metric.TT, Link_TT, AvgSpeed_hour)

    if node_nums:
        _print_Metric("* LOS", Metric.Node, node_nums, LOS_hour)
        _print_Metric("* Emissions CO", Metric.Node, node_nums, EmissionCO_hour)
        _print_Metric("* Emissions VOC", Metric.Node, node_nums,
                      EmissionVOC_hour)

    _fill_color(ws, 36, start_row, 1, row - 2, 1)

    return
