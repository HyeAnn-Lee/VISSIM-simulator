# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import json
import logging
import logging.config
from pathlib import Path

config = json.load(open("./logger.json"))
logging.config.dictConfig(config)
logger = logging.getLogger(__name__)


def _read_result_att(att_file):
    # Input
    # > 'att_file' : result att file object.
    #
    # Output
    # > 'column_name.split(';')' : 1D-list of str.
    #
    # Using the fact that all result att files have the same layout.

    # Find the line NOT starts with *.
    att_file.readline()         # Pass the first line.
    line = att_file.readline()  # Start from the second line.
    while '*' == line[0]:
        line = att_file.readline()
    # Here, 'line' starts with $.

    line = line.rstrip()                # Remove the last NULL character.
    column_name = line.split(':')[1]    # str containing column names.

    return column_name.split(';')       # Return 1D-list of column names.


def _str_2_float(string):
    # Input
    # > 'string' : str.
    #
    # Output
    # > float.

    if string == '':
        return 0
    return float(string)


def cal_SH_per_link(lanes_with_SH):
    # Input
    # > 'lanes_with_SH' : 1D list of (int, int, double, double)
    #
    # Output
    # > 'sh_per_link' : 1D list of (int, int)
    #
    # Calculate the number of signal heads per link.
    # Output is further used to merge Excel cells.

    sh_per_link = []

    # init
    link_track = lanes_with_SH[0][0]
    num_SH = 0

    # Fill 'sh_per_link'
    for linkNo, *_ in lanes_with_SH:
        if linkNo != link_track:    # If new link is met,
            sh_per_link.append((link_track, num_SH))  # Append prev link info.
            link_track = linkNo     # Update to current 'linkNo'.
            num_SH = 1              # Update to current 'num_SH'.
        else:       # If still the same link,
            num_SH += 1     # Increment the number of SH.

    sh_per_link.append((link_track, num_SH))    # Append the last link info.

    return sh_per_link


def cal_occuprate_overall(OccupRate_hour, OccupRate_overall, sim_len):
    # Input
    # > 'OccupRate_hour'    : 2D list of non-negative numbers.
    # > 'OccupRate_overall' : Empty list.
    # > 'sim_len' : int.

    # Calculate how many minutes the last time interval is.
    last_min = (sim_len - 1) % 3600 + 1

    # Fill 'OccupRate_overall'.
    for index_DC in range(len(OccupRate_hour[0])):
        OccupRate = 0
        for index_hour in range(len(OccupRate_hour)-1):
            OccupRate += 3600 * OccupRate_hour[index_hour][index_DC]
        OccupRate += last_min * OccupRate_hour[-1][index_DC]
        OccupRate_overall.append(OccupRate / sim_len)

    # 'OccupRate_overall' becomes 1D list of non-negative numbers.
    return


def cal_qstop_overall(QStop_hour, QStop_overall):
    # Input
    # > 'QStop_hour'    : 2D list of non-negative numbers.
    # > 'QStop_overall' : Empty list.

    # Fill 'QStop_overall'.
    for index_QC in range(len(QStop_hour[0])):
        qstop = 0
        for temp_1dlist in QStop_hour:
            qstop += temp_1dlist[index_QC]
        QStop_overall.append(qstop)

    # 'QStop_overall' becomes 1D list of non-negative numbers.
    return


def cal_qstop_per_meter(QStop_hour, QStop_overall, lanes_with_SH):
    # Input
    # > 'QStop_hour'    : 2D list of non-negative numbers.
    # > 'QStop_overall' : 1D list of non-negative numbers.
    # > 'lanes_with_SH' : 1D list of (int, int, double, double).
    #
    # Divide all values in 'QStop_hour' and 'QStop_overall'
    # by [length of each link].

    for index_hour in range(len(QStop_hour)):
        for index_lane in range(len(QStop_hour[index_hour])):
            QStop_hour[index_hour][index_lane] /= lanes_with_SH[index_lane][3]

    for index_lane in range(len(QStop_overall)):
        QStop_overall[index_lane] /= lanes_with_SH[index_lane][3]

    return


def extract_from_linkseg(file, lanes_with_SH, Density_overall,
                         DelayRel_overall, AvgSpeed_overall):
    # Input
    # > 'file'              : Absolute path of Link Segment Results att file.
    # > 'lanes_with_SH'     : 1D list of (int, int, double, double).
    # > 'Density_overall'   : Empty list.
    # > 'DelayRel_overall'  : Empty list.
    # > 'AvgSpeed_overall'  : Empty list.

    def _remove_unit(string):
        # Input
        # > 'string' : str.
        #
        # Output
        # > str.

        return string.split(" ")[0]

    if not Path(file).exists():
        logger.error("extract_from_linkseg():\
                        Link Segment Results att file is missing.")

    # Read att file to find column names.
    att_file = open(file, "r")
    parse = _read_result_att(att_file)  # 'parse' : 1D-list of str.
    pLinkname = parse.index('LINKEVALSEGMENT')    # int
    pDensity = parse.index('DENSITY(ALL)')       # int
    pDelayRel = parse.index('DELAYREL(ALL)')      # int
    pSpeed = parse.index('SPEED(ALL)')         # int

    # Find links with signal heads.
    links_with_SH = []
    for linkNo, *_ in lanes_with_SH:
        if linkNo not in links_with_SH:
            links_with_SH.append(linkNo)
    # 'links_with_SH' : 1D list of int.

    # Read the rest of 'att_file' and fill 'Density_overall',
    # 'DelayRel_overall' and 'AvgSpeed_overall'.
    i = 0
    while i < len(links_with_SH):
        line = att_file.readline().rstrip()
        parse = line.split(';')     # 1D-list of str containing actual data.

        # If 'line' is about links with signal head,
        if parse[pLinkname].split('-')[0] == str(links_with_SH[i]):
            density = _str_2_float(_remove_unit(parse[pDensity]))
            delayrel = _str_2_float(_remove_unit(parse[pDelayRel]))
            speed = _str_2_float(_remove_unit(parse[pSpeed]))
            Density_overall.    append(density)
            DelayRel_overall.   append(delayrel)
            AvgSpeed_overall.   append(speed if speed else -1)
            i += 1
    # 'Density_overall', 'DelayRel_overall' and 'AvgSpeed_overall' becomes
    # 1D list of floats. -1 value means that actual data was 0.

    att_file.close()
    return


def prep_extract_from_node(hour_step, LOS_hour, EmissionCO_hour,
                           EmissionVOC_hour):
    # Input
    # > 'hour_step'         : int.
    # > 'LOS_hour'          : Empty list.
    # > 'EmissionCO_hour'   : Empty list.
    # > 'EmissionVOC_hour'  : Empty list.

    LOS_hour.           extend([[] for _ in range(hour_step)])
    EmissionCO_hour.    extend([[] for _ in range(hour_step)])
    EmissionVOC_hour.   extend([[] for _ in range(hour_step)])

    # 'LOS_hour', 'EmissionCO_hour' and 'EmissionVOC_hour' become
    # empty 2D-lists.

    return


def extract_from_node(file, node_nums, EmissionCO, EmissionVOC, LOS_hour,
                      EmissionCO_hour, EmissionVOC_hour):
    # Input
    # > 'file'          : Absolute path of Node Results att file.
    # > 'node_nums'       : 1D list of int.
    # > 'EmissionCO'    : Empty list.
    # > 'EmissionVOC'   : Empty list.
    # > 'LOS_hour'          : Empty 2D-list.
    # > 'EmissionCO_hour'   : Empty 2D-list.
    # > 'EmissionVOC_hour'  : Empty 2D-list.

    if not Path(file).exists():
        logger.error("extract_from_node() : Node Results att file is missing.")

    num_nodes = len(node_nums)

    # Change form of each list first.
    EmissionCO. extend([0.0 for _ in range(num_nodes)])
    EmissionVOC.extend([0.0 for _ in range(num_nodes)])

    for index_hour in range(len(LOS_hour)):
        LOS_hour[index_hour] = ["" for _ in range(num_nodes)]
        EmissionCO_hour[index_hour] = [0.0 for _ in range(num_nodes)]
        EmissionVOC_hour[index_hour] = [0.0 for _ in range(num_nodes)]

    # Read att file to find column names.
    att_file = open(file, "r")
    parse = _read_result_att(att_file)  # 'parse' : 1D-list of str.
    pTimeInt = parse.index('TIMEINT')        # int
    pMovement = parse.index('MOVEMENT')       # int
    pLOS = parse.index('LOS(ALL)')       # int
    pCO = parse.index('EMISSIONSCO')    # int
    pVOC = parse.index('EMISSIONSVOC')   # int

    # Read the rest of 'att_file'.
    collection = []
    line = att_file.readline().rstrip()
    while line:
        parse = line.split(';')     # 1D-list of str containing actual data.
        # If 'line' contains aggregated data of a node,
        if '@' not in parse[pMovement]:
            collection.append(parse)
        line = att_file.readline().rstrip()
    # 'collection' becomes 2D list of str.

    att_file.close()

    # Fill 'EmissionCO', 'EmissionVOC', 'LOS_hour', 'EmissionCO_hour' and
    # 'EmissionVOC_hour'
    for parse in collection:
        CO = _str_2_float(parse[pCO])
        VOC = _str_2_float(parse[pVOC])

        node_index = node_nums.index(int(parse[pMovement].split(':')[0]))
        EmissionCO[node_index] += CO
        EmissionVOC[node_index] += VOC

        hour_index = int(parse[pTimeInt].split('-')[0]) // 3600
        LOS_hour[hour_index][node_index] = parse[pLOS][-1]
        EmissionCO_hour[hour_index][node_index] = CO
        EmissionVOC_hour[hour_index][node_index] = VOC

    # 'EmissionCO' and 'EmissionVOC' become 1D-list of float.
    # 'LOS_hour' becomes 2D-list of str.
    # 'EmissionCO_hour' and 'EmissionVOC_hour' become 2D-list of float.

    return
