# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
import datetime
import json
import logging
import logging.config
import math
from pathlib import Path

import win32com.client as com

import cal
import readinput
import report
import runsimul
import setvissim
from variable import *

# from tqdm import tqdm

datainfo = dict()
datainfo['random_seed'] = -1
datainfo['quick_mode'] = True
datainfo['simulation_time'] = 600
datainfo['vehicle_input_period'] = 900
datainfo['comment'] = ""
start_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# 1. Read Excel
logger.info("Reading an input file...")
readinput.read_json(datainfo, Path().absolute()/"resources/init.json")

try:
    excel = com.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb1 = excel.Workbooks.Open(datainfo['signal_xlsx'])
    readinput.read_signal(wb1, Signal, datainfo['simulation_time'])
    wb2 = excel.Workbooks.Open(datainfo['vehicle_input_xlsx'])
    readinput.read_vehicleinput(wb2, VehicleInput)
    wb3 = excel.Workbooks.Open(datainfo['vehicle_routes_xlsx'])
    readinput.read_static_vehicle_routes(wb3, Static_Vehicle_Routes)
    excel.Quit()

except Exception as e:
    print(e)

finally:
    wb1 = wb2 = wb3 = None
    excel = None

BreakAt = readinput.set_accum_break(Signal)

# 2. Set Vissim
setvissim.convert_signal_to_enum(Signal)
datainfo['random_seed'] = setvissim.set_randomseed(datainfo['random_seed'])

# Connecting the COM Server => Open a new Vissim Window:
logger.info("Setting Vissim...")
Vissim = com.Dispatch("Vissim.Vissim")

# Load a Vissim Network:
Vissim.LoadNet(datainfo['vissim_inpx'])

setvissim.check_sig_file(Vissim)

Link_TT = setvissim.get_travtm_info(Vissim)
node_nums = setvissim.get_all_node(Vissim)
setvissim.find_incoming_lane(Vissim, lanes_with_SH)

setvissim.set_Vissim(Vissim, datainfo)
setvissim.set_link_segment(Vissim)
setvissim.set_queue_counter(Vissim, lanes_with_SH)
setvissim.set_data_collection(Vissim, lanes_with_SH)
setvissim.set_vehicleinput(Vissim, datainfo, VehicleInput)
setvissim.set_static_vehicle_route(Vissim, Static_Vehicle_Routes)


# 3. Run Simulation
logger.info("Running simulation...")
Vissim.Simulation.RunSingleStep()

# Extract data per signal period
"""
break_at = pbar_update = 0
with tqdm(total=datainfo['simulation_time']) as pbar:
    runsimul.set_signal(Vissim, Signal, break_at)
    for break_at in BreakAt:
        Vissim.Simulation.SetAttValue('SimBreakAt', break_at)   # Set break_at
        Vissim.Simulation.RunContinuous()   # Run simulation until 'break_at'
        runsimul.set_signal(Vissim, Signal, break_at)   # Set signal

        pbar.update(break_at-pbar_update)
        pbar_update = break_at
    Vissim.Simulation.RunContinuous()
"""
break_at = 0
runsimul.set_signal(Vissim, Signal, break_at)
for break_at in runsimul.progressbar(BreakAt):
    Vissim.Simulation.SetAttValue('SimBreakAt', break_at)   # Set break_at
    Vissim.Simulation.RunContinuous()   # Run simulation until 'break_at'
    runsimul.set_signal(Vissim, Signal, break_at)   # Set signal
Vissim.Simulation.RunContinuous()

# Extract data per hour
hour_step = math.ceil(break_at/3600)
for hour in range(hour_step):
    time_str = str(hour + 1)
    runsimul.extract_from_datacollection_per_hour(Vissim, time_str, VehNum_hour, OccupRate_hour)
    runsimul.extract_from_queue_per_hour(Vissim, time_str, QStop_hour)
    runsimul.extract_from_travtm_per_hour(Vissim, time_str, AvgSpeed_hour)

# Close COM server:
logger.info("Closing Vissim...")
Vissim = None


# 4. Calculate overall data
logger.info("Calculating...")
SH_per_link = cal.cal_SH_per_link(lanes_with_SH)

cal.cal_occuprate_overall(OccupRate_hour, OccupRate_overall, datainfo['simulation_time'])
cal.cal_qstop_overall(QStop_hour, QStop_overall)
cal.cal_qstop_per_meter(QStop_hour, QStop_overall, lanes_with_SH)

network_filename, _extention = datainfo['vissim_inpx'].split('.')

linkseg_result = f'{network_filename}_Link Segment Results_001.att'
cal.extract_from_linkseg(linkseg_result, lanes_with_SH, Density_overall, DelayRel_overall, AvgSpeed_overall)

cal.prep_extract_from_node(hour_step, LOS_hour, EmissionCO_hour, EmissionVOC_hour)
if node_nums:     # If there was any node in Vissim network,
    node_result = f'{network_filename}_Node Results_001.att'
    cal.extract_from_node(node_result, node_nums, EmissionCO, EmissionVOC, LOS_hour, EmissionCO_hour, EmissionVOC_hour)


# 5. Report
logger.info("Reporting...")
try:
    excel = com.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False     # To merge cells
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets("Sheet1")

    report.print_simul_info(ws, datainfo)
    report.print_explanation(ws)
    report.print_overall(ws, lanes_with_SH, SH_per_link, node_nums, DelayRel_overall, Density_overall, AvgSpeed_overall, QStop_overall, OccupRate_overall, EmissionCO, EmissionVOC)
    report.print_hour(ws, lanes_with_SH, SH_per_link, Link_TT, node_nums, VehNum_hour, QStop_hour, OccupRate_hour, AvgSpeed_hour, LOS_hour, EmissionCO_hour, EmissionVOC_hour)

    ws.Columns(2).AutoFit()
    wb.SaveAs(str(Path().absolute()/f'output_{start_time}.xlsx'))
    excel.Quit()

except Exception as e:
    print(e)

finally:
    ws = None
    wb = None
    excel = None
