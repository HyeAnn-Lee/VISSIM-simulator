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

Path('./log').mkdir(parents=True, exist_ok=True)
config = json.load(open("./logger.json"))
logging.config.dictConfig(config)
logger = logging.getLogger(__name__)

start_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# 1. Read Excel
logger.info("Reading an input file...")
DataInfo = readinput.\
            read_json(Path().absolute()/"vissim_simulator/resources/init.json")

try:
    excel = com.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb1 = excel.Workbooks.Open(DataInfo.Signal)
    readinput.read_signal(wb1, Signal, DataInfo.simulation_time)
    wb2 = excel.Workbooks.Open(DataInfo.VehicleInput)
    readinput.read_vehicleinput(wb2, VehicleInput)
    excel.Quit()

except Exception as e:
    print(e)

finally:
    wb1 = None
    wb2 = None
    excel = None

BreakAt = readinput.set_accum_break(Signal)

# 2. Set Vissim
setvissim.convert_signal_to_enum(Signal)
DataInfo.RandomSeed = setvissim.set_randomseed(DataInfo.RandomSeed)

# Connecting the COM Server => Open a new Vissim Window:
logger.info("Setting Vissim...")
Vissim = com.Dispatch("Vissim.Vissim")

# Load a Vissim Network:
Vissim.LoadNet(DataInfo.VissimInput)

setvissim.check_sig_file(Vissim)

Link_TT = setvissim.get_travtm_info(Vissim)
node_nums = setvissim.get_all_node(Vissim)
setvissim.find_incoming_lane(Vissim, lanes_with_SH)

setvissim.set_Vissim(Vissim, DataInfo)
setvissim.set_link_segment(Vissim)
setvissim.set_queue_counter(Vissim, lanes_with_SH)
setvissim.set_data_collection(Vissim, lanes_with_SH)
setvissim.set_vehicleinput(Vissim, DataInfo, VehicleInput)


# 3. Run Simulation
logger.info("Running simulation...")
Vissim.Simulation.RunSingleStep()

# Extract data per signal period
break_at = 0
runsimul.set_signal(Vissim, Signal, break_at)
for break_at in BreakAt:
    Vissim.Simulation.SetAttValue('SimBreakAt', break_at)   # Set break_at
    Vissim.Simulation.RunContinuous()       # Run simulation until 'break_at'
    runsimul.set_signal(Vissim, Signal, break_at)               # Set signal
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

cal.cal_occuprate_overall(OccupRate_hour, OccupRate_overall, DataInfo.simulation_time)
cal.cal_qstop_overall(QStop_hour, QStop_overall)
cal.cal_qstop_per_meter(QStop_hour, QStop_overall, lanes_with_SH)

network_filename = Path(DataInfo.VissimInput).stem  # without extension (.inpx)

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

    report.print_simul_info(ws, DataInfo)
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
