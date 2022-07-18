# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
from variable    import *
from readinput   import *
from setvissim   import *
from runsimul    import *
from cal         import *
from report      import *
import win32com.client as com
import datetime, os, math


start_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# 1. Read Excel
print("Reading an input file...")
DataInfo = read_json(os.getcwd() + "\\init.json")

excel = com.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
read_signal         (excel.Workbooks.Open(DataInfo.Signal),         Signal)
read_vehicleinput   (excel.Workbooks.Open(DataInfo.VehicleInput),   VehicleInput)
excel.Quit()

BreakAt = set_accum_break(Signal)

# 2. Set Vissim
convert_signal_to_enum(Signal)
DataInfo.RandomSeed = set_randomseed(DataInfo.RandomSeed)

## Connecting the COM Server => Open a new Vissim Window:
print("Setting Vissim...")
Vissim = com.Dispatch("Vissim.Vissim")

## Load a Vissim Network:
Vissim.LoadNet      (DataInfo.VissimInput)

check_sig_file      (Vissim)

Link_TT = get_travtm_info   (Vissim)
No_Node = get_node_info     (Vissim)
find_incoming_lane  (Vissim, lanes_with_SH)

SimLen = BreakAt.pop(-1)
set_Vissim          (Vissim, SimLen, DataInfo.RandomSeed)
set_link_segment    (Vissim)
set_queue_counter   (Vissim, lanes_with_SH)
set_data_collection (Vissim, lanes_with_SH)
set_vehicleinput    (Vissim, SimLen, DataInfo.TimeInterval, VehicleInput)


# 3. Run Simulation
print("Running simulation...")
Vissim.Simulation.RunSingleStep()

## Extract data per signal period
break_at = 0
set_signal(Vissim, Signal, break_at)
for break_at in BreakAt:
    Vissim.Simulation.SetAttValue('SimBreakAt', break_at)   # Set break_at
    Vissim.Simulation.RunContinuous()                       # Run simulation until 'break_at'
    set_signal(Vissim, Signal, break_at)                    # Set signal
Vissim.Simulation.RunContinuous()

## Extract data per hour
hour_step = math.ceil(break_at/3600)
for hour in range(hour_step):
    time_str = str(hour + 1)
    extract_from_datacollection_per_hour(Vissim, time_str, VehNum_hour, OccupRate_hour)
    extract_from_queue_per_hour         (Vissim, time_str, QStop_hour)
    extract_from_travtm_per_hour        (Vissim, time_str, AvgSpeed_hour)

## Close COM server:
print("Closing Vissim...")
Vissim = None


# 4. Calculate overall data
print("Calculating...")
SH_per_link = cal_SH_per_link(lanes_with_SH)

cal_occuprate_overall   (OccupRate_hour, OccupRate_overall, SimLen)
cal_qstop_overall       (QStop_hour, QStop_overall)
cal_qstop_per_meter     (QStop_hour, QStop_overall, lanes_with_SH)

linkseg_result = DataInfo.VissimInput[:-5] + '_Link Segment Results_001.att'
extract_from_linkseg(linkseg_result, lanes_with_SH, Density_overall, DelayRel_overall, AvgSpeed_overall)

prep_extract_from_node(hour_step, LOS_hour, EmissionCO_hour, EmissionVOC_hour)
if No_Node:     # If there was any node in Vissim network,
    node_result = DataInfo.VissimInput[:-5] + '_Node Results_001.att'
    extract_from_node(node_result, No_Node, EmissionCO, EmissionVOC, LOS_hour, EmissionCO_hour, EmissionVOC_hour)


# 5. Report
print("Reporting...")
excel = com.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False     # To merge cells
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

print_simul_info    (ws, DataInfo, SimLen)
print_explanation   (ws)
print_overall       (ws, lanes_with_SH, SH_per_link, No_Node, DelayRel_overall, Density_overall, AvgSpeed_overall, QStop_overall, OccupRate_overall, EmissionCO, EmissionVOC)
print_hour          (ws, lanes_with_SH, SH_per_link, Link_TT, No_Node, VehNum_hour, QStop_hour, OccupRate_hour, AvgSpeed_hour, LOS_hour, EmissionCO_hour, EmissionVOC_hour)

ws.Columns(2).AutoFit()
wb.SaveAs(os.getcwd() + '\\output_' + start_time + '.xlsx')
excel.Quit()
