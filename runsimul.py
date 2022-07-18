# ==========================================================================
# Author : HyeAnn Lee
# ==========================================================================
from warning import *


def _get_AttValue_num(item, attribute):
    # Input
    # > 'item'      : CDispatch. Vissim element.
    # > 'attribute' : str.
    #
    # Output
    # > 'value'     : Non-negative number.

    value = item.AttValue(attribute)
    if value is None:
        return 0

    if value < 0:
        warning(" ERROR : Negative AttValue has been detected...")

    return value


def set_signal(Vissim, list_of_sigcon, break_at):
    # Input
    # > 'list_of_sigcon'    : 1D-list of SigControl().
    # > 'break_at'          : int.
    #
    # Set signals of particular time 'break_at'.
    # Do nothing if signals in a signal controller does not have to be changed.

    for sigcon in list_of_sigcon:   # For each signal controller,
        # Check if signal of 'sigcon' should be changed or not.
        try:
            index = sigcon.BreakAt.index(break_at)
        except ValueError:
            continue

        # Find appropriate Signal Controller.
        SC_Iter = Vissim.Net.SignalControllers.Iterator
        while SC_Iter.Valid:
            if SC_Iter.Item.AttValue('Name') == sigcon.Name:
                break
            SC_Iter.Next()

        # Find appropriate signal indicator set.
        signal_Iter = iter(sigcon.SigInd[index % len(sigcon.SigInd)])

        # Set signals.
        SG_Iter = SC_Iter.Item.SGs.Iterator
        while SG_Iter.Valid:
            signal = next(signal_Iter)
            SG_Iter.Item.SetAttValue('SigState', signal)
            SG_Iter.Next()

    return


def extract_from_datacollection_per_hour(Vissim, time_str, VehNum_hour, OccupRate_hour):
    # Input
    # > 'time_str'          : str.
    # > 'VehNum_hour'       : list.
    # > 'OccupRate_hour'    : list.

    # Add another list, containing information of the hour.
    VehNum_hour.append([])
    OccupRate_hour.append([])

    subattr = '(Current,' + time_str + ',All)'
    DC_Iter = Vissim.Net.DataCollectionMeasurements.Iterator
    while DC_Iter.Valid:
        DC = DC_Iter.Item
        VehNum_hour[-1].    append(_get_AttValue_num(DC, 'Vehs' + subattr)             )
        OccupRate_hour[-1]. append(_get_AttValue_num(DC, 'OccupRate' + subattr) * 100  )
        DC_Iter.Next()

    # Now, last element of 'VehNum_hour' and 'OccupRate_hour' is list of non-negative numbers.

    return


def extract_from_queue_per_hour(Vissim, time_str, QStop_hour):
    # Input
    # > 'time_str'      : str.
    # > 'QStop_hour'    : list.

    # Add another list, containing information of the hour.
    QStop_hour.append([])

    QC_Iter = Vissim.Net.QueueCounters.Iterator
    while QC_Iter.Valid:
        QC = QC_Iter.Item
        QStop_hour[-1].append(_get_AttValue_num(QC, 'QStops(Current,' + time_str + ')'))
        QC_Iter.Next()

    # Now, last elements of 'QStop_hour' is a list of non-negative numbers.

    return


def extract_from_travtm_per_hour(Vissim, time_str, AvgSpeed_hour):
    # Input
    # > 'time_str'      : str.
    # > 'AvgSpeed_hour' : list.

    def _change_unit(meter_per_sec):
        km_per_hour = meter_per_sec * 18 / 5
        return km_per_hour

    # Add another list, containing information of the hour.
    AvgSpeed_hour.append([])

    TT_Iter = Vissim.Net.VehicleTravelTimeMeasurements.Iterator
    while TT_Iter.Valid:
        TT = TT_Iter.Item
        Dist    = _get_AttValue_num(TT, 'Dist')
        TravTm  = _get_AttValue_num(TT, 'TravTm(Current,' + time_str + ',All)')
        if TravTm == 0:
            AvgSpeed_hour[-1].append(-1)
        else:
            AvgSpeed_hour[-1].append(_change_unit(Dist / TravTm))

        TT_Iter.Next()

    # Now, last elements of 'AvgSpeed_hour' is a list of numbers.
    # -1 value of the last element means that there was no vehicles passing through the TT.

    return
