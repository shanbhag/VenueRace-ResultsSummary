# -*- coding: utf-8 -*-
'''
DESCRIPTION


REVISION HISTORY
YYYY-MM-DD ; Created
'''
# Module-level imports
import datetime as dt
import glob
import matplotlib.pyplot as plt
import numpy as np
import os
import pandas as pd
import scipy.interpolate as ipl
import xlsxwriter as xlwt
import sys

# Add path to system paths
PathList = []
for p in PathList:
    sys.path.append(p)

# Inputs
# Paths
PathToProjects = r'S:\parag\Concept2\Venue Race\Data'
Project = '2017-01-10'
PathToWP = r''
SearchString = '*Stroke Data.txt'
SplitsFile = 'Splits.xlsx'
ResampledFile = 'ResampledData.xlsx'

# pandas and DataFrame options
separator = ','
DataDict = {'heat':  'Heat',
            'team':  'Roeier',
            'time':  'Tijd',
            'dist':  'Afstand',
            'pace':  '500m Tijd',
            'sr':    'Tempo (spm)',
            'hr':    'Heart Rate',
            'ivl':   'Interval',
            'split': 'Split',
            'lap':   'Lap'}
DFcolumns = [DataDict['ivl'], DataDict['team'], DataDict['time'],
             DataDict['dist'], DataDict['pace'], DataDict['sr'],
             DataDict['hr']]
DropCols = [DataDict['ivl'], DataDict['hr']]

# Output values
DesiredSplits = np.array([250, 500, 750, 1000, 3500, 6000, 8500, 11000], dtype=float) #np.arange(0.0, 6000.1, 500.0, dtype=float)
ResamplePeriod = {'time': 5.0, 'dist': 10.0}
TimeFactor = 24*60*60
DistFactor = 1.0

BadNameList = ['Leeg', 'PM3']

colwidth = 10
ExcelFormats = {'TF1':'m:ss',
                'TF2':'m:ss.0',
                'DF1':'0',
                'DF2':'0.0'}
DataTypeList = [DataDict['split'], DataDict['lap'],
                DataDict['pace'], DataDict['sr']]

# Plotting
PageWidth = 297 / 25.4
PageHeight = 210 / 25.4
PageSize = (PageWidth, PageHeight)
PlotDPI = 60
LineStyles = ['b-', 'r-', 'g-', 'k-', 'm-', 'c-']
NxTicks = 11
MinPace = 90 / TimeFactor
MaxPace = 150 / TimeFactor
MinSR = 10
MaxSR = 40
yTicks = np.arange(MinPace, MaxPace+0.0001, 10/TimeFactor)


# Local functions / classes
def WriteExcel(fp, outDF, DataDict, FullList, TR=True, unit='Time',
               outDataList=['Split'], cw=10,
               xlFmt={'TF1': 'm:ss', 'TF2': 'm:ss.0', 'DF1': '0', 'DF2': '0.0'}):
    '''
    DESCRIPTION

    INPUTS
    fp              absolute file path
    outDF           DataFrame with data to b written
    DataDict        dictionary of ataFrame column names
    FullList        list of [heat,team] combinations names
    TR              Time race (True) or distance race (False)
    unit            column name in dataframe that is based on the race type
    outDataList     list of data types to be output (one per sheet)
    cw              column width in Excel file
    xlFmt           cell formatting dictionary

    OUTPUT - none

    REVISION HISTORY
    '''
    # Open file for writing
    wbk = xlwt.Workbook(fp)

    # Set formats
    fmt_head = wbk.add_format()
    fmt_head.set_bold(True)
    fmt_head.set_align('center')
    fmt_head.set_align('vcenter')
    fmt_head.set_text_wrap(True)

    fmt_time1 = wbk.add_format()
    fmt_time1.set_num_format(xlFmt['TF1'])
    fmt_time1.set_align('center')
    fmt_time1.set_align('vcenter')

    fmt_time2 = wbk.add_format()
    fmt_time2.set_num_format(xlFmt['TF2'])
    fmt_time2.set_align('center')
    fmt_time2.set_align('vcenter')

    fmt_dist1 = wbk.add_format()
    fmt_dist1.set_num_format(xlFmt['DF1'])
    fmt_dist1.set_align('center')
    fmt_dist1.set_align('vcenter')

    fmt_dist2 = wbk.add_format()
    fmt_dist2.set_num_format(xlFmt['DF2'])
    fmt_dist2.set_align('center')
    fmt_dist2.set_align('vcenter')

    fmt = dict()

    # For each data type, open worksheet and output data to the file
    for DT in outDataList:
        wsh = wbk.add_worksheet(DT)

        # Set column width
        wsh.set_column('A:XFD', cw)

        # Write headers
        if TR:
            wsh.merge_range('A1:A2', DataDict['time'], fmt_head)
            fmt_base = fmt_time1
            fmt[DataDict['split']] = fmt[DataDict['lap']] = fmt_dist2
        else:
            wsh.merge_range('A1:A2', DataDict['dist'], fmt_head)
            fmt_base = fmt_dist1
            fmt[DataDict['split']] = fmt[DataDict['lap']] = fmt_time2
        fmt[DataDict['pace']] = fmt_time2
        fmt[DataDict['sr']] = fmt_dist2

        for i, val in enumerate(outDF[unit].unique()):
            wsh.write(i+2, 0, val, fmt_base)

        icol = 0
        for [Heat, Team] in FullList:
            mask_heat = (outDF[DataDict['heat']] == Heat)
            mask_team = (outDF[DataDict['team']] == Team)

            # increment column number
            icol += 1

            # Write header
            wsh.write(0, icol, Heat, fmt_head)
            wsh.write(1, icol, Team, fmt_head)

            # Write splits
            for i, val in enumerate(outDF[mask_heat & mask_team][DT].values):
                wsh.write(i+2, icol, val, fmt[DT])

    wbk.close()
    # End of funciton

# Main program
if __name__ == '__main__':
    # List of files to parse
    FileList = glob.glob(os.path.join(PathToProjects,
                                      Project,
                                      PathToWP,
                                      SearchString))

    # Read in each file using pandas and add to data  DataFrame
    data = pd.DataFrame()
    for FilePath in FileList:
        heat = os.path.split(FilePath)[-1][:os.path.split(FilePath)[-1].
                                           find(SearchString[1:])].strip()
        temp = pd.read_csv(FilePath, sep=separator)
        temp.columns = DFcolumns
        temp[DataDict['heat']] = heat
        data = data.append(temp)

    # Fill missing fields and re-index
    data = data.fillna(method='ffill')
    data.reset_index(drop=True, inplace=True)
    for col in DropCols:
        del(data[col])

    # Drop all rows that have an invalid team name
    for s in BadNameList:
        data = data[~data[DataDict['team']].str.contains(s)]

    # Make a list of heats and teams (temporary for teams)
    HeatList = data[DataDict['heat']].unique()
    TeamList = data[data[DataDict['heat']] == HeatList[0]][DataDict['team']].unique()

    # Determine if it is a distance race or time race based on results from
    # first two rowers
    LastVal1 = data[(data[DataDict['heat']] == HeatList[0]) &
                    (data[DataDict['team']] == TeamList[0])][-1:]
    LastVal2 = data[(data[DataDict['heat']] == HeatList[0]) &
                    (data[DataDict['team']] == TeamList[1])][-1:]
    if np.allclose(LastVal1[DataDict['time']], LastVal2[DataDict['time']]):
        TimeRace = True
    else:
        TimeRace = False

    SplitDF = pd.DataFrame()
    ResampDF = pd.DataFrame()
    FullList = []
    for Heat in HeatList:
        mask_heat = (data[DataDict['heat']] == Heat)
        # Generate list of teams in heat
        TeamList = data[mask_heat][DataDict['team']].unique()

        for Team in TeamList:
            FullList.append([Heat, Team])
            mask_team = (data[DataDict['team']] == Team)

            if TimeRace:
                # x=time, y=distance
                x = data[mask_heat & mask_team][DataDict['time']].values
                y = data[mask_heat & mask_team][DataDict['dist']].values
                ResampleT = ResamplePeriod['time']
                xFactor = TimeFactor
                yFactor = DistFactor
                BaseUnit = DataDict['time']
            else:
                # x=distance, y=time
                x = data[mask_heat & mask_team][DataDict['dist']].values
                y = data[mask_heat & mask_team][DataDict['time']].values
                ResampleT = ResamplePeriod['dist']
                xFactor = DistFactor
                yFactor = TimeFactor
                BaseUnit = DataDict['dist']
            sr = data[mask_heat & mask_team][DataDict['sr']].values

            # Remove x values that are idential. This is required for the
            # interpolation algorithm
            # Then interpolate to desired splits and resampling frequency
            mask_x = np.append(True, (x[1:] != x[:-1]))
            f = ipl.pchip(x[mask_x], y[mask_x])
            f_sr = ipl.pchip(x[mask_x], sr[mask_x])

            # Splits
            Splits = f(DesiredSplits)
            Laps = np.array(Splits)
            Laps[1:] = Splits[1:] - Splits[:-1]
            if TimeRace:
                PaceTimes = (np.insert(DesiredSplits, 0, 0.0)[1:] -
                             np.insert(DesiredSplits, 0, 0.0)[:-1]) * 500 / Laps
            else:
                PaceTimes = Laps * 500 / (np.insert(DesiredSplits, 0, 0.0)[1:] -
                                          np.insert(DesiredSplits, 0, 0.0)[:-1])

            srSplit = np.zeros_like(Splits)
            dx = x[1:] - x[:-1]
            for iSplit, Split in enumerate(DesiredSplits[1:]):
                srmask_x = ((x >= DesiredSplits[iSplit]) & (x <= Split))[1:]
                srSplit[iSplit+1] = np.sum(sr[1:][srmask_x] *
                                           dx[srmask_x]) / np.sum(dx[srmask_x])

            temp = pd.DataFrame({DataDict['heat']:  Heat,
                                 DataDict['team']:  Team,
                                 DataDict['split']: Splits / yFactor,
                                 DataDict['lap']:   Laps / yFactor,
                                 DataDict['pace']:  PaceTimes / TimeFactor,
                                 DataDict['sr']:    srSplit,
                                 BaseUnit:          DesiredSplits / xFactor})
            SplitDF = SplitDF.append(temp)

            # Resampled data
            ResampleSplits = np.arange(0, x[-1]+0.1*ResampleT, ResampleT)
            Splits = f(ResampleSplits)
            Laps = np.array(Splits)
            Laps[1:] = Splits[1:] - Splits[:-1]
            if TimeRace:
                PaceTimes = (np.insert(ResampleSplits, 0, 0.0)[1:] -
                             np.insert(ResampleSplits, 0, 0.0)[:-1]) * 500 / Laps
            else:
                PaceTimes = Laps * 500 / (np.insert(ResampleSplits, 0, 0.0)[1:] -
                                          np.insert(ResampleSplits, 0, 0.0)[:-1])
            sr = f_sr(ResampleSplits)

            temp = pd.DataFrame({DataDict['heat']:  Heat,
                                 DataDict['team']:  Team,
                                 DataDict['split']: Splits / yFactor,
                                 DataDict['lap']:   Laps / yFactor,
                                 DataDict['pace']:  PaceTimes / TimeFactor,
                                 DataDict['sr']:    sr,
                                 BaseUnit:          ResampleSplits / xFactor})
            ResampDF = ResampDF.append(temp)

    # Replace infinity and nan values with 0
    # These occur at the start fo some record because of divide by zero
    # or zero / zero
    SplitDF.replace([np.inf, -np.inf], np.nan, inplace=True)
    SplitDF.fillna(0, inplace=True)
    ResampDF.replace([np.inf, -np.inf], np.nan, inplace=True)
    ResampDF.fillna(0, inplace=True)

    # Write splits to Excel
    WriteExcel(os.path.join(PathToProjects, Project, PathToWP, SplitsFile),
               SplitDF, DataDict, FullList, TR=TimeRace, unit=BaseUnit,
               outDataList=DataTypeList, cw=colwidth, xlFmt=ExcelFormats)

    WriteExcel(os.path.join(PathToProjects, Project, PathToWP, ResampledFile),
               ResampDF, DataDict, FullList, TR=TimeRace, unit=BaseUnit,
               outDataList=DataTypeList, cw=colwidth, xlFmt=ExcelFormats)

    # Generate plots for each rower
    for iPlot, [Heat, Team] in enumerate(FullList):
        mask_heat = (ResampDF[DataDict['heat']] == Heat)
        mask_team = (ResampDF[DataDict['team']] == Team)

        fig = plt.figure(iPlot+1, PageSize, dpi=PlotDPI)
        ax0 = fig.add_subplot(1, 1, 1)
        ax0.plot(ResampDF[mask_heat & mask_team][BaseUnit].values,
                 ResampDF[mask_heat & mask_team][DataDict['pace']].values,
                 LineStyles[0])

        # Set plot bounds, etc
        xmax = ResampDF[mask_heat & mask_team][BaseUnit].values[-1]
        ax0.set_xbound(0, xmax)
        ax0.set_xlabel(BaseUnit)
        xTicks = np.linspace(0, xmax, NxTicks+1)
        if TimeRace:
            xTickLabels = [dt.datetime.strftime(dt.datetime(2000, 1, 1) +
                                                dt.timedelta(xTick), '%M:%S')
                           for xTick in xTicks]
        else:
            xTickLabels = [format(xTick, '.0f') for xTick in xTicks]
        ax0.set_xticklabels(xTickLabels)
        ax0.set_ylabel(DataDict['pace'])
        yTickLabels = [dt.datetime.strftime(dt.datetime(2000, 1, 1) +
                                            dt.timedelta(yTick), '%M:%S')
                       for yTick in yTicks]
        ax0.set_ybound(MinPace, MaxPace)
        plt.xticks(xTicks, xTickLabels)
        plt.yticks(yTicks, yTickLabels)
        plt.grid(True)

        # Add stroke rat to plot, on scondary y axis
        ax1 = ax0.twinx()
        ax1.plot([0, 0], [0, 0], LineStyles[0], label=DataDict['pace'])
        ax1.plot(ResampDF[mask_heat & mask_team][BaseUnit].values,
                 ResampDF[mask_heat & mask_team][DataDict['sr']].values,
                 LineStyles[1],
                 label=DataDict['sr'])

        ax1.set_ylabel(DataDict['sr'])
        ax1.set_ybound(MinSR, MaxSR)
        plt.legend(loc='lower center')

        plt.title(Heat + ' - ' + Team, fontsize=14)

        fig.savefig(os.path.join(PathToProjects,
                                 Project,
                                 PathToWP,
                                 Heat + ' - ' + Team + '.png'),
                    dpi=PlotDPI,
                    format='png')

        plt.close(iPlot+1)
