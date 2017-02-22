#[dyntools_demo.py]  Demo for using functions from dyntools module
# ====================================================================================================
'''
'dyntools' module provide access to data in PSS(R)E Dynamic Simulation Channel Output file.
This module has functions:
- to get channel data in Python scripts for further processing
- to get channel information and their min/max range
- to export data to text files, excel spreadsheets
- to open multiple channel output files and post process their data using Python scripts
- to plot selected channels
- to plot and insert plots in word document

This is an example file showing how to use various functions available in dyntools module.

Other Python modules 'matplotlib', 'numpy' and 'python win32 extension' are required to be
able to use 'dyntools' module.
Self installation EXE files for these modules are available at:
   PSSE User Support Web Page and follow link 'Python Modules used by PSSE Python Utilities'.

- The dyntools is developed and tested using these versions of with matplotlib and numpy.
  When using Python 2.5
  Python 2.5 matplotlib-1.1.1
  Python 2.5 numpy-1.7.0
  Python 2.5 pywin32-218

  When using Python 2.7
  Python 2.7 matplotlib-1.2.0
  Python 2.7 numpy-1.7.0
  Python 2.7 pywin32-218

  Versions later than these may work.

---------------------------------------------------------------------------------
How to use this file?
- Open Python IDLE (or any Python Interpreter shell)
- Open this file
- run (F5)

Note: Do NOT run this file from PSS(R)E GUI. The 'xyplots' function from dyntools can
save plots to eps, png, pdf or ps files. However, creating only 'eps' files from inside
PSS(R)E GUI works. This is because different backends matplotlib uses to create different
plot types.
When run from any Python interpreter (outside PSS(R)E GUI) plots can be saved to any of
these four (eps, png, pdf or ps) file types.

Get information on functions available in dyntools as:
import dyntools
help(dyntools)

'''

import os, sys, collections
from subsystem_info import subsystem_info
# =====================================================================================================

def check_psse_example_folder():
    # if called from PSSE's Example Folder, create report in subfolder 'Output_Pyscript'

    outdir = os.getcwd()
    cwd = outdir.lower()
    i = cwd.find('pti')
    j = cwd.find('psse')
    k = cwd.find('example')
    if i>0 and j>i and k>j:     # called from Example folder
        outdir = os.path.join(outdir, 'Output_Pyscript')
        if not os.path.exists(outdir): os.mkdir(outdir)

    return outdir

# =============================================================================================

def get_demotest_file_names(outpath):

    if outpath:
        outdir = outpath
    else:
        outdir = check_psse_example_folder()

    outfile1 = os.path.join(outdir,'out_dyntools_demo_bus154_fault.out')
    outfile2 = os.path.join(outdir,'out_dyntools_demo_bus3018_gentrip.out')
    outfile3 = os.path.join(outdir,'out_dyntools_demo_brn3005_3007_trip.out')
    prgfile  = os.path.join(outdir,'out_dyntools_demo_progress.txt')

    return outfile1, outfile2, outfile3, prgfile

# =============================================================================================
# Run Dynamic simulation on SAVNW to generate .out files

def run_savnw_simulation(datapath, outfile1, outfile2, outfile3, prgfile):
    _F1_start=100
    _F1_end=100+70*1/60.0;

    _F2_start = 300
    _F2_end = 400;

    _runto=300;
    import psspy
    psspy.psseinit()

    savfile = 'IEEE 9 Bus_modifiedj4ab.sav'
    snpfile = 'IEEE 9 Bus_modifiedj4ab.snp'
    _i = psspy.getdefaultint()
    _f = psspy.getdefaultreal()
    _s = psspy.getdefaultchar()
    INTGAR = [_i] * 7
    REALAR = [_f] * 8

    if datapath:
        savfile = os.path.join(datapath, savfile)
        snpfile = os.path.join(datapath, snpfile)

    psspy.lines_per_page_one_device(1,90)
    psspy.progress_output(2,prgfile,[0,0])  #Use this API to specify the progress output device.

    ierr = psspy.case(savfile)   #Use this API to open a PSSE Saved Case file and transfers its data into the PSSE working case
    if ierr:
        psspy.progress_output(1,"",[0,0])
        print(" psspy.case Error")
        return
    ierr = psspy.rstr(snpfile)#Use this API to read a dynamics Snapshot File into PSSE working memory (activity RSTR).
    if ierr:
        psspy.progress_output(1,"",[0,0])
        print(" psspy.rstr Error")
        return

    psspy.strt(0,outfile1)  #strt(option, outfile)   #Use this API to initialize a PSSE dynamic simulation for state-space simulations (i.e., in preparation for activity RUN) and to specify the Channel Output File into which the output channel values are to be recorded during the dynamic simulation (activity STRT).
    psspy.run(0, _F1_start,5000,1,0) #Use this API to calculate PSSE state-space dynamic simulations (activity RUN).
    #psspy.dist_bus_fault(8,1, 230.0,[0.0,-0.2E+10]) #Use this API routine to apply a fault at a bus during dynamic simulations. (Note: use DIST_BUS_FAULT_2 if phase voltages are to be calculated during the simulation.)


    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo
    psspy.dist_branch_fault(8,9, '1',3,0.0,[0.0,0.000001])
    ##psspy.dist_branch_trip(8, 7, '1')
    ##psspy.load_chng_5(11, r"""1""", [0, _i, _i, _i, _i, _i, _i], [_f, _f, _f, _f, _f, _f, _f, _f])
    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo

    psspy.run(0, _F1_end+1.5*1/60,5000,1,0)
    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo



    #psspy.dist_branch_close(8,7,'1')
   ## psspy.load_chng_5(11, r"""1""", [1, _i, _i, _i, _i, _i, _i], [_f, _f, _f, _f, _f, _f, _f, _f])
    psspy.dist_clear_fault(1)
    psspy.run(0, _F1_end+2.0, 5000, 1, 0)
    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo

    #psspy.dist_clear_fault(1)  #Use this API to clear a fault during dynamic simulations. The fault must have previously been applied using one of the following APIs:
    psspy.run(0, _runto,5000,1,0)
#trigger machine
#    psspy.case(savfile)    #Use this API to open a PSSE Saved Case file and transfers its data into the PSSE working case
#    psspy.rstr(snpfile)     #Use this API to read a dynamics Snapshot File into PSSE working memory (activity RSTR).
#    psspy.strt(0,outfile2)
#    psspy.run(0, 1.0,1000,1,0)
#    psspy.dist_machine_trip(2,'1')
#    psspy.run(0, 10.0,1000,1,0)
#trigger line
#    psspy.case(savfile)
#    psspy.rstr(snpfile)
#    psspy.strt(0,outfile3)
#    psspy.run(0, 1.0,1000,1,0)
#    psspy.dist_branch_trip(7,8,'1')
#    psspy.run(0, 10.0,1000,1,0)

    psspy.lines_per_page_one_device(2,10000000)
    psspy.progress_output(1,"",[0,0])




# =============================================================================================
# 0. Run savnw dynamics simulation to create .out files

def test0_run_simulation(datapath=None, outpath=None):

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    run_savnw_simulation(datapath, outfile1, outfile2, outfile3, prgfile)

    print(" Done SAVNW dynamics simulation")

# =============================================================================================
# 1. Data extraction/information
Matrix_VOLT =[];
Matrix_FREQ =[];
Matrix_ANGL =[];
Matrix_POWR =[];
Matrix_VARS =[];
ANGL_hdr=[];
FREQ_hdr=[];
VOLT_hdr=[];
POWR_hdr=[];
VARS_hdr=[];
def test1_data_extraction(outpath=None, show=True):
    import numpy as np
    import dyntools
    global Matrix_VOLT
    global Matrix_FREQ
    global Matrix_ANGL
    global Matrix_POWR
    global Matrix_VARS

    global ANGL_hdr
    global FREQ_hdr
    global VOLT_hdr
    global POWR_hdr
    global VARS_hdr
    print outpath
    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    # create object
    chnfobj = dyntools.CHNF(outfile1)

    print '\n Testing call to get_data'
    sh_ttl, ch_id, ch_data = chnfobj.get_data()
    print sh_ttl
    print ch_id

    ch_data.get('time')

    ANGL_ctr=0; ANGL_pos=[];ANGL_hdr=[];
    FREQ_ctr = 0;FREQ_pos=[];FREQ_hdr=[];
    VOLT_ctr = 0;VOLT_pos=[];VOLT_hdr=[];
    POWR_ctr = 0;POWR_pos=[];POWR_hdr=[];
    VARS_ctr = 0;VARS_pos=[];VARS_hdr=[];
    nchannels=len(ch_data)
    ANGL_hdr.append("Time")
    FREQ_hdr.append("Time")
    VOLT_hdr.append("Time")
    POWR_hdr.append("Time")
    VARS_hdr.append("Time")
    #iterate the channels to find angle channels
    for x in range(1, nchannels):
        tclass=ch_id.get(x);
        if (tclass.find('ANGL')>=0):
            ANGL_ctr+=1;
            ANGL_pos.append(x)
            ANGL_hdr.append(tclass)


        if (tclass.find('FREQ')>=0):
            FREQ_ctr+=1;
            FREQ_pos.append(x)
            FREQ_hdr.append(tclass)

        if (tclass.find('VOLT') >= 0):
            VOLT_ctr += 1;
            VOLT_pos.append(x)
            VOLT_hdr.append(tclass)

        if (tclass.find('POWR') >= 0):
            POWR_ctr += 1;
            POWR_pos.append(x)
            POWR_hdr.append(tclass)

        if (tclass.find('VARS') >= 0):
            VARS_ctr += 1;
            VARS_pos.append(x)
            VARS_hdr.append(tclass)

    xp1 = len(ch_data.get('time'))
    Matrix_ANGL = [[0 for x in range(ANGL_ctr + 1)] for y in range(xp1)]
    Matrix_FREQ = [[0 for x in range(FREQ_ctr + 1)] for y in range(xp1)]
    Matrix_VOLT = [[0 for x in range(VOLT_ctr + 1)] for y in range(xp1)]
    Matrix_POWR = [[0 for x in range(POWR_ctr + 1)] for y in range(xp1)]
    Matrix_VARS = [[0 for x in range(VARS_ctr + 1)] for y in range(xp1)]



    for xi in range(0, xp1):
        Matrix_ANGL[xi][0]= ch_data.get('time')[xi]
        for xj in range(0, ANGL_ctr):
            Matrix_ANGL[xi][1+xj] = ch_data.get(ANGL_pos[xj])[xi]-ch_data.get(ANGL_pos[0])[xi]

    for xi in range(0, xp1):
        Matrix_FREQ[xi][0]= ch_data.get('time')[xi]
        for xj in range(0, FREQ_ctr):
            Matrix_FREQ[xi][1+xj] = ch_data.get(FREQ_pos[xj])[xi]+60

    for xi in range(0, xp1):
        Matrix_VOLT[xi][0]= ch_data.get('time')[xi]
        for xj in range(0, VOLT_ctr):
            Matrix_VOLT[xi][1+xj] = ch_data.get(VOLT_pos[xj])[xi]

    for xi in range(0, xp1):
        Matrix_POWR[xi][0] = ch_data.get('time')[xi]
        for xj in range(0, POWR_ctr):
            Matrix_POWR[xi][1 + xj] = ch_data.get(POWR_pos[xj])[xi]

    for xi in range(0, xp1):
        Matrix_VARS[xi][0] = ch_data.get('time')[xi]
        for xj in range(0, VARS_ctr):
            Matrix_VARS[xi][1 + xj] = ch_data.get(VARS_pos[xj])[xi]

    np.savetxt(outpath + "\\" + "ANG.csv", Matrix_ANGL,fmt="%12.6f",delimiter="; ",header='; '.join(ANGL_hdr))
    np.savetxt(outpath + "\\" + "FREQ.csv", Matrix_FREQ,fmt="%12.6f",delimiter="; ",header='; '.join(FREQ_hdr))
    np.savetxt(outpath+"\\"+"VOLT.csv",Matrix_VOLT,fmt="%12.6f",delimiter="; ",header='; '.join(VOLT_hdr))
    np.savetxt(outpath + "\\" + "P.csv", Matrix_POWR, fmt="%12.6f", delimiter="; ", header='; '.join(POWR_hdr))
    np.savetxt(outpath + "\\" + "Q.csv", Matrix_VARS, fmt="%12.6f", delimiter="; ", header='; '.join(VARS_hdr))


    #np.savetxt("test.txt",np.asarray( [ch_data.get('time')], [ch_data.get(1)],[ch_data.get(2)] ) );

    print '\n Testing call to get_id'
    sh_ttl, ch_id = chnfobj.get_id()
    print sh_ttl
    print ch_id

    print '\n Testing call to get_range'
    ch_range = chnfobj.get_range()
    print ch_range

    print '\n Testing call to get_scale'
    ch_scale = chnfobj.get_scale()
    print ch_scale

    print '\n Testing call to print_scale'
    chnfobj.print_scale()

    print '\n Testing call to txtout'
    chnfobj.txtout(channels=[1,2,21,22])

    #print '\n Testing call to xlsout'
#    chnfobj.xlsout(channels=[2,3,4,7,8,10],show=show)

# =============================================================================================
# 2. Multiple subplots in a figure, but one trace in each subplot
#    Channels specified with normal dictionary

# See how "set_plot_legend_options" method can be used to place and format legends
def format_label(input_label):
    ki=input_label.find("[")+1
    kf = input_label.find("]")
    text1=input_label[ki:kf]
    return text1

def plot_qt(plottype,title="",filename="out.png",list_buses=[], t_start=-1,t_len=10000):
    import numpy as np
    import math
    import matplotlib.pyplot as plt
    #plottype=upper(plottype)
    if (plottype=="V"):
        a_TEMP = np.array(Matrix_VOLT)
        h_TEMP=VOLT_hdr;
        ylabel="Voltage (PU)"

    if (plottype=="F"):
        a_TEMP = np.array(Matrix_FREQ)
        h_TEMP = FREQ_hdr;
        ylabel = "Frequency (Hz)"

    if (plottype=="A"):
        a_TEMP = np.array(Matrix_ANGL)
        h_TEMP = ANGL_hdr;
        ylabel = "Angle (DEG)"

    if (plottype=="P"):
        a_TEMP = np.array(Matrix_POWR)
        h_TEMP = POWR_hdr;
        ylabel = "Power (MW)"

    if (plottype=="Q"):
        a_TEMP = np.array(Matrix_VARS)
        h_TEMP = VARS_hdr;
        ylabel = "R Power (VARS)"

    fig, ax = plt.subplots(nrows=1, ncols=1)  # create figure & 1 axis

    timect = a_TEMP[1,0] - a_TEMP[0,0]

    startpos =math.floor( t_start / timect);


    if startpos<0:
        startpos=0;

    endpos = math.floor( t_len / timect)+startpos;


    if endpos>len(a_TEMP[:, 1]):
        endpos=len(a_TEMP[:, 1]);

    startpos = int(startpos)
    endpos = int(endpos)

    if (len(list_buses)==0):
        for xi in range(1, len(a_TEMP[1, :])):
            ax.plot(a_TEMP[startpos:endpos, 0], a_TEMP[startpos:endpos, xi],  linewidth=.75,label=str(xi)+": "+format_label(h_TEMP[xi]))
    else:
        for xi in range(0, len(list_buses)):
            xj=list_buses[xi];
            ax.plot(a_TEMP[startpos:endpos, 0], a_TEMP[startpos:endpos, xj],  linewidth=.75,label=format_label(h_TEMP[xj]))

    plt.grid( linestyle='--', linewidth=.5)
    plt.legend(prop={'size':6})
    plt.title(title)
    plt.ylabel(ylabel)
    plt.xlabel('Time (s)')
    fig.savefig(outpath + '/'+filename)  # save the figure to file
    plt.close(fig)  # close the figure



def test2_subplots_one_trace(outpath=None, show=True):
    '''
    plot_qt("V","","V_all.png")
    plot_qt("V","Bus voltages","V_filtered.png",[3,4,5,6,7,8,9,10])

    plot_qt("F", "","F_all.png")
    plot_qt("F","Bus frequencies", "F_filtered.png", [1,2])

    plot_qt("A","", "A_all.png")
    plot_qt("A","Bus angles relative to BUS 2", "A_filtered.png", [1,2])

    plot_qt("P","", "P_all.png")
    plot_qt("P","P at generation buses", "P_Gen.png", [3,4])
    plot_qt("P","P at load buses", "P_Load.png", [10, 16])
    plot_qt("P", "P flows between buses", "P_Buses.png", [ 6, 7, 13])

    plot_qt("Q","","Q_all.png")
    plot_qt("Q","Q at generation buses", "Q_Gen.png", [3, 4])
    plot_qt("Q","Q at load buses","Q_Load.png", [ 10, 16])
    plot_qt("Q","Q flows between buses", "Q_Buses.png", [ 6, 7, 13])
'''
    tstart=99.99
    tlen = 5.0

    plot_qt("V", "", "V_all.png",[],tstart,tlen)
    plot_qt("V", "Bus voltages", "V_filtered.png", [3, 4, 5, 6, 7, 8, 9, 10],tstart,tlen)

    plot_qt("F", "", "F_all.png",[],tstart,tlen)
    plot_qt("F", "Bus frequencies", "F_filtered.png", [1, 2],tstart,tlen)

    plot_qt("A", "", "A_all.png",[],tstart,tlen)
    plot_qt("A", "Bus angles relative to BUS 2", "A_filtered.png", [1, 2],tstart,tlen)

    plot_qt("P", "", "P_all.png",[],tstart,tlen)
    plot_qt("P", "P at generation buses", "P_Gen.png", [3, 4],tstart,tlen)
    plot_qt("P", "P at load buses", "P_Load.png", [10, 16],tstart,tlen)
    plot_qt("P", "P flows between buses", "P_Buses.png", [6, 7, 13],tstart,tlen)

    plot_qt("Q", "", "Q_all.png",[],tstart,tlen)
    plot_qt("Q", "Q at generation buses", "Q_Gen.png", [3, 4],tstart,tlen)
    plot_qt("Q", "Q at load buses", "Q_Load.png", [10, 16],tstart,tlen)
    plot_qt("Q", "Q flows between buses", "Q_Buses.png", [6, 7, 13],tstart,tlen)

    print 'Write Data//'

# =============================================================================================
# 3. Multiple subplots in a figure and more than one trace in each subplot
#    Channels specified with normal dictionary

def test3_subplots_mult_trace(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':2,'columns':2,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':True}

    optnchn1 = {1:{'chns':[1]},2:{'chns':[2]},3:{'chns':[3]},4:{'chns':[4]},5:{'chns':[5]}}
    pn,x     = os.path.splitext(outfile1)
    pltfile1 = pn+'.png'

   # optnchn2 = {1:{'chns':{outfile2:1}},
   #             2:{'chns':{'v82_test1_bus_fault.out':3}},
   #             3:{'chns':4},
   #             4:{'chns':[5]}
   #            }
   # pn,x     = os.path.splitext(outfile2)
   # pltfile2 = pn+'.pdf'

   # optnchn3 = {1:{'chns':{'v80_test1_bus_fault.out':1}},
   #             2:{'chns':{'v80_test2_complex_wind.out':[1,5]}},
   #             3:{'chns':{'v82_test1_bus_fault.out':3}},
   #             5:{'chns':[4,5]},
   #            }
   # pn,x     = os.path.splitext(outfile3)
   # pltfile3 = pn+'.eps'

    figfiles1 = chnfobj.xyplots(optnchn1,optnfmt,pltfile1)
   # figfiles2 = chnfobj.xyplots(optnchn2,optnfmt,pltfile2)
   # figfiles3 = chnfobj.xyplots(optnchn3,optnfmt,pltfile3)

    figfiles = figfiles1[:]
   # figfiles.extend(figfiles2)
   # figfiles.extend(figfiles3)
    if figfiles:
        print 'Plot fils saved:'
        for f in figfiles:
            print '    ', f

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

# =============================================================================================
# 4. Multiple subplots in a figure, but one trace in each subplot
#    Channels specified with Ordered dictionary

def test4_subplots_mult_trace_OrderedDict(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':1,'columns':1,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':True}

    optnchn  = {1:{'chns':collections.OrderedDict([(outfile1,3), (outfile2,3), (outfile3,3)]),
                  }
               }
    p,nx     = os.path.split(outfile1)
    pltfile  = os.path.join(p, 'plot_chns_ordereddict.png')

    figfiles = chnfobj.xyplots(optnchn,optnfmt,pltfile)

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

# =============================================================================================
# 5. Do XY plots and insert them into word file
# Does not work because win32 API to Word does not work.

def test5_plots2word(outpath=None, show=True):

    import dyntools

    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2, outfile3)

    p,nx       = os.path.split(outfile1)
    docfile    = os.path.join(p,'savnw_response')
    overwrite  = True
    caption    = True
    align      = 'center'
    captionpos = 'below'
    height     = 0.0
    width      = 0.0
    rotate     = 0.0

    optnfmt  = {'rows':3,'columns':1,'dpi':300,'showttl':True}

    optnchn  = {1:{'chns':{outfile1:1,  outfile2:1,  outfile3:1} },
                2:{'chns':{outfile1:7,  outfile2:7,  outfile3:7} },
                3:{'chns':{outfile1:17, outfile2:17, outfile3:17} },
                4:{'chns':[1,2,3,4,5]},
                5:{'chns':{outfile2:[1,2,3,4,5]} },
                6:{'chns':{outfile3:[1,2,3,4,5]} },
               }
    ierr, docfile = chnfobj.xyplots2doc(optnchn,optnfmt,docfile,show,overwrite,caption,align,
                        captionpos,height,width,rotate)

    if not ierr:
        print 'Plots saved to file:'
        print '    ', docfile

# =============================================================================================
# Run all tests and save plot and report files.

def run_all_tests(datapath=None, outpath=None):

    show = False        # This must be false in this test.

    test0_run_simulation(datapath, outpath)

    test1_data_extraction(outpath, show)

    test2_subplots_one_trace(outpath, show)

    #test3_subplots_mult_trace(outpath, show)

    #test4_subplots_mult_trace_OrderedDict(outpath, show)

   # test5_plots2word(outpath, show)

# =============================================================================================

if __name__ == '__main__':

    import psspy

    #(a) Run one test a time
    # Need to run "test0_run_simulation(..)" before running other tests.
    # After running "test0_run_simulation(..)", run other tests one at a time.
    datapath = 'C:\Users\jonathans\PycharmProjects\PSSE_DYN\SampleSystem\IN'
    outpath  = 'C:\Users\jonathans\PycharmProjects\PSSE_DYN\SampleSystem\OUT'
    show     = True     # True  --> create, save and show Excel spreadsheets and Plots when done
                        # False --> create, save but do not show Excel spreadsheets and Plots when done

    #test0_run_simulation(datapath, outpath)

    #test1_data_extraction(outpath, show)

    #test2_subplots_one_trace(outpath, show)

    #test3_subplots_mult_trace(outpath, show)

    #test4_subplots_mult_trace_OrderedDict(outpath, show)

    #test5_plots2word(outpath, show)

    #(b) Run all tests

    run_all_tests(datapath, outpath)

# =============================================================================================
