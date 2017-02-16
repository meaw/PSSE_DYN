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
    _start=50
    _end=100;
    _runto=200;
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
    psspy.run(0, _start,5000,1,0) #Use this API to calculate PSSE state-space dynamic simulations (activity RUN).
    #psspy.dist_bus_fault(8,1, 230.0,[0.0,-0.2E+10]) #Use this API routine to apply a fault at a bus during dynamic simulations. (Note: use DIST_BUS_FAULT_2 if phase voltages are to be calculated during the simulation.)


    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo

    psspy.load_chng_5(12, r"""1""", [0, _i, _i, _i, _i, _i, _i], [_f, _f, _f, _f, _f, _f, _f, _f])
    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo

    psspy.run(0, _end+1.5*1/60,5000,1,0)
    businfo = subsystem_info('bus', ['NUMBER', 'NAME', 'PU'], sid=-1)
    print businfo



    #psspy.dist_branch_close(8,11,'1')
    psspy.load_chng_5(12, r"""1""", [1, _i, _i, _i, _i, _i, _i], [_f, _f, _f, _f, _f, _f, _f, _f])
    psspy.run(0, _end+2.0, 5000, 1, 0)
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
ANGL_hdr=[];
FREQ_hdr=[];
VOLT_hdr=[];

def test1_data_extraction(outpath=None, show=True):
    import numpy as np
    import dyntools
    global Matrix_VOLT
    global Matrix_FREQ
    global Matrix_ANGL

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
    nchannels=len(ch_data)
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

    xp1 = len(ch_data.get('time'))
    Matrix_ANGL = [[0 for x in range(ANGL_ctr + 1)] for y in range(xp1)]
    Matrix_FREQ = [[0 for x in range(FREQ_ctr + 1)] for y in range(xp1)]
    Matrix_VOLT = [[0 for x in range(VOLT_ctr + 1)] for y in range(xp1)]


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

    np.savetxt(outpath + "\\" + "ANG.csv", Matrix_ANGL,fmt="%12.6f",delimiter="; ",header="time;"+'; '.join(ANGL_hdr))
    np.savetxt(outpath + "\\" + "FREQ.csv", Matrix_FREQ,fmt="%12.6f",delimiter="; ",header="time;"+'; '.join(FREQ_hdr))
    np.savetxt(outpath+"\\"+"VOLT.csv",Matrix_VOLT,fmt="%12.6f",delimiter="; ",header="time;"+'; '.join(VOLT_hdr))

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

def test2_subplots_one_trace(outpath=None, show=True):

    import dyntools
    import numpy as np
    global Matrix_VOLT
    global Matrix_FREQ
    global Matrix_ANGL
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(nrows=1, ncols=1)  # create figure & 1 axis

    a_VOLT = np.array(Matrix_VOLT)
    a_ANGL = np.array(Matrix_ANGL)
    a_FREQ = np.array(Matrix_FREQ)
    for xi in range(1, len(a_VOLT[1,:])):
        ax.plot(a_VOLT[:,0], a_VOLT[:,xi],label=VOLT_hdr[xi])

    plt.legend()
    plt.ylabel('Voltage (PU)')
    plt.xlabel('Time (s)')

    fig.savefig(outpath+'/to.png')  # save the figure to file
    plt.close(fig)  # close the figure


    print 'Write Data//'
    outfile1, outfile2, outfile3, prgfile = get_demotest_file_names(outpath)

    chnfobj = dyntools.CHNF(outfile1, outfile2)

    chnfobj.set_plot_page_options(size='letter', orientation='portrait')
    chnfobj.set_plot_markers('square', 'triangle_up', 'thin_diamond', 'plus', 'x',
                             'circle', 'star', 'hexagon1')
    chnfobj.set_plot_line_styles('solid', 'dashed', 'dashdot', 'dotted')
    chnfobj.set_plot_line_colors('blue', 'red', 'black', 'green', 'cyan', 'magenta', 'pink', 'purple')

    optnfmt  = {'rows':4,'columns':2,'dpi':300,'showttl':True, 'showoutfnam':True, 'showlogo':False,
                'legendtype':1, 'addmarker':True}

    optnchn1 = {1:{'chns':1,    'title':'ANGL 2'},
                2:{'chns':2,    'title':'ANGL 3'},
                3:{'chns':3,    'title':'POWR 2'},
                4:{'chns':4,    'title':'POWR 2'},
                5:{'chns':15,   'title':'FREQ 2'},
                6:{'chns':16,   'title':'FREQ 3'},
                7: {'chns': Matrix_VOLT[:][1], 'title': 'VOLT 2'},
                8: {'chns': 22, 'title': 'VOLT 3'},
                }
    pn,x     = os.path.splitext(outfile1)
    pltfile1 = pn+'.png'

    optnchn2 = {1:{'chns':{outfile2:1}, 'title':'Channel 1 from bus3018_gentrip'},
                2:{'chns':{outfile2:6}, 'title':'Channel 6 from bus3018_gentrip'},
                3:{'chns':{outfile2:11}},
                4:{'chns':{outfile2:16}},
                5:{'chns':{outfile2:26}},
                6:{'chns':{outfile2:40}},
                }
    pn,x     = os.path.splitext(outfile2)
    pltfile2 = pn+'.png'

    figfiles1 = chnfobj.xyplots(optnchn1,optnfmt,pltfile1)

    chnfobj.set_plot_legend_options(loc='lower center', borderpad=0.2, labelspacing=0.5,
                                    handlelength=1.5, handletextpad=0.5, fontsize=8, frame=False)

    optnfmt  = {'rows':3,'columns':1,'dpi':300,'showttl':False, 'showoutfnam':True, 'showlogo':False,
                'legendtype':2, 'addmarker':False}

    figfiles2 = chnfobj.xyplots(optnchn2,optnfmt,pltfile2)

    if figfiles1 or figfiles2:
        print 'Plot fils saved:'
        if figfiles1: print '   ', figfiles1[0]
        if figfiles2: print '   ', figfiles2[0]

    if show:
        chnfobj.plots_show()
    else:
        chnfobj.plots_close()

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
    datapath = 'C:\Users\jonathans\PycharmProjects\PSSEDYN\SampleSystem\IN'
    outpath  = 'C:\Users\jonathans\PycharmProjects\PSSEDYN\SampleSystem\OUT'
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
