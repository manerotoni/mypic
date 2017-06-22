# -*- coding: utf-8 -*-
"""
cleanup_notrigger.py
Move directories that do not contain Trigger subdirectory to toDelete.
The direectories are generated during adaptive feedback microscopy workflow (pipelineconstructor)
Created on Thu Jun 02 18:12:02 2016. Modified 23.01.2017
@author: Antonio Politi
"""

import os
import re
import shutil
import Tkinter, Tkconstants, tkFileDialog
Tkinter.Tk().withdraw()
opt = {}
opt['initialdir'] = r'Z:\AntonioP_elltier1\CelllinesImaging\MitoSysPipelines\160309_gfpNUP107z26z31\MitoSys2\LSM'
rootdir = tkFileDialog.askdirectory(**opt)

root, dirs, files = os.walk(rootdir).next()
for adir in dirs:
    m = re.match('DE_W(?P<well>\d+)_P(?P<position>\d+)', adir)
    if  m is not None:
        root2, dirs2, files2 = os.walk(os.path.join(root, adir)).next()
        print(len(dirs2) == 0)
        if len(dirs2) == 0:
            print adir
            try:
                os.mkdir(os.path.join(root, 'toDelete'))
            except:
                pass
            print(adir)
            shutil.move(os.path.join(root, adir), os.path.join(root, 'toDelete'))
            print(adir)     