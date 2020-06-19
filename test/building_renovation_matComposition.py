# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 15:15:40 2020

@author: spauliuk
"""

import numpy as np

Sc = np.array([[0,0,0,0,0,0],[0,1,0,0,0,0],[0,1,1,0,0,0],[0,1,1,3,0,0],[0,0,1,3,2,0],[0,0,0,3,2,2]]) # tc
S  = Sc.sum(axis=1)
I  = np.array([1,1,3,2,2]) # t
O  = np.array([[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,1,0,0,0,0],[0,0,1,0,0,0]]) # tc

MB = I - np.diff(Sc,1,axis=0).sum(axis=1) - O[1::,:].sum(axis=1)

# changing MC:
MC = np.array([[1,2,2,2,1,3],[1,2,2,2,2,3],[1,3,2,2,3,3],[1,4,3,2,3,3],[1,4,4,4,3,3],[1,4,4,4,3,3]]) # tc

SM = np.einsum('tc,tc->tc',Sc,MC)
MO = np.einsum('tc,tc->tc',O, MC)
MI = np.einsum('t,tt->t',  I, MC[1::,1::])

MC_diff = np.diff(MC,1,axis=0)

F_ren = np.einsum('tc,tc->t',Sc[1::,:],MC_diff)

# correct mass balance at material layer:
MBal = F_ren + MI - np.diff(SM,1,axis=0).sum(axis=1) - MO[1::,:].sum(axis=1)

I_tot = (np.diff(SM,1,axis=0) + MO[1::,:]).sum(axis=1)
