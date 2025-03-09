import scipy
import numpy as np
import polars as pl


dosenames = ['ff','dsf']

h = pl.read_csv('./data/Table_A_5_1b_photon_kerma_h_kermaapprox_extend.csv')
h_E = h.select("E_keV").to_numpy().squeeze()
h_coef = h.select("h_Sv/Gy").to_numpy().squeeze()
xout = np.arange(5, 400, 0.5, dtype=float)

temp = scipy.interpolate.CubicSpline(h_E, h_coef, bc_type = 'natural')
yout = temp(xout)