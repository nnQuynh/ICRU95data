import numpy as np
import polars as pl


dataFolder = './data/photon'


h            = pl.read_csv(f'{dataFolder}/Table_A_5_1b_photon_kerma_h_kermaapprox_extend.csv')
hp           = pl.read_csv(f'{dataFolder}/Table_A_5_2b_photon_kerma_hp_kermaapprox_extend.csv')
dlens        = pl.read_csv(f'{dataFolder}/Table_A_5_3b_photon_kerma_dlens_kermaapprox_extend.csv')
dskin_slab   = pl.read_csv(f'{dataFolder}/Table_A_5_4_1b_photon_kerma_dskin_slab_kermaapprox_extend.csv')
dskin_pillar = pl.read_csv(f'{dataFolder}/Table_A_5_4_2b_photon_kerma_dskin_pillar_kermaapprox.csv')
dskin_rod    = pl.read_csv(f'{dataFolder}/Table_A_5_4_3b_photon_kerma_dskin_rod_kermaapprox.csv')


