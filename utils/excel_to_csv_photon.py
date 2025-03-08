import pandas as pd


excel_photon = './data/ICRU95 - data - Photon.xlsx'

# h* - fluence
Table_A_1_1a = pd.read_excel(excel_photon, sheet_name='A.1.1a-b', skiprows = 2, usecols = 'C,D')
Table_A_1_1a.dropna(inplace=True)
Table_A_1_1a.rename(columns={Table_A_1_1a.columns[0]: 'E_keV', 
                             Table_A_1_1a.columns[1]: 'h_pSv.cm2'}, 
                             inplace=True)
Table_A_1_1a[['E_keV']] = Table_A_1_1a[['E_keV']].astype('float64')
Table_A_1_1a.to_csv('./data/Table_A_1_1a_photon_fluence_h.csv', index=False)

# h* - kerma
Table_A_1_1b = pd.read_excel(excel_photon, sheet_name='A.1.1a-b', skiprows = 2, usecols = 'H,I')
Table_A_1_1b.dropna(inplace=True)
Table_A_1_1b.rename(columns={Table_A_1_1b.columns[0]: 'E_keV', 
                             Table_A_1_1b.columns[1]: 'h_Sv/Gy'}, 
                             inplace=True)
Table_A_1_1b[['E_keV']] = Table_A_1_1b[['E_keV']].astype('float64')
Table_A_1_1b.to_csv('./data/Table_A_1_1b_photon_kerma_h.csv', index=False)  

# hp - fluence
Table_A_2_1a = pd.read_excel(excel_photon, sheet_name='A.2.1a-b', skiprows = 2, usecols = 'C:O')
Table_A_2_1a.dropna(inplace=True)
Table_A_2_1a.rename(columns={Table_A_2_1a.columns[0]: 'E_keV', 
                             Table_A_2_1a.columns[1]: 'hp_0_pSv.cm2',
                             Table_A_2_1a.columns[2]: 'hp_15',
                             Table_A_2_1a.columns[3]: 'hp_30',
                             Table_A_2_1a.columns[4]: 'hp_45',
                             Table_A_2_1a.columns[5]: 'hp_60',
                             Table_A_2_1a.columns[6]: 'hp_75',
                             Table_A_2_1a.columns[7]: 'hp_90',
                             Table_A_2_1a.columns[8]: 'hp_180',
                             Table_A_2_1a.columns[9]: 'hp_ROT',
                             Table_A_2_1a.columns[10]: 'hp_ISO',
                             Table_A_2_1a.columns[11]: 'hp_SS-ISO',
                             Table_A_2_1a.columns[12]: 'hp_IS-ISO',
                             }, inplace=True)
Table_A_2_1a[['E_keV']] = Table_A_2_1a[['E_keV']].astype('float64')
Table_A_2_1a.to_csv('./data/Table_A_2_1a_photon_fluence_hp.csv', index=False) 

# hp - kerma
Table_A_2_1b = pd.read_excel(excel_photon, sheet_name='A.2.1a-b', skiprows = 2, usecols = 'S:AE')
Table_A_2_1b.dropna(inplace=True)
Table_A_2_1b.rename(columns={Table_A_2_1b.columns[0]: 'E_keV', 
                             Table_A_2_1b.columns[1]: 'hp_0_Sv/Gy',
                             Table_A_2_1b.columns[2]: 'hp_15',
                             Table_A_2_1b.columns[3]: 'hp_30',
                             Table_A_2_1b.columns[4]: 'hp_45',
                             Table_A_2_1b.columns[5]: 'hp_60',
                             Table_A_2_1b.columns[6]: 'hp_75',
                             Table_A_2_1b.columns[7]: 'hp_90',
                             Table_A_2_1b.columns[8]: 'hp_180',
                             Table_A_2_1b.columns[9]: 'hp_ROT',
                             Table_A_2_1b.columns[10]: 'hp_ISO',
                             Table_A_2_1b.columns[11]: 'hp_SS-ISO',
                             Table_A_2_1b.columns[12]: 'hp_IS-ISO',
                             }, inplace=True)
Table_A_2_1b[['E_keV']] = Table_A_2_1b[['E_keV']].astype('float64')
Table_A_2_1b.to_csv('./data/Table_A_2_1b_photon_kerma_hp.csv', index=False) 

# dlens - fluence
Table_A_3_1a = pd.read_excel(excel_photon, sheet_name='A.3.1a-b', skiprows = 2, usecols = 'C:K')
Table_A_3_1a.dropna(inplace=True)
Table_A_3_1a.rename(columns={Table_A_3_1a.columns[0]: 'E_keV', 
                             Table_A_3_1a.columns[1]: 'dlens_0_pGy.cm2',
                             Table_A_3_1a.columns[2]: 'dlens_15',
                             Table_A_3_1a.columns[3]: 'dlens_30',
                             Table_A_3_1a.columns[4]: 'dlens_45',
                             Table_A_3_1a.columns[5]: 'dlens_60',
                             Table_A_3_1a.columns[6]: 'dlens_75',
                             Table_A_3_1a.columns[7]: 'dlens_90',
                             Table_A_3_1a.columns[8]: 'dlens_ROT',
                             }, inplace=True)
Table_A_3_1a[['E_keV']] = Table_A_3_1a[['E_keV']].astype('float64')
Table_A_3_1a.to_csv('./data/Table_A_3_1a_photon_fluence_dlens.csv', index=False) 


# dlens - kerma
Table_A_3_1b = pd.read_excel(excel_photon, sheet_name='A.3.1a-b', skiprows = 2, usecols = 'O:W')
Table_A_3_1b.dropna(inplace=True)
Table_A_3_1b.rename(columns={Table_A_3_1b.columns[0]: 'E_keV', 
                             Table_A_3_1b.columns[1]: 'dlens_0_Sv/Gy',
                             Table_A_3_1b.columns[2]: 'dlens_15',
                             Table_A_3_1b.columns[3]: 'dlens_30',
                             Table_A_3_1b.columns[4]: 'dlens_45',
                             Table_A_3_1b.columns[5]: 'dlens_60',
                             Table_A_3_1b.columns[6]: 'dlens_75',
                             Table_A_3_1b.columns[7]: 'dlens_90',
                             Table_A_3_1b.columns[8]: 'dlens_ROT',
                             }, inplace=True)
Table_A_3_1b[['E_keV']] = Table_A_3_1b[['E_keV']].astype('float64')
Table_A_3_1b.to_csv('./data/Table_A_3_1b_photon_kerma_dlens.csv', index=False) 


# dskin_slab - fluence
Table_A_4_1_1a = pd.read_excel(excel_photon, sheet_name='A.4.1.1a-b', skiprows = 2, usecols = 'C:I')
Table_A_4_1_1a.dropna(inplace=True)
Table_A_4_1_1a.rename(columns={Table_A_4_1_1a.columns[0]: 'E_keV', 
                               Table_A_4_1_1a.columns[1]: 'dskin_slab_0_pGy.cm2',
                               Table_A_4_1_1a.columns[2]: 'dskin_slab_15',
                               Table_A_4_1_1a.columns[3]: 'dskin_slab_30',
                               Table_A_4_1_1a.columns[4]: 'dskin_slab_45',
                               Table_A_4_1_1a.columns[5]: 'dskin_slab_60',
                               Table_A_4_1_1a.columns[6]: 'dskin_slab_75',
                             }, inplace=True)
Table_A_4_1_1a[['E_keV']] = Table_A_4_1_1a[['E_keV']].astype('float64')
Table_A_4_1_1a.to_csv('./data/Table_A_4_1_1a_photon_fluence_dskin_slab.csv', index=False) 

# dskin_slab - kerma
Table_A_4_1_1b = pd.read_excel(excel_photon, sheet_name='A.4.1.1a-b', skiprows = 2, usecols = 'M:S')
Table_A_4_1_1b.dropna(inplace=True)
Table_A_4_1_1b.rename(columns={Table_A_4_1_1b.columns[0]: 'E_keV', 
                               Table_A_4_1_1b.columns[1]: 'dskin_slab_0_Sv/Gy',
                               Table_A_4_1_1b.columns[2]: 'dskin_slab_15',
                               Table_A_4_1_1b.columns[3]: 'dskin_slab_30',
                               Table_A_4_1_1b.columns[4]: 'dskin_slab_45',
                               Table_A_4_1_1b.columns[5]: 'dskin_slab_60',
                               Table_A_4_1_1b.columns[6]: 'dskin_slab_75',
                             }, inplace=True)
Table_A_4_1_1b[['E_keV']] = Table_A_4_1b[['E_keV']].astype('float64')
Table_A_4_1_1b.to_csv('./data/Table_A_4_1_1b_photon_kerma_dskin_slab.csv', index=False) 
  

# dskin_pillar - fluence
Table_A_4_1_2a = pd.read_excel(excel_photon, sheet_name='A.4.1.2a-b', skiprows = 2, usecols = 'C:Q')
Table_A_4_1_2a.dropna(inplace=True)
Table_A_4_1_2a.rename(columns={Table_A_4_1_2a.columns[0]: 'E_keV', 
                               Table_A_4_1_2a.columns[1]: 'dskin_slab_0_pGy.cm2',
                               Table_A_4_1_2a.columns[2]: 'dskin_slab_15',
                               Table_A_4_1_2a.columns[3]: 'dskin_slab_30',
                               Table_A_4_1_2a.columns[4]: 'dskin_slab_45',
                               Table_A_4_1_2a.columns[5]: 'dskin_slab_60',
                               Table_A_4_1_2a.columns[6]: 'dskin_slab_75',
                               Table_A_4_1_2a.columns[7]: 'dskin_slab_90',
                               Table_A_4_1_2a.columns[8]: 'dskin_slab_105',
                               Table_A_4_1_2a.columns[9]: 'dskin_slab_120',
                               Table_A_4_1_2a.columns[10]: 'dskin_slab_135',
                               Table_A_4_1_2a.columns[11]: 'dskin_slab_150',
                               Table_A_4_1_2a.columns[12]: 'dskin_slab_165',
                               Table_A_4_1_2a.columns[13]: 'dskin_slab_180',
                               Table_A_4_1_2a.columns[14]: 'dskin_slab_ROT',
                              }, inplace=True)
Table_A_4_1_2a[['E_keV']] = Table_A_4_1_2a[['E_keV']].astype('float64')
Table_A_4_1_2a.to_csv('./data/Table_A_4_1_2a_photon_fluence_dskin_slab.csv', index=False) 
 

# dskin_pillar - fluence
Table_A_4_1_2a = pd.read_excel(excel_photon, sheet_name='A.4.1.2a-b', skiprows = 2, usecols = 'C:Q')
Table_A_4_1_2a.dropna(inplace=True)
Table_A_4_1_2a.rename(columns={Table_A_4_1_2a.columns[0]: 'E_keV', 
                               Table_A_4_1_2a.columns[1]: 'dskin_slab_0_pGy.cm2',
                               Table_A_4_1_2a.columns[2]: 'dskin_slab_15',
                               Table_A_4_1_2a.columns[3]: 'dskin_slab_30',
                               Table_A_4_1_2a.columns[4]: 'dskin_slab_45',
                               Table_A_4_1_2a.columns[5]: 'dskin_slab_60',
                               Table_A_4_1_2a.columns[6]: 'dskin_slab_75',
                               Table_A_4_1_2a.columns[7]: 'dskin_slab_90',
                               Table_A_4_1_2a.columns[8]: 'dskin_slab_105',
                               Table_A_4_1_2a.columns[9]: 'dskin_slab_120',
                               Table_A_4_1_2a.columns[10]: 'dskin_slab_135',
                               Table_A_4_1_2a.columns[11]: 'dskin_slab_150',
                               Table_A_4_1_2a.columns[12]: 'dskin_slab_165',
                               Table_A_4_1_2a.columns[13]: 'dskin_slab_180',
                               Table_A_4_1_2a.columns[14]: 'dskin_slab_ROT',
                              }, inplace=True)
Table_A_4_1_2a[['E_keV']] = Table_A_4_1_2a[['E_keV']].astype('float64')
Table_A_4_1_2a.to_csv('./data/Table_A_4_1_2a_photon_fluence_dskin_slab.csv', index=False) 


# dskin_pillar - kerma
Table_A_4_1_2b = pd.read_excel(excel_photon, sheet_name='A.4.1.2a-b', skiprows = 2, usecols = 'U:AI')
Table_A_4_1_2b.dropna(inplace=True)
Table_A_4_1_2b.rename(columns={Table_A_4_1_2b.columns[0]: 'E_keV', 
                               Table_A_4_1_2b.columns[1]: 'dskin_slab_0_Sv/Gy',
                               Table_A_4_1_2b.columns[2]: 'dskin_slab_15',
                               Table_A_4_1_2b.columns[3]: 'dskin_slab_30',
                               Table_A_4_1_2b.columns[4]: 'dskin_slab_45',
                               Table_A_4_1_2b.columns[5]: 'dskin_slab_60',
                               Table_A_4_1_2b.columns[6]: 'dskin_slab_75',
                               Table_A_4_1_2b.columns[7]: 'dskin_slab_90',
                               Table_A_4_1_2b.columns[8]: 'dskin_slab_105',
                               Table_A_4_1_2b.columns[9]: 'dskin_slab_120',
                               Table_A_4_1_2b.columns[10]: 'dskin_slab_135',
                               Table_A_4_1_2b.columns[11]: 'dskin_slab_150',
                               Table_A_4_1_2b.columns[12]: 'dskin_slab_165',
                               Table_A_4_1_2b.columns[13]: 'dskin_slab_180',
                               Table_A_4_1_2b.columns[14]: 'dskin_slab_ROT',
                              }, inplace=True)
Table_A_4_1_2b[['E_keV']] = Table_A_4_1_2b[['E_keV']].astype('float64')
Table_A_4_1_2b.to_csv('./data/Table_A_4_1_2b_photon_kerma_dskin_slab.csv', index=False) 

# dskin_rod - fluence
Table_A_4_1_3a = pd.read_excel(excel_photon, sheet_name='A.4.1.3a-b', skiprows = 2, usecols = 'C:Q')
Table_A_4_1_3a.dropna(inplace=True)
Table_A_4_1_3a.rename(columns={Table_A_4_1_3a.columns[0]: 'E_keV', 
                               Table_A_4_1_3a.columns[1]: 'dskin_rod_0_pGy.cm2',
                               Table_A_4_1_3a.columns[2]: 'dskin_rod_15',
                               Table_A_4_1_3a.columns[3]: 'dskin_rod_30',
                               Table_A_4_1_3a.columns[4]: 'dskin_rod_45',
                               Table_A_4_1_3a.columns[5]: 'dskin_rod_60',
                               Table_A_4_1_3a.columns[6]: 'dskin_rod_75',
                               Table_A_4_1_3a.columns[7]: 'dskin_rod_90',
                               Table_A_4_1_3a.columns[8]: 'dskin_rod_105',
                               Table_A_4_1_3a.columns[9]: 'dskin_rod_120',
                               Table_A_4_1_3a.columns[10]: 'dskin_rod_135',
                               Table_A_4_1_3a.columns[11]: 'dskin_rod_150',
                               Table_A_4_1_3a.columns[12]: 'dskin_rod_165',
                               Table_A_4_1_3a.columns[13]: 'dskin_rod_180',
                               Table_A_4_1_3a.columns[14]: 'dskin_rod_ROT',
                              }, inplace=True)
Table_A_4_1_3a[['E_keV']] = Table_A_4_1_3a[['E_keV']].astype('float64')
Table_A_4_1_3a.to_csv('./data/Table_A_4_1_3a_photon_fluence_dskin_rod.csv', index=False) 

# dskin_rod - kerma
Table_A_4_1_3b = pd.read_excel(excel_photon, sheet_name='A.4.1.3a-b', skiprows = 2, usecols = 'U:AI')
Table_A_4_1_3b.dropna(inplace=True)
Table_A_4_1_3b.rename(columns={Table_A_4_1_3b.columns[0]: 'E_keV', 
                               Table_A_4_1_3b.columns[1]: 'dskin_rod_0_Sv/Gy',
                               Table_A_4_1_3b.columns[2]: 'dskin_rod_15',
                               Table_A_4_1_3b.columns[3]: 'dskin_rod_30',
                               Table_A_4_1_3b.columns[4]: 'dskin_rod_45',
                               Table_A_4_1_3b.columns[5]: 'dskin_rod_60',
                               Table_A_4_1_3b.columns[6]: 'dskin_rod_75',
                               Table_A_4_1_3b.columns[7]: 'dskin_rod_90',
                               Table_A_4_1_3b.columns[8]: 'dskin_rod_105',
                               Table_A_4_1_3b.columns[9]: 'dskin_rod_120',
                               Table_A_4_1_3b.columns[10]: 'dskin_rod_135',
                               Table_A_4_1_3b.columns[11]: 'dskin_rod_150',
                               Table_A_4_1_3b.columns[12]: 'dskin_rod_165',
                               Table_A_4_1_3b.columns[13]: 'dskin_rod_180',
                               Table_A_4_1_3b.columns[14]: 'dskin_rod_ROT',
                              }, inplace=True)
Table_A_4_1_3b[['E_keV']] = Table_A_4_1_3b[['E_keV']].astype('float64')
Table_A_4_1_3b.to_csv('./data/Table_A_4_1_3b_photon_fluence_dskin_rod.csv', index=False) 

# h* - fluence - kermaapprox
Table_A_5_1a = pd.read_excel(excel_photon, sheet_name='A.5.1a-b', skiprows = 2, usecols = 'C,D')
Table_A_5_1a.dropna(inplace=True)
Table_A_5_1a.rename(columns={Table_A_5_1a.columns[0]: 'E_keV', 
                             Table_A_5_1a.columns[1]: 'h_pSv.cm2'}, 
                             inplace=True)
Table_A_5_1a[['E_keV']] = Table_A_5_1a[['E_keV']].astype('float64')
Table_A_5_1a.to_csv('./data/Table_A_5_1a_photon_fluence_h_kermaapprox.csv', index=False)

# h* - kerma - kermaapprox
Table_A_5_1b = pd.read_excel(excel_photon, sheet_name='A.5.1a-b', skiprows = 2, usecols = 'C,D')
Table_A_5_1b.dropna(inplace=True)
Table_A_5_1b.rename(columns={Table_A_5_1b.columns[0]: 'E_keV', 
                             Table_A_5_1b.columns[1]: 'h_Sv/Gy'}, 
                             inplace=True)
Table_A_5_1b[['E_keV']] = Table_A_5_1b[['E_keV']].astype('float64')
Table_A_5_1b.to_csv('./data/Table_A_5_1b_photon_kerma_h_kermaapprox.csv', index=False)

# hp - fluence - kermaapprox
Table_A_5_2a = pd.read_excel(excel_photon, sheet_name='A.5.2a-b', skiprows = 2, usecols = 'C:O')
Table_A_5_2a.dropna(inplace=True)
Table_A_5_2a.rename(columns={Table_A_5_2a.columns[0]: 'E_keV', 
                             Table_A_5_2a.columns[1]: 'hp_0_pSv.cm2',
                             Table_A_5_2a.columns[2]: 'hp_15',
                             Table_A_5_2a.columns[3]: 'hp_30',
                             Table_A_5_2a.columns[4]: 'hp_45',
                             Table_A_5_2a.columns[5]: 'hp_60',
                             Table_A_5_2a.columns[6]: 'hp_75',
                             Table_A_5_2a.columns[7]: 'hp_90',
                             Table_A_5_2a.columns[8]: 'hp_180',
                             Table_A_5_2a.columns[9]: 'hp_ROT',
                             Table_A_5_2a.columns[10]: 'hp_ISO',
                             Table_A_5_2a.columns[11]: 'hp_SS-ISO',
                             Table_A_5_2a.columns[12]: 'hp_IS-ISO',
                             }, inplace=True)
Table_A_5_2a[['E_keV']] = Table_A_5_2a[['E_keV']].astype('float64')
Table_A_5_2a.to_csv('./data/Table_A_2_2a_photon_fluence_hp_kermaapprox.csv', index=False) 

# hp - kerma - kermaapprox
Table_A_5_2b = pd.read_excel(excel_photon, sheet_name='A.5.2a-b', skiprows = 2, usecols = 'C:O')
Table_A_5_2b.dropna(inplace=True)
Table_A_5_2b.rename(columns={Table_A_5_2b.columns[0]: 'E_keV', 
                             Table_A_5_2b.columns[1]: 'hp_0_Sv/Gy',
                             Table_A_5_2b.columns[2]: 'hp_15',
                             Table_A_5_2b.columns[3]: 'hp_30',
                             Table_A_5_2b.columns[4]: 'hp_45',
                             Table_A_5_2b.columns[5]: 'hp_60',
                             Table_A_5_2b.columns[6]: 'hp_75',
                             Table_A_5_2b.columns[7]: 'hp_90',
                             Table_A_5_2b.columns[8]: 'hp_180',
                             Table_A_5_2b.columns[9]: 'hp_ROT',
                             Table_A_5_2b.columns[10]: 'hp_ISO',
                             Table_A_5_2b.columns[11]: 'hp_SS-ISO',
                             Table_A_5_2b.columns[12]: 'hp_IS-ISO',
                             }, inplace=True)
Table_A_5_2b[['E_keV']] = Table_A_5_2b[['E_keV']].astype('float64')
Table_A_5_2b.to_csv('./data/Table_A_2_2b_photon_kerma_hp_kermaapprox.csv', index=False) 