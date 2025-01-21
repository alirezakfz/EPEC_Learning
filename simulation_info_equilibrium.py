# -*- coding: utf-8 -*-
"""
HIGH SPEED COMMUNICATION NETWORKS LABORATORY
NATIONAL TECHNICAL UNIVERSITY OF ATHENS

Created on Mon Jan 29 18:31:16 2024

@author: Alireza Khaksari
@email:  alirezakfz@mail.ntua.gr
@mail:   alireza_kfz@yahoo.com
  

"""

import os
import sys
import numpy as np
import random
import math
import pandas as pd
from copy import deepcopy
import pathlib
import os


root_path = os.path.dirname(__file__)
root_path = os.path.abspath(os.path.join(root_path, os.pardir))
sys.path.append(root_path)



from utility.data_utilities import dataset_loader, random_irrediance_solar_power
from utility.plot_util import plot_scenario_info



# 15 November 2019 forecasted temprature

outside_temp = [16.784803,16.094803,15.764802,14.774801,14.834802,14.184802,14.144801,15.314801,16.694803,19.734802,24.414803,25.384802,26.744802,27.144802,27.524803,27.694803,26.834803,26.594803,25.664803,22.594803,21.394802,20.164803,19.584803,20.334803]
outside_temp = [x + 5 if x < 21 else x+2 for x in outside_temp]

# outside_temp = [22.694803, 21.834803, 21.594803, 20.664803, 17.594803, 16.394802, 15.164803, 14.584803, 15.334803, 11.784803, 11.094802999999999, 10.764802, 9.774801, 9.834802, 9.184802, 9.144801, 10.314801, 11.694803, 14.734801999999998, 19.414803, 20.384802, 21.744802, 22.144802, 22.524803]
irrediance_nov = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 101.55, 237.82, 290.98, 224.05, 96.78, 141.85, 60.03, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

irradiance_april = [0, 0, 0, 0, 0, 0, 211, 1200, 3188, 5954, 9317, 6609, 6178, 7082, 5790, 4117, 2321, 1399, 780, 186, 0, 0, 0, 0]

default_grid_info =  os.path.join(root_path,'network_data', '6_Bus_Transmission_Test_System.xlsx')
# root_path = os.path.abspath(os.path.join(os.getcwd(), os.pardir))


class da_profile:
    
    def __init__(self, id, IN_loads, SL_profiles, OCC_profiles, MVA,  load_multiply, conversion_metric, bus):
        self.id = id  # Id of the DA respnsible for managing this profile
        self.bus= bus # Bus number that this profile belongs to it.
        
        # Inflexible loads profiles
        self.IN_loads = IN_loads/3.0 #/10
        
        # Occupancy Profiles
        self.OCC_profiles = OCC_profiles
        
        # EVs properties 
        self.arrival = SL_profiles['Arrival'].to_numpy()
        self.depart  = SL_profiles['Depart'].to_numpy()
        self.charge_power = SL_profiles['EV_Power'].to_numpy()*load_multiply/(conversion_metric*MVA)
        self.EV_soc_low   = SL_profiles['EV_soc_low'].to_numpy()*load_multiply/(conversion_metric*MVA)
        self.EV_soc_up   = SL_profiles['EV_soc_up'].to_numpy()*load_multiply/(conversion_metric*MVA)
        self.EV_soc_arrive = SL_profiles['EV_soc_arr'].to_numpy()*load_multiply/(conversion_metric*MVA)
        self.EV_demand = self.EV_soc_up - self.EV_soc_arrive
         
             
        # Shiftable loads
        # self.SL_baseline_loads=[]
        # self.SL_baseline_loads.append(SL_profiles['SL_loads1'].to_numpy()*load_multiply/(conversion_metric*MVA))#*load_multiply/10)
        # self.SL_baseline_loads.append(SL_profiles['SL_loads2'].to_numpy()*load_multiply/(conversion_metric*MVA))#*load_multiply/10)
        self.SL_loads = SL_profiles['SL_loads'].to_numpy()*load_multiply/(conversion_metric*MVA)
        self.SL_low   = SL_profiles['SL_low'].to_numpy()
        self.SL_up    = SL_profiles['SL_up'].to_numpy()
        self.SL_cycle = SL_profiles['SL_cycle'].to_numpy()
    
        # Thermostatically loads
        self.TCL_R   = SL_profiles['TCL_R']
        self.TCL_C   = SL_profiles['TCL_C']
        self.TCL_COP = SL_profiles['TCL_COP']
        self.TCL_Max = SL_profiles['TCL_MAX']   # This load will scale and convert into MVA in the model power balance
        self.TCL_Beta= SL_profiles['TCL_Beta']
        self.TCL_temp_low = SL_profiles['TCL_temp_low'] 
        self.TCL_temp_up  = SL_profiles['TCL_temp_up'] 
    
    def set_arrival(self, new_arrival):
        self.arrival = new_arrival
    
    def set_depart(self, new_depart):
        self.depart = new_depart
        
    def set_charge_power(self, new_charge_power):
        self.charge_power = new_charge_power
    
    def set_soc_low(self, new_value):
        self.EV_soc_low = new_value
    
    def set_soc_up(self, new_value):
        self.EV_soc_up = new_value
    
    def set_demand(self, new_value):
        self.EV_demand = new_value
    
    def set_soc_arrive(self, new_value):
        self.EV_soc_arrive = new_value
        
    def set_sl_low(self, new_value):
        self.SL_low = new_value
    
    def set_sl_up(self, new_value):
        # print("Set SL Up new value", new_value)
        self.SL_up = new_value
    
    def set_sl_cycle(self, new_value):
        self.SL_cycle = new_value
       
    


class scenario_profiles:
    def __init__(self,
                 no_DAs,                            # Number of demand aggregators in the simulation scenarios
                 no_nodes,                          # Number of nodes controlled by all the DAs. This is equal to number of prosumers.
                 strategic_DA,                      # Default nodes which acts Strategically during simulation scenarios and are controlled by Strategic DA
                 horizon = 24,                      # Simulation Horizon
                 MVA=30,                            # Voltage PU metric
                 irradiance = irradiance_april,     # Value dor solar power generation
                 load_multiply = {1:100},           # Value for number of prosumers pofile for each cluster as defined with no_prosumers
                 conversion_metric=1000,            # Conversion metric from KW to MW for prosumers
                 NO_prosumers = 50,                 # Default number of prosumers clustering profile for each demand aggregator (DA)     
                 network_data = default_grid_info,  # Gid data excell file
                 inf_RES_factor = 1.0,              # Factor to change RES power production
                 inflexible_load_factor = 1.0,      # Factor to scale inflexible loads
                 sl_loads_factor = 0.7,             # Factor to change SL power consumption
                 ev_loads_factor = 1.0,             # Scale EVs charge power
                 plot_simulation_info = False,      # Boolean to plot arrival/departure and SL start/end
                 DAs_dict = None,                   # dictionary from grid data containing bus information of the DAs
                 results_path=None,                 # Path where results of the simelation to be saved in
                 ):
                            
        
        self.MVA = MVA
        self.network_data = network_data
        self.horizon = horizon
        self.NO_prosumers = NO_prosumers
        self.load_multiply = load_multiply
        self.irradiance = irradiance                # Time horizon starts from 1 of today for 24 hours. 
        self.conversion_metric = conversion_metric  # Conversion parmeter to convert KW to MW in prosumers profile
        self.inf_RES_factor = inf_RES_factor        # with this parameter we can control the amount of PVs production if asked to be in simulation scenarios
        self.no_DAs = no_DAs                        # There are c DAs in the simulation
        self.no_nodes = no_nodes                    # There are n nodes/buses as prosumers controlled by c DAs
        self.strategic_DA = strategic_DA            
        self.sl_loads_factor = sl_loads_factor      # Scale shiftable loads
        self.ev_loads_factor = ev_loads_factor      # Scale EVs loads power
        self.inflexible_load_factor = inflexible_load_factor
        self.evs_time_flexibility = 12               # Maximum flexibility and EV can have from its minimum depart time
        self.sl_time_flexibility = 10                # Maximum flexibility an shiftable load can have after its cycle 
        self.plot_simulation_info = plot_simulation_info              # Boolean to plot arrival/departure and SL start/end
        self.DAs_dict = DAs_dict
        self.results_path = results_path
        
        self.V2G_baseline = np.zeros(no_nodes)                        # [0,0,..,0] In baseline scenario none of the prosumers have V2G capability
        self.PV_baseline =  np.ones(no_nodes)*0.2                     # [0,0,..,0] PVs Penetration rate in prosumers.
        self.Peak_RES_Load_Ratio_baseline = np.ones(no_nodes) * 0.4   # [0.2, 0.2,..., 0.2] 
        self.EV_max_flex_baseline = np.ones(no_nodes) * 4             # [4,4,..,4] Charging flexibility of EVs.
        self.SL_max_flex_baseline = np.ones(no_nodes) * 0
        
        self.priceMakers = self.get_strategic_DA_controlled_nodes()   # There are b nodes controlled by strategic DA. Load from network data excel file in 'Structure' sheet.
        self.DAs_profiles = self.load_DAs_profiles()                  # Load profiles for all prosumers in the market. Maximum 20 DAs profiles are available.
        
        self.scenarios = self.generate_scenario_profiles()            # Create separate list of profiles for each DA in each Scenario
        self.save_simulation_info_stats()
        
    def scenario(self, i):
        V2G_rate = 0.5 # V2G Penetration rate in the strategic DA portfolio prosumers
        PV_rate  = 0.5 # Solar Penetration Rate in the strategic DA portfolio prosumers
        
        if(i == 1): # Baseline Scenario
            V2G = np.ones(self.no_nodes) * V2G_rate# np.copy(self.V2G_baseline)
            PV =  np.ones(self.no_nodes) * PV_rate
            Peak_RES_Load_Ratio = np.copy(self.Peak_RES_Load_Ratio_baseline)
            
        
        return V2G, Peak_RES_Load_Ratio, PV
        
    
    def load_DAs_profiles(self):
        das_profiles = dict()
        
        counter = 1
        
        for key, DA in zip(self.DAs_dict.keys(),self.DAs_dict.values()):
            
            for bus in DA.control_buses:
                # Read prosumers data for each DA 
                IN_loads, SL_profiles, OCC_profiles = dataset_loader(str(counter), self.load_multiply[bus], self.NO_prosumers, self.MVA, self.conversion_metric)
                 
            
                # Scale Inflexible loads
                IN_loads *= self.inflexible_load_factor
            
                # Scale Shiftable Loads
                SL_profiles['SL_loads'] *= self.sl_loads_factor
                
            
            
                # Create seperate profile for each DA
                profile = da_profile(id=DA.id, IN_loads= IN_loads, SL_profiles=SL_profiles, OCC_profiles= OCC_profiles, MVA = self.MVA, load_multiply = self.load_multiply[bus] , conversion_metric = self.conversion_metric, bus=bus)
                
                # Scale EVs loads
                profile = self.scale_evs_loads(profile)
            
                # Add current da profile to all created list of DAs profiles
                das_profiles[bus] = profile
                
                # Increament th counter to load next profile from the prosumers profiles folder
                counter += 1 
        return das_profiles
    
    
    
    def generate_scenario_profiles(self):
        
        scenarios = dict()    # Stores profiles for each scenario
        
        for scenario_id in range(1, 2): # There are currently 7 scenarios from 1 to 7
            scenario = dict()           # Scenario information
            V2G, Peak_RES_Load_Ratio, PV = self.scenario(scenario_id)
            
            EVs_list = dict()       # Contains list of prosumers with V2G capability
            Solar_list = dict()     # Contains list of prosumers with PV panels installed
            DA_solar_power = dict() # List representing forecast of solar power production for each prosumer in each DA
            DAs_profiles = dict()   # New modified profile for current scenario
            count_das = 0
            
            for j in self.DAs_profiles.keys():
                temp_profile = deepcopy(self.DAs_profiles[j])  # Load prifle for j DA in the simulation
                
                # Rndomly generate V2G capable prosumers using penetration rate
                # V2G is the scenario penetration rate for 
                v2g_penetration = int( V2G[count_das] * self.NO_prosumers)
                random_prosumers = random.sample([i+1 for i in range(self.NO_prosumers)], k=v2g_penetration )
                EVs_list[j] = random_prosumers
                
                # randomly generate PV installed prosumers usin scenario penetration rate
                # PV[j-1] holds penetration rate of the current DA j in selected scenario
                pv_penetration = int(PV[count_das] * self.NO_prosumers)
                random_prosumers = random.sample([i+1 for i in range(self.NO_prosumers)], k=pv_penetration )
                Solar_list[j] = random_prosumers
                # Based on the list of the prosumers in the Solar List, the DA will have solar power capacity 
                DA_solar_power[j] = self.inf_RES_factor * random_irrediance_solar_power(self.irradiance, self.NO_prosumers,  j, Solar_list[j], self.load_multiply[j], outside_temp, self.MVA, self.conversion_metric)
                DA_solar_power[j] = self.scale_res_based_on_ratio(Peak_RES_Load_Ratio[count_das], DA_solar_power[j], j)
                
                DAs_profiles[j] = deepcopy(temp_profile)
                
                count_das += 1
                
            # Plot scenario info
            if self.plot_simulation_info:
                plot_scenario_info(DAs_profiles, scenario_id)
            
            scenario['profiles']    = DAs_profiles
            scenario['EVs_list']   = EVs_list
            scenario['Solar_list'] = Solar_list
            scenario['DA_solar_power'] =  DA_solar_power
            
            scenarios['scenario'+str(scenario_id)] = scenario
        
        return scenarios
    
    def offer_bid_values_competitve(self):
        c_d_o = dict()
        c_d_b = dict()
        col1 = ['t='+str(i) for i in range(1, self.horizon+1)]
        col2 = ['t='+str(i)+'.1' for i in range(1, self.horizon+1)]
        
        if os.path.isabs(self.network_data):
            path = self.network_data
        else:
            path =  os.path.join(root_path, self.network_data)
            
        
        df = pd.read_excel(path, sheet_name = "CDA Price Offers_Bids")
        d_o = df[col1].to_numpy()
        d_b = df[col2].to_numpy()
        bus_numbers = df["Bus no."].to_numpy()
        
        counter = 0
        for i in bus_numbers:
            c_d_o[i] = d_o[counter]
            c_d_b[i] = d_b[counter]
            counter+=1
            
        return c_d_o, c_d_b
        
    
    def scale_res_based_on_ratio(self, Peak_RES_Load_Ratio, DA_solar_power, j):
        InfLoad_peak = max(self.DAs_profiles[j].IN_loads.sum())
        DA_solar_power = np.array(DA_solar_power)*InfLoad_peak*Peak_RES_Load_Ratio
        return DA_solar_power
        
        
    def scale_evs_loads(self, da_profile):
        da_profile.set_charge_power(da_profile.charge_power * self.ev_loads_factor)
        da_profile.set_soc_up(da_profile.EV_soc_up * self.ev_loads_factor)
        da_profile.set_soc_arrive(da_profile.EV_soc_arrive * self.ev_loads_factor)
        da_profile.set_demand(da_profile.EV_soc_up - da_profile.EV_soc_arrive)
        
        return da_profile
    
    def simulation_info_stats(self):
        # Inflexible loads:
        inf_loads = []
        EVs_loads = []
        SLs_loads = []
        max_inf_loads = []
        max_evs_loads = []
        max_sl_loads = []
        bus_number = []
        da_id = []
        for profile in self.DAs_profiles.values():
            inf_loads.append(sum(profile.IN_loads.sum()) * self.MVA)
            max_inf_loads.append(max(profile.IN_loads.sum()) * self.MVA)
            
            EVs_loads.append(sum(profile.EV_demand) * self.MVA)
            max_evs_loads.append(max(profile.EV_demand) * self.MVA)
            
            loads = sum(profile.SL_loads*profile.SL_cycle) * self.MVA
            SLs_loads.append(loads)
            max_sl_loads.append(max(profile.SL_loads) * self.MVA)
            da_id.append(profile.id)
            bus_number.append(profile.bus)
            
            
        
        print('Total Infelexible Loads are:', inf_loads )
        print('Total EVs demands are:', EVs_loads)
        print('Total SL loads are:',SLs_loads)
        
        data = {"Inflexible_loads": inf_loads,
               "EVs_loads":EVs_loads,
               "Shiftable_loads":SLs_loads,
               "MAX_inf_loads": max_inf_loads,
               "MAX_EVS_loads": max_evs_loads,
               "MAX_SL_loads": max_sl_loads,
               "DA":da_id,
               "Bus":bus_number}
        
        return data
    
    def save_simulation_info_stats(self):
        data = self.simulation_info_stats()    
        
        df = pd.DataFrame.from_dict(data)
        if(self.results_path == None):
            path = os.path.join(root_path, "Results")
        elif(os.path.isabs(self.results_path)):
            path = self.results_path
        else:
            path = os.path.join(root_path, self.results_path)
        
        if not os.path.exists(path):
            os.makedirs(path)
            
        path = os.path.join(path, "DAs_general_info.csv")
        # print(root_path)
        # print(pathlib.Path().resolve())
        df.to_csv(path, index=None)
    
    def get_strategic_DA_controlled_nodes(self):
        """
        There are n nodes in the network. To find nodes (prosumers) that are 
        controlled by strategic DA, we have to look into excel file.
        If in 'Structure' sheet, the number of DA cell is equal to strategic DA
        selected by the user,then add it to price maker list.
        Returns
        -------
        strategic_DA_nodes : List
            DESCRIPTION: prosumers that are controlled by the strategic DA.
        """
        strategic_DA_nodes =[]
        
        if os.path.isabs(self.network_data):
            path = self.network_data
        else:
            path =  os.path.join(root_path, self.network_data)
            
        print(path)
        df = pd.read_excel(path, sheet_name = "Structure")
        for j in range(1, len(df)+1):
            if df['DA'][j-1] == self.strategic_DA:
                strategic_DA_nodes.append(j)
        return strategic_DA_nodes
            