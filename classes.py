import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import math
import sys
import os

class Generator:
    component_type = {'eletric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance, control):
        self.name_of_instance = name_of_instance

        #list with variables used in this class. This list should not contain variables that are used for connecting elements,
        #e.g. P_from_gen, or P_gen_bat. These variables are generated automatically when connections are constructed. 
        self.list_var = ['gen_op_costs',
                         'gen_emissions',
                         'gen_inv_costs'] # this ist contains variables that appear in the constraints of this class. Conection variables
                                         # should not be mentioned here.
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                               'within = pyo.NonNegativeReals',
                               'within = pyo.NonNegativeReals'] # this list contains the type of variable that has to be created.
                                                                # 

        #these lists are used when transforming parameters into variables, when selecting size_optimization This list should stay empty 
        self.list_altered_var = []
        self.list_text_altered_var = []

        #All parameters that are used in the methods of this class must be given am default value
        self.param_gen_eff = 0.10 # efficency of generation device, usually unitless 
        self.param_gen_size = 50 # size of the generation device. Units can be m2, kW etc. 
        self.param_gen_spec_op_costs = 0.05 # specific operation costs, unit is defined by the rule of accounting costs
        self.param_gen_spec_em = 0.3 # specific emissions, unit is defined by the rule of accounting emissions
        self.param_gen_series = [10] * control.time_span # Units in kWh by default, can be altered if desired
        self.param_gen_spec_inv = 100 # specific emissions, unit is defined by the rule of accounting emissions
        self.param_gen_lifetime = 20 * 8760 # lifetime of the device in hours. After one lifetime, investment cost are repeated
        self.param_gen_compensation = 0.09 # parameter for financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        self.write_gen_series(control) 

    #Every method with a name that start with 'contraint_' will be turned into a constraint of the pyomo model.
    
    #constraint_generation should contain the rule for the generation ef energy of the generation unit. The method below is an example.
    def constraint_generation_rule(model,t):
        return model.P_from_gen[t] == (model.param_gen_series[t] / model.time_step) * model.param_gen_eff * model.param_gen_size
    
    #constraint_operation_costs has the rule for operation cost accounting. The method below is an example.
    def constraint_operation_costs(model,t):
        return model.gen_op_costs[t] == model.param_gen_size * model.param_gen_spec_op_costs
    
    #constraint_emissions_costs has the rule for emissions accounting. The method below is an example.
    def constraint_emissions(model,t):
        return model.gen_emissions[t] == model.P_from_gen[t] * model.param_gen_spec_em
    
    #constraint_investment_costs has the rule for emissions accounting. In this example, investment is accounted in the first time step 
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.gen_inv_costs[t] == model.param_gen_size * model.param_gen_spec_inv
        
        if t % int(model.param_gen_lifetime) == 0:
            return model.gen_inv_costs[t] == model.param_gen_size * model.param_gen_spec_inv
        
        else:
            return model.gen_inv_costs[t] == 0
    
    #write_gen_series is a function that insert the time series for param_gen_series. This will be later used as input for the optimization
    #If input sheet "series" already contains data for param_gen_series, the default value is ignored.
    def write_gen_series(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_' + self.name_of_instance + '_series' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_series': self.param_gen_series})
            df_input_series = pd.concat([df_input_series, df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

class pv(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'no'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['pv_op_costs',
                         'pv_emissions',
                         'pv_inv_costs'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']
        
        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_pv_eff = 0.15 # aproximate overall efficiency of pv cells [-] https://www.eon.de/de/pk/solar/kwp-bedeutung-umrechnung.html
        self.param_pv_area = 100 # total size of pv device [m^2]
        self.param_pv_maintenance = 5 # maintenance costs per kWp and year [€/kWp/yr] (IEP)
        self.param_pv_repair = 10 # repair costs per kWp and year [€/kWp/yr] (IEP)
        self.param_pv_kWp_per_area  = 0.12 # max power of pv device [kWP/m^2]
        self.param_pv_inv_per_kWp = 1000 # invesment cost per kWp of generation [€/kWp] (IEP)
        self.param_pv_lifetime = 25 * 8760 # lifetime of panels in hours [hs]
        self.param_pv_spec_em = 0 # specifc emissions per kWh generated [kgCO2eq/kWh]
        self.param_pv_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        self.param_E_pv_solar = [0.12] * control.time_span # [kWh/m^2] series for solar radiation input, in case no data is given in "series" sheet of input file

        self.write_E_pv_solar(control)

    # write_E_pv_solar writes default solar radiaton to input file in case none is given
    def write_E_pv_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_pv_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    def constraint_generation_rule(model,t):
        return model.P_from_pv[t] == (model.param_E_pv_solar[t] / model.time_step) * model.param_pv_eff * model.param_pv_area 

    def constraint_operation_costs(model,t):
        return model.pv_op_costs[t] == model.param_pv_area * model.param_pv_kWp_per_area * (model.param_pv_maintenance + model.param_pv_repair) / (365 * 24 / model.time_step)
    
    def constraint_emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.param_pv_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pv_inv_costs[t] == model.param_pv_area * model.param_pv_kWp_per_area * model.param_pv_inv_per_kWp
        
        elif t % int(model.param_pv_lifetime) == 0:
            return model.pv_inv_costs[t] == model.param_pv_area * model.param_pv_kWp_per_area * model.param_pv_inv_per_kWp
        
        else:
            return model.pv_inv_costs[t] == 0

class solar_th(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['solar_th_op_costs',
                         'solar_th_emissions',
                         'solar_th_inv_costs']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input 
        self.param_solar_th_eff = 0.30 # aproximate overall efficiency of solar_th cells https://www.solaranlagen-portal.com/solarthermie/kollektoren/flachkollektor (Flachkollektor)
        self.param_solar_th_area = 50 # total size of solar thermal device [m^2] 
        self.param_solar_th_maintenance = 7 # maintenance costs per year and area of installed device [€/m^2] (IEP)
        self.param_solar_th_repair = 3.5 # repair costs per year and per area of installed device [€/m^2] (IEP)
        self.param_solar_th_spec_em = 0 # specifc emissions per kWh generated [kgCO2eq/kWh]
        self.param_solar_th_lifetime = 20 * 8760  # lifetime of panels in years [yrs]
        self.param_solar_th_inv_per_area = 700 # # invesment cost per m^2 of installed device [€/m^2] (IEP)
        self.param_solar_th_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        #default series
        self.param_E_solar_th_solar = [0.12] * control.time_span # kWh/m^2 series for solar irradiation input, in case none is given

        self.write_E_solar_th_solar(control)

    # write_E_solar_th_solar writes default solar radiaton to input file in case none is given
    def write_E_solar_th_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_solar_th_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)
        
    def constraint_generation_rule(model,t):
        return model.Q_from_solar_th[t] == (model.param_E_solar_th_solar[t] * model.time_step) * model.param_solar_th_area * model.param_solar_th_eff
        
    def constraint_operation_costs(model,t):
        return model.solar_th_op_costs[t] == model.param_solar_th_area * (model.param_solar_th_maintenance + model.param_solar_th_repair) / (365 * 24 / model.time_step)
        
    def constraint_emissions(model,t):
        return model.solar_th_emissions[t] == model.Q_from_solar_th[t] * model.param_solar_th_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.solar_th_inv_costs[t] == model.param_solar_th_area * model.param_solar_th_inv_per_area
        
        elif t % int(model.param_solar_th_lifetime) == 0:
            return model.solar_th_inv_costs[t] == model.param_solar_th_area * model.param_solar_th_inv_per_area
        
        else:
            return model.solar_th_inv_costs[t] == 0

class pvt(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['pvt_op_costs',
                         'pvt_emissions',
                         'pvt_inv_costs'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values for scalar parameters
        self.param_pvt_P_eff = 0.20 # approximate overall efficiency of pvt cells for electricity generation [-] https://www.solarserver.de/wissen/basiswissen/hocheffizientes-heizsystem-pvt-kollektor-und-waermepumpe/
        self.param_pvt_Q_eff = 0.40 # approximate overall efficiency of pvt cells for thermal energy generation [-] https://www.solarserver.de/wissen/basiswissen/hocheffizientes-heizsystem-pvt-kollektor-und-waermepumpe/
        self.param_pvt_area = 50 # m^2
        self.param_pvt_maintenace = 8.5 # maintenance costs per m^2 of installed device per year [€/m^2/yr] (IEP)
        self.param_pvt_repair = 4.25 # repair costs per m^2 of installed device per year [€/m^2/yr] (IEP)
        self.param_pvt_spec_em = 0 # specific emission values [kgCO2eq/kWh] 
        self.param_pvt_life_time = 20 * 8760 # lifetime of panels in years, same as PV [yr]
        self.param_pvt_inv_per_area = 850 # investments costs in relation to total area of installed device  [€/m^2] (IEP)
        self.param_pvt_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        #default values for series parameters
        self.param_E_pvt_solar = [0.12] * control.time_span # kWh/m^2 series for solar irradiation input, in case none is given

        self.write_E_pvt_solar(control)

    def write_E_pvt_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_pvt_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)
        
    def constraint_generation_rule(model,t):
        return model.P_from_pvt[t] == (model.param_E_pvt_solar[t] * model.time_step) * model.param_pvt_area * model.param_pvt_P_eff
    
    def constraint_generation_rule_thermal(model,t):
        return model.Q_from_pvt[t] == (model.param_E_pvt_solar[t] * model.time_step) * model.param_pvt_area * model.param_pvt_Q_eff
        
    def constraint_operation_costs(model,t):
        return model.pvt_op_costs[t] == model.param_pvt_area * (model.param_pvt_maintenace + model.param_pvt_repair) / (365 * 24 /  model.time_step)
        
    def constraint_emissions(model,t):
        return model.pvt_emissions[t] == (model.Q_from_pvt[t] + model.P_from_pvt[t]) * model.param_pvt_spec_em

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pvt_inv_costs[t] == model.param_pvt_area * model.param_pvt_inv_per_area
        
        elif t % int(model.param_pvt_life_time) == 0:
            return model.pvt_inv_costs[t] == model.param_pvt_area * model.param_pvt_inv_per_area
        
        else:
            return model.pvt_inv_costs[t] == 0
        
class CHP(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['CHP_op_costs',
                         'CHP_emissions',
                         'CHP_inv_costs',
                         'CHP_fuel_cons',
                         'CHP_K',
                         'CHP_z1',
                         'CHP_z2']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.Binary',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values for scalar parameters
        self.param_P_CHP_max = 20 # max power of CHP device [kW]
        self.param_P_CHP_min = 0.2 * self.param_P_CHP_max # min operation power of CHP [kW]
        self.param_CHP_P_to_Q_ratio = 0.5 # In german, Stromkennzahl, relation Pel/Pth, [-] https://www.energieatlas.bayern.de/thema_energie/kwk/anlagentypen
        self.param_CHP_eff = 0.8 # Overall efficiency of CHP [-] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden.
        self.param_CHP_fuel_price = 0.09463 # cost per kWh of fuel consumed [€/kWh] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden
        self.param_CHP_spec_em = 0.200 # emissions due to burning of 1 kWh of the fuel [kgCO2eq/kWh] https://www.ris.bka.gv.at/GeltendeFassung.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20008075
        self.param_CHP_inv_costs_per_power = 1700 # investments costs per electric power https://www.heizungsfinder.de/bhkw/kosten-preise/anschaffungskosten
        self.param_CHP_maintenance_costs = 0.04 # costs of maintenance per generated kWh [€/kWh] https://partner.mvv.de/blog/welche-bhkw-kosten-fallen-in-der-anschaffung-und-beim-betrieb-an-bhkw#:~:text=Wartung%20und%20Bedienung,75%20Cent%20pro%20kWh%20rechnen.
        self.param_CHP_bonus = 0.09 # KWK-Bonus, compensation for energy from CHP sold to net [€/kWh] https://www.heizungsfinder.de/bhkw/wirtschaftlichkeit/einspeiseverguetung#3     https://www.bhkw-infozentrum.de/wirtschaftlichkeit-bhkw-kwk/ueblicher_preis_bhkw.html
        self.param_CHP_not_used_energy_compensation = 0.01 #  compensaton for decentralised energy generation [€/kWh] https://www.heizungsfinder.de/bhkw/wirtschaftlichkeit/einspeiseverguetung#3
        self.param_CHP_compensation = self.param_CHP_bonus + self.param_CHP_not_used_energy_compensation # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]
        self.param_CHP_lifetime = 20 * 8760 # total lifespan of device in hours [hs]

        self.param_CHP_M1 = 10000
        self.param_CHP_M2 = 10000


    # linearized equations to define min power limit. Original equation was: model.P_from_CHP[t] >= model.param_P_CHP_min * model.CHP_K[t] https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    def constraint_min_generation1(model,t):
        return model.P_from_CHP[t] >= model.CHP_z1[t]
    
    def constraint_min_generation2(model,t):
        return model.CHP_z1[t] <= model.param_CHP_M1 * model.CHP_K[t]
    
    def constraint_min_generation3(model,t):
        return model.CHP_z1[t] <= model.param_P_CHP_min 
    
    def constraint_min_generation4(model,t):
        return model.CHP_z1[t] >= model.param_P_CHP_min - (1 - model.CHP_K[t]) * model.param_CHP_M1
    
    def constraint_min_generation5(model,t):
        return model.CHP_z1[t] >= 0



    # linearized equations to define max power limit. Original equation was: model.P_from_CHP[t] <= model.param_P_CHP_max * model.CHP_K[t] https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    def constraint_max_generation1(model,t):
        return model.P_from_CHP[t] <= model.CHP_z2[t]
    
    def constraint_max_generation2(model,t):
        return model.CHP_z2[t] <= model.param_CHP_M2 * model.CHP_K[t]
    
    def constraint_max_generation3(model,t):
        return model.CHP_z2[t] <= model.param_P_CHP_max
    
    def constraint_max_generation4(model,t):
        return model.CHP_z2[t] >= model.param_P_CHP_max - (1 - model.CHP_K[t]) * model.param_CHP_M2
    
    def constraint_max_generation5(model,t):
        return model.CHP_z2[t] >= 0
    


    def constraint_generation_rule(model,t):
        return model.P_from_CHP[t] == model.Q_from_CHP[t] * model.param_CHP_P_to_Q_ratio 
    
    def constraint_fuel_consumption(model,t):
        return model.CHP_fuel_cons[t] == (model.P_from_CHP[t] + model.Q_from_CHP[t]) / model.param_CHP_eff
    
    def constraint_operation_costs(model,t):
        return model.CHP_op_costs[t] == model.CHP_fuel_cons[t] * model.param_CHP_fuel_price + model.CHP_fuel_cons[t] * model.param_CHP_eff * model.param_CHP_maintenance_costs
    
    def constraint_emissions(model,t):
        return model.CHP_emissions[t] == model.CHP_fuel_cons[t] * model.param_CHP_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.CHP_inv_costs[t] == model.param_P_CHP_max * model.param_CHP_inv_costs_per_power
        
        elif t % int(model.param_CHP_lifetime) == 0:
            return model.CHP_inv_costs[t] == model.param_P_CHP_max * model.param_CHP_inv_costs_per_power
        
        else:
            return model.CHP_inv_costs[t] == 0

class gas_boiler(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['gas_boiler_fuel_cons',
                         'gas_boiler_op_costs',
                         'gas_boiler_emissions',
                         'gas_boiler_inv_costs',
                         'gas_boiler_z1',
                         'gas_boiler_z2',
                         'gas_boiler_K'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.Binary']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values for scalar parameters
        self.param_Q_gas_boiler_max = 20 # max power that can be generated with this device [kW]
        self.param_Q_gas_boiler_min = 0.2 * self.param_Q_gas_boiler_max # min power limitation when this device is in operation [kW]
        self.param_gas_boiler_eff = 0.95 # efficency when converting fuel into thermal energy [-] 
        # self.param_gas_boiler_fuel_cons_ratio = 90.09 #dm3 per kWh of P_from_CHP * 1h  
        self.param_gas_boiler_fuel_price = 0.09463 # cost per kWh of fuel consumed [€/kWh] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden
        self.param_gas_boiler_spec_em = 0.200 # emissions due to burning of 1 kWh of the fuel [kgCO2eq/kWh] https://www.ris.bka.gv.at/GeltendeFassung.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20008075
        self.param_gas_boiler_maintenance = 1875 #maintenace costs per kW of installed capacity per  year [€/kW/yr] (IEP)
        self.param_gas_boiler_repair = 1875 #repair costs per kW of installed capacity per  year [€/kW/yr] (IEP)
        self.param_gas_boiler_inv_costs_per_power = 125 # investment costs per max kW of installed device [€/kW]
        self.param_gas_boiler_lifetime = 20 * 8760 # total life span of the device [hs]
        self.param_gas_boiler_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]
        
        self.param_gas_boiler_M1 = 10000
        self.param_gas_boiler_M2 = 10000
    

    # linearized equations to define min power limit. Original equation was: model.Q_from_gas_boiler[t] >= model.param_Q_gas_boiler_min https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    
    def constraint_min_generation1(model,t):
        return model.Q_from_gas_boiler[t] >= model.gas_boiler_z1[t]
    
    def constraint_min_generation2(model,t):
        return model.gas_boiler_z1[t] <= model.param_gas_boiler_M1 * model.gas_boiler_K[t]
    
    def constraint_min_generation3(model,t):
        return model.gas_boiler_z1[t] <= model.param_Q_gas_boiler_min 
    
    def constraint_min_generation4(model,t):
        return model.gas_boiler_z1[t] >= model.param_Q_gas_boiler_min - (1 - model.gas_boiler_K[t]) * model.param_gas_boiler_M1
    
    def constraint_min_generation5(model,t):
        return model.gas_boiler_z1[t] >= 0


    # linearized equations to define max power limit. Original equation was: model.Q_from_gas_boiler[t] <= model.param_Q_gas_boiler_max https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    
    def constraint_max_generation1(model,t):
        return model.Q_from_gas_boiler[t] <= model.gas_boiler_z2[t]
    
    def constraint_max_generation2(model,t):
        return model.gas_boiler_z2[t] <= model.param_gas_boiler_M2 * model.gas_boiler_K[t]
    
    def constraint_max_generation3(model,t):
        return model.gas_boiler_z2[t] <= model.param_Q_gas_boiler_max 
    
    def constraint_max_generation4(model,t):
        return model.gas_boiler_z2[t] >= model.param_Q_gas_boiler_max - (1 - model.gas_boiler_K[t]) * model.param_gas_boiler_M2
    
    def constraint_max_generation5(model,t):
        return model.gas_boiler_z2[t] >= 0
    


    def constraint_generation_rule(model,t):
        return model.Q_from_gas_boiler[t] == model.param_gas_boiler_eff *  model.gas_boiler_fuel_cons[t] / model.time_step

    def constraint_operation_costs(model,t):
        return model.gas_boiler_op_costs[t] == (model.gas_boiler_fuel_cons[t] * model.param_gas_boiler_fuel_price +
                                               model.param_Q_gas_boiler_max * (model.param_gas_boiler_maintenance + model.param_gas_boiler_repair) / (365 * 24 / model.time_step))
    
    def constraint_emissions(model,t):
        return model.gas_boiler_emissions[t] == model.gas_boiler_fuel_cons[t] * model.param_gas_boiler_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.gas_boiler_inv_costs[t] == model.param_Q_gas_boiler_max * model.param_gas_boiler_inv_costs_per_power
        
        elif t % int(model.param_gas_boiler_lifetime) == 0:
            return model.gas_boiler_inv_costs[t] == model.param_Q_gas_boiler_max * model.param_gas_boiler_inv_costs_per_power
        
        else:
            return model.gas_boiler_inv_costs[t] == 0

class heat_pump(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self, name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['heat_pump_emissions',
                         'heat_pump_inv_costs',
                         'heat_pump_op_costs',
                         'heat_pump_K',
                         'heat_pump_z1',
                         'heat_pump_z2']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.Binary',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values for scalar parameters
        self.param_P_heat_pump_max = 20 # max possible electric power consumption [kW]
        self.param_P_heat_pump_min = 0.2 * self.param_P_heat_pump_max # minimum power consumption for device while in operation [kW] 
        self.param_heat_pump_COP = 4 # overall average coefficent of performance, Pth/Pel [-] https://www.heizungsfinder.de/waermepumpe/wirtschaftlichkeit/cop-wert
        self.param_heat_pump_spec_em = 0 # specific emissions per consumed kWh of electricity [kgCO2eq/kWh] https://gshp.org.uk/gshps/what-are-gshps/#:~:text=element%20of%20a%20ground%20source,life%20of%20over%20100%20years.&text=Unlike%20burning%20oil%2C%20gas%2C%20LPG,is%20used%20to%20power%20them).
        self.param_heat_pump_spec_inv_costs = 733 # investment costs per kW of installed capacity (IEP)
        self.param_heat_pump_maintenance = 11 # maintenance costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_repair = 11 # repair costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_spec_op_costs = 12 # operation costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_lifetime = 20 * 8760 # total lifespan of device [hs]
        self.param_heat_pump_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        self.param_heat_pump_M1 = 10000
        self.param_heat_pump_M2 = 10000


    def constraint_generation_rule(model,t):
        return model.Q_from_heat_pump[t] == model.P_to_heat_pump[t] * model.param_heat_pump_COP
    
    
    # linearized equations to define max power limit. Original equation was: model.Q_from_heat_pump[t] == model.P_to_heat_pump[t] * model.param_heat_pump_COP https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable

    def constraint_max_power1(model,t):
        return model.P_to_heat_pump[t] <= model.heat_pump_z1[t]
    
    def constraint_max_power2(model,t):
        return model.heat_pump_z1[t] <= model.param_heat_pump_M1 * model.heat_pump_K[t]
    
    def constraint_max_power3(model,t):
        return model.heat_pump_z1[t] <= model.param_P_heat_pump_max 
    
    def constraint_max_power4(model,t):
        return model.heat_pump_z1[t] >= model.param_P_heat_pump_max - (1 - model.heat_pump_K[t]) * model.param_heat_pump_M1
    
    def constraint_max_power5(model,t):
        return model.heat_pump_z1[t] >= 0


    # linearized equations to define min power limit. Original equation was: model.P_to_heat_pump[t] >= model.param_P_heat_pump_min * model.heat_pump_K[t] https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    
    def constraint_min_power1(model,t):
        return model.P_to_heat_pump[t] >= model.heat_pump_z2[t]
    
    def constraint_min_power2(model,t):
        return model.heat_pump_z2[t] <= model.param_heat_pump_M2 * model.heat_pump_K[t]
    
    def constraint_min_power3(model,t):
        return model.heat_pump_z2[t] <= model.param_P_heat_pump_min
    
    def constraint_min_power4(model,t):
        return model.heat_pump_z2[t] >= model.param_P_heat_pump_min - (1 - model.heat_pump_K[t]) * model.param_heat_pump_M2
    
    def constraint_min_power5(model,t):
        return model.heat_pump_z2[t] >= 0


    
    def constraint_operation_costs(model,t):
        return model.heat_pump_op_costs[t] == model.param_P_heat_pump_max * (model.param_heat_pump_maintenance + model.param_heat_pump_spec_op_costs + model.param_heat_pump_repair) / (365 * 24 / model.time_step)

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.heat_pump_inv_costs[t] == model.param_heat_pump_spec_inv_costs * model.param_P_heat_pump_max
        
        elif t % int(model.param_heat_pump_lifetime) == 0:
            return model.heat_pump_inv_costs[t] == model.param_heat_pump_spec_inv_costs * model.param_P_heat_pump_max

        else:
            return model.heat_pump_inv_costs[t] == 0

    def constraint_emissions(model,t): 
        return model.heat_pump_emissions[t] == model.P_to_heat_pump[t] * model.param_heat_pump_spec_em



class Storage:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'no',
                      'thermal_source':'yes'}
    
    def __init__(self, name_of_instance,control):
        self.name_of_instance = name_of_instance
    
        self.list_var = ['storage_emissions',
                         'storage_inv_costs',
                         'storage_op_costs']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_storage_eff = 4
        self.param_storage_spec_op_costs = 0.05
        self.param_storage_inv_costs = 10000
        self.param_storage_spec_em = 0.1
        self.param_storage_lifetime = 20 * 8760


    def constraint_function_rule(model,t):
        return model.Q_from_storage[t] == model.P_to_storage[t] * model.param_storage_eff
    
    def constraint_operation_costs(model,t):
        return model.storage_op_costs[t] == model.param_storage_spec_op_costs * model.time_step

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.storage_inv_costs[t] == model.param_storage_inv_costs
        if t % int (model.param_storage_lifetime) == 0:
            return model.storage_inv_costs[t] == model.param_storage_inv_costs
        else:
            return model.storage_inv_costs[t] == 0

    def constraint_emissions(model,t): 
        return model.storage_emissions[t] == model.P_to_storage[t] * model.param_storage_spec_em
    
class bat(Storage):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'no'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['bat_energy',
                         'bat_K_ch',
                         'bat_K_dis',
                         'bat_op_costs',
                         'bat_emissions',
                         'bat_inv_costs',
                         'bat_z1',
                         'bat_z2'] #no powers
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'domain = pyo.Binary',
                              'domain = pyo.Binary',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']
        
        self.list_altered_var = []
        self.list_text_altered_var =[]
        
        #default values for scalar parameters
        self.param_bat_E_max_initial = 100 #max energy capacity of the battery [kWh]
        self.param_bat_starting_SOC = 0.5 # starting state of charge of the battery [-]
        self.param_bat_ch_eff = 0.95 # efficiency for charging the battery [-]
        self.param_bat_dis_eff = 0.95 # efficiency for discharging the battery [-]
        self.param_bat_c_rate_ch = 1 # c-rate of the battery for charging. Specifices the max power of charging in relation to its capacity [1/hr]
        self.param_bat_c_rate_dis = 1 # c-rate of the battery for discharging. Specifices the max power of discharging in relation to its capacity [1/hr]
        self.param_bat_spec_op_costs = 0 # specifies the operation cost of the battery per kWh flow in the battery [€/kWh]
        self.param_bat_spec_em = 50 # embodied emission per kWh capacity of the battery in EU. file: 1-s2.0-S0921344922004402-main [kgCO2eq/kWh]
        self.param_bat_DoD = 0.7 # maximum depth of discharge of the battery [-]
        self.param_bat_inv_per_capacity = 650 # investment costs per installed kWh of capacity. [€/kWh]
        self.param_bat_cycles = 9000 # max number of full cyces before battery is worn out beyond operation [# cycles]
        self.param_bat_lifetime = 10 * 8760 # total lifespan of the device in hours [hs]
        self.param_bat_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        self.param_bat_energy_starting_index = 1 # parameter created to connect energy of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_inv_costs_starting_index = 1 # parameter created to connect investment costs with previous optimization horizon when using receding horizon optimization

        self.param_bat_M1 = 10000 # parameter created for linearization using Big M method.
        self.param_bat_M2 = 10000 # parameter created for linearization using Big M method.

        if control.receding_horizon == 'yes':
            self.param_bat_receding_horizon = 1
        else:
            self.param_bat_receding_horizon = 0

    def constraint_depth_of_discharge(model,t):
        return model.bat_energy[t] >= (1 - model.param_bat_DoD) * model.param_bat_E_max_initial
    
    def constraint_max_state_of_charge(model,t):
        return model.bat_energy[t] <= model.param_bat_E_max_initial
        
    def constraint_function_rule(model,t):
        if t == 1:
            return model.bat_energy[t] == model.param_bat_E_max_initial * model.param_bat_starting_SOC + (model.P_to_bat[t] * model.param_bat_ch_eff 
                                                                - model.P_from_bat[t] / model.param_bat_dis_eff) * model.time_step
         
        elif t == model.starting_index and model.param_bat_receding_horizon == 1:
            return model.bat_energy[t] == model.param_bat_energy_starting_index
        
        else:
            return model.bat_energy[t] == model.bat_energy[t-1] + (model.P_to_bat[t] * model.param_bat_ch_eff 
                                                                - model.P_from_bat[t] / model.param_bat_dis_eff) * model.time_step
        

    # linearized equations to define charging power limit. Original equation was: model.P_to_bat[t] <= model.param_bat_E_max_initial * model.bat_K_ch[t] * model.param_bat_c_rate_ch https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable
    def constraint_charge_limit1(model,t):
        return model.P_to_bat[t] <= model.param_bat_c_rate_ch * model.bat_z1[t]

    def constraint_charge_limit2(model,t):
        return model.bat_z1[t] <= model.param_bat_M1 * model.bat_K_ch[t]

    def constraint_charge_limit3(model,t):
        return model.bat_z1[t] <= model.param_bat_E_max_initial

    def constraint_charge_limit4(model,t):
        return model.bat_z1[t] >= model.param_bat_E_max_initial - (1 - model.bat_K_ch[t]) * model.param_bat_M1
    
    def constraint_charge_limit5(model,t):
        return model.bat_z1[t] >= 0 


    # linearized equations to define discharging power limit. Original equation was: model.P_from_bat[t] <= model.param_bat_E_max_initial *  model.bat_K_dis[t] * model.param_bat_c_rate_dis https://or.stackexchange.com/questions/39/how-to-linearize-the-product-of-a-binary-and-a-non-negative-continuous-variable

    def constraint_discharge_limit1(model,t):
        return model.P_from_bat[t] <= model.param_bat_c_rate_dis * model.bat_z2[t] 

    def constraint_discharge_limit2(model,t):
        return model.bat_z2[t] <= model.param_bat_M2 * model.bat_K_dis[t]

    def constraint_discharge_limit3(model,t):
        return model.bat_z2[t] <= model.param_bat_E_max_initial

    def constraint_discharge_limit4(model,t):
        return model.bat_z2[t] >= model.param_bat_E_max_initial - (1 - model.bat_K_dis[t]) * model.param_bat_M2
    
    def constraint_discharge_limit5(model,t):
        return model.bat_z2[t] >= 0
    

    # constraint_keys_rule prevents battery from charging and discharging in the same time step
    def constraint_keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1

    def constraint_operation_costs(model,t):
        return model.bat_op_costs[t] == (model.P_to_bat[t] + model.P_from_bat[t]) * model.param_bat_spec_op_costs
    
    #emissions equation accounts for battery embodied CO2 emissions 
    def constraint_emissions(model,t):
        return model.bat_emissions[t] == (model.P_from_bat[t] + model.P_to_bat[t]) * model.param_bat_spec_em / (model.param_bat_cycles*2)
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.bat_inv_costs[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity
        
        elif t % int(model.param_bat_lifetime) == 0:
            return model.bat_inv_costs[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity
        
        else:
            return model.bat_inv_costs[t] == 0

class bat_with_aging(Storage):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'no'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['bat_with_aging_SOC',
                         'bat_with_aging_K_ch',
                         'bat_with_aging_K_dis',
                         'bat_with_aging_op_costs',
                         'bat_with_aging_emissions',
                         'bat_with_aging_SOC_max',
                         'bat_with_aging_integer',
                         'bat_with_aging_cumulated_aging',
                         'bat_with_aging_inv_costs'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',
                              'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'within = pyo.Integers', 'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values for scalar parameters
        self.param_bat_with_aging_E_max_initial = 100 # max stored energy of device at the beginning of lifetime [kWh]
        self.param_bat_with_aging_starting_SOC = 0.5 # initial default state of charge of device [-]
        self.param_bat_with_aging_ch_eff = 0.95 # efficiency for charging device [-]
        self.param_bat_with_aging_dis_eff = 0.95 # efficiency for discharging device [-]
        self.param_bat_with_aging_c_rate_ch = 1 # c-rate of the battery for charging. Specifices the max power of charging in relation to its capacity [1/hr]
        self.param_bat_with_aging_c_rate_dis = 1 # c-rate of the battery for discharging. Specifices the max power of discharging in relation to its capacity [1/hr]
        self.param_bat_with_aging_spec_op_costs = 0.01 # specifies the operation cost of the battery per kWh flow in the battery [€/kWh]
        self.param_bat_with_aging_spec_em = 50 # embodied emission per kWh capacity of the battery in EU. file: 1-s2.0-S0921344922004402-main [kgCO2eq/kWh]
        self.param_bat_with_aging_DoD = 0.7 # max depth of discharge of energy without damaging battery [-]
        self.param_bat_with_aging_final_SoH = 0.7 # max state of charge of battery at the end of lifetime, state of heatlh [-]
        self.param_bat_with_aging_cycles = 9000 # full cycles before final SoH is reached and battery is replaced [# cycles]
        self.param_bat_with_aging_aging = (self.param_bat_with_aging_E_max_initial * (1 -self.param_bat_with_aging_final_SoH)) / (self.param_bat_with_aging_cycles * 2 * self.param_bat_with_aging_E_max_initial) / self.param_bat_with_aging_E_max_initial # state of charge lost for kWh going through the battery.
        self.param_bat_with_aging_inv_per_capacity = 650 # investment costs of battery per kWh of installed capacity [kWh]
        self.param_bat_with_aging_compensation = 0.01 # financial compensation for each kWh sold to the net. This value is accounted additionally to the market's spot price [€/kWh]

        self.param_bat_with_aging_SOC_starting_index = 1 # parameter created to connect state of charge of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_with_aging_cumulated_aging_starting_index = 1 # parameter created to connect accumulated aging of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_with_aging_inv_costs_starting_index = 1 # parameter created to connect investment costs of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_with_aging_integer_starting_index = 1 # parameter created to connect logical variable of the battery with previous optimization horizon when using receding horizon optimization
        
        if control.receding_horizon == 'yes':
            self.param_bat_with_aging_receding_horizon = 1
        else:
            self.param_bat_with_aging_receding_horizon = 0

        #defining energy type to build connections with other componets correctly

    def constraint_depth_of_discharge(model,t):
        return model.bat_with_aging_SOC[t] >= (1 - model.param_bat_with_aging_DoD)
    
    def constraint_max_state_of_charge(model,t):
        if t == 1:
            return model.bat_with_aging_SOC[t] <= 1
        else:
            return model.bat_with_aging_SOC[t] <= model.bat_with_aging_SOC_max[t]
    
    def constraint_cumulated_aging(model,t):
        if t == 1:
            return model.bat_with_aging_cumulated_aging[t] == (model.P_from_bat_with_aging[t] + 
                                                    model.P_to_bat_with_aging[t]) * model.time_step * model.param_bat_with_aging_aging
        
        elif t == model.starting_index and model.param_bat_with_aging_receding_horizon == 1:

            return model.bat_with_aging_cumulated_aging[t] == model.param_bat_with_aging_cumulated_aging_starting_index
        
        else:
            return model.bat_with_aging_cumulated_aging[t] == model.bat_with_aging_cumulated_aging[t-1] + (model.P_from_bat_with_aging[t] + 
                                                                                     model.P_to_bat_with_aging[t]) * model.time_step * model.param_bat_with_aging_aging
        
    def constraint_upper_integer_rule(model,t):
        if t == 1:
            return model.bat_with_aging_integer[t] == 0 
        else:
            return model.bat_with_aging_integer[t] <= model.bat_with_aging_cumulated_aging[t] / (1-model.param_bat_with_aging_final_SoH)
    
    def constraint_lower_integer_rule(model,t):
        if t == 1:
            return model.bat_with_aging_integer[t] == 0
        
        elif t == model.starting_index and model.param_bat_with_aging_receding_horizon == 1:
            return model.bat_with_aging_integer[t] == model.param_bat_with_aging_integer_starting_index
        
        else:
            return model.bat_with_aging_integer[t] >= model.bat_with_aging_cumulated_aging[t] / (1-model.param_bat_with_aging_final_SoH) - 1
     
    def constraint_aging(model,t):
        if t == 1:
            return model.bat_with_aging_SOC_max[t] == 1
        else:
            return model.bat_with_aging_SOC_max[t] == 1 - model.bat_with_aging_cumulated_aging[t] + model.bat_with_aging_integer[t] * (1 - model.param_bat_with_aging_final_SoH)
    
    def constraint_function_rule(model,t):
        if t == 1:
            return model.bat_with_aging_SOC[t] == model.param_bat_with_aging_starting_SOC + (model.P_to_bat_with_aging[t] * model.param_bat_with_aging_ch_eff 
                                                                - model.P_from_bat_with_aging[t]/model.param_bat_with_aging_dis_eff) * model.time_step / model.param_bat_with_aging_E_max_initial
        elif t == model.starting_index and model.param_bat_with_aging_receding_horizon == 1:
            return model.bat_with_aging_SOC[t] == model.param_bat_with_aging_SOC_starting_index
        else:
            return model.bat_with_aging_SOC[t] == model.bat_with_aging_SOC[t-1] + (model.P_to_bat_with_aging[t] * model.param_bat_with_aging_ch_eff 
                                                                - model.P_from_bat_with_aging[t]/model.param_bat_with_aging_dis_eff) * model.time_step / model.param_bat_with_aging_E_max_initial

    def constraint_charge_limit(model,t):
        return model.P_to_bat_with_aging[t] <= model.param_bat_with_aging_E_max_initial * model.bat_with_aging_K_ch[t] * model.param_bat_with_aging_c_rate_ch

    def constraint_discharge_limit(model,t):
        return model.P_from_bat_with_aging[t] <= model.param_bat_with_aging_E_max_initial *  model.bat_with_aging_K_dis[t] * model.param_bat_with_aging_c_rate_dis
    
    def constraint_keys_rule(model,t):
        return model.bat_with_aging_K_ch[t] + model.bat_with_aging_K_dis[t] <= 1

    def constraint_operation_costs(model,t):
        return model.bat_with_aging_op_costs[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_spec_op_costs
    
    def constraint_emissions(model,t):
        return model.bat_with_aging_emissions[t] == (model.P_from_bat_with_aging[t] + model.P_to_bat_with_aging[t]) * model.param_bat_with_aging_spec_em/(model.param_bat_with_aging_cycles*2)
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.bat_with_aging_inv_costs[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_inv_per_capacity
        
        elif t == model.starting_index and model.param_bat_with_aging_receding_horizon == 1:
            return model.bat_with_aging_inv_costs[t] == model.param_bat_with_aging_inv_costs_starting_index
        
        else:
            return model.bat_with_aging_inv_costs[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_inv_per_capacity * (model.bat_with_aging_integer[t] - model.bat_with_aging_integer[t-1])



class Consumer:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'yes',
                      'thermal_source':'no'}
    
    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance

        self.list_var = ['cons_inv_costs',
                         'cons_op_costs']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_P_to_cons = [500] * control.time_span # this is tranformed into a series in case user does not give input

        self.write_P_to_cons(control)

    def constraint_investment_costs(model,t):
        return model.cons_inv_costs[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.cons_op_costs[t] == 0
    
    def write_P_to_cons(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_P_to_' + self.name_of_instance: self.param_P_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

class demand(Consumer):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'yes',
                      'thermal_source':'no'}

    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance
        
        self.list_var = ['demand_inv_costs',
                         'demand_op_costs']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #Setting up default values for series if none are given in input file:
        self.param_P_to_demand = [2.5] * control.time_span # default electric power of demand that needs to be covered [kW]
        self.param_Q_to_demand = [7] * control.time_span # default thermal power of demand that needs to be covered  [kW]]

        self.write_P_to_demand(control)
        self.write_Q_to_demand(control)

    def write_P_to_demand(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_P_to_' + self.name_of_instance: self.param_P_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    def write_Q_to_demand(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_Q_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_Q_to_' + self.name_of_instance: self.param_Q_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    def constraint_investment_costs(model,t):
        return model.demand_inv_costs[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.demand_op_costs[t] == 0

class charging_station(Consumer):

    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'no'}

    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance

        #default values in case of no input
        # self.list_var = ['charging_station_op_costs',
        #                  'charging_station_inv_costs',
        #                  'charging_station_emissions',
        #                  'charging_station_revenue']
        
        self.list_var = ['charging_station_op_costs',
                         'charging_station_inv_costs',
                         'charging_station_emissions']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_charging_station_inv_specific_costs = 200000 # investment costs related to a full operational charging station [€]
        self.param_charging_station_selling_price = 0.60 # price of selling energy [€/kWh]
        self.param_charging_station_spec_emissions = 0.05 # specific emissions generated per sold kWh [kgCO2eq/kWh]
        self.param_charging_station_spec_op_costs = 200 # specific operation costs per charging station [€/month]

        self.time_span = control.time_span
        self.reference_date = control.reference_date
        self.path_charts = control.path_charts
        self.name_file = 'input.xlsx'

        #defining paramenters for functions that are not constraints, includiing default values:
        self.dict_parameters  = {'list_sheets':['param_scalars','param_scalars','param_scalars','param_scalars','param_scalars','param_scalars','param_scalars'],
                                 'charging_station_mult': 1.2,
                                 'reference_date': self.reference_date,
                                 'number_hours': self.time_span + 1,
                                 'number_days': math.ceil((self.time_span + 1)/24),
                                 'charging_points' : 2, # number of charging points in the charging station model [# number charging points]
                                 'charging_station_time_step' : 15,
                                 'charging_station_max_power': 100} # timesteps used for building the power demand per charging station [min]
        
        self.dict_series = {'list_sheets':['c_staton_mix_capacity','c_staton_mix_capacity','c_staton_mix_capacity','c_station_SoC','c_station_SoC','c_station_SoC',
                                           'c_station_schedule','c_station_schedule','c_station_schedule'],
                       
                       'list_models':['Tesla MODEL 3','VW E-UP','VW ID.3','Renault ZOE','Smart FORTWO',
                                      'Hyundai KONA 39 kWh','Hyundai KONA 64 kWh','Hyundai IONIQ5 58kWh',
                                      'Hyundai IONIQ5 77kWh','VW ID.4','Fiat 500 E','BMW I3 22kWh','BMW I3 33kWh',
                                      'BMW I3 42kWh','Opel CORSA','MINI Cooper SE','Audi E-TRON','Peugeot 208',
                                      'Renault TWINGO','Opel MOKKA','Nissan LEAF 24 kWh','Nissan LEAF 30 kWh',
                                      'Nissan LEAF 40 kWh','Nissan LEAF 62 kWh','Audi Q4','Dacia SPRING','TESLA MODEL Y',
                                      'VW ID.5','SKODA ENYAQ 77kWh','SKODA ENYAQ 58kWh','CUPRA BORN 77kWh',
                                      'CUPRA BORN 58kWh','RENAULT MEGANE E-TECH','Polestar 2','KIA EV6','Porsche Taycan','Others'],

                        'list_mix_models':[0.065,0.033,0.036,0.036,0.022,0.023,0.023,0.015,0.015,0.021,0.067,
                                           0.009,0.009,0.009,0.019,0.03,0.033,0.013,0.016,0.018,0.002,0.002,
                                           0.002,0.002,0.025,0.021,0.045,0.021,0.018,0.018,0.007,0.007,0.005,
                                           0.012,0.01,0.009,0.282],

                        'list_capacity':[50,36.8,58,52,17.6,39,64,58,77,77,24,22,33,42,45,28.9,85,45,21.3,45,
                                         24,30,40,62,76.6,26.8,75,77,77,58,77,58,60,75,74,83.7,52.631],

                        'list_SoC':[0,0.05,0.1,0.15,0.2,0.25,0.3,0.35,0.4,0.45,0.5,0.55,0.6,0.65,0.7,0.75,0.8,0.85,0.9,0.95,1],

                        'list_mix_start_SoC':[0,0.07,0.1,0.13,0.15,0.17,0.14,0.12,0.08,0.03,0.01,0,0,0,0,0,0,0,0,0,0],

                        'list_mix_end_SoC':[0,0,0,0,0,0,0,0,0,0,0,0,0.03,0.07,0.1,0.11,0.13,0.17,0.15,0.13,0.11],

                        f"{self.name_of_instance}_list_means": [8, 12, 16],
                        f"{self.name_of_instance}_list_std": [0.5, 0.5, 0.5],
                        f"{self.name_of_instance}_list_size": [5, 5, 5] }
        
        self.dict_charging_curves = {'Tesla MODEL 3' : [125 ,140 ,250 ,240 ,225 ,210 ,200 ,180 ,175 ,155 ,125 ,100 ,90 ,75 ,65 ,60 ,45 ,40 ,30 ,20 ,20],
                                     'VW E-UP' : [35 ,35 ,36 ,37 ,37 ,38 ,39 ,36 ,33 ,31 ,30 ,28 ,27 ,25 ,24 ,22.5 ,15 ,13 ,13 ,13 ,13],
                                     'VW ID.3' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,125 ,115 ,105 ,100 ,90 ,75 ,65 ,60 ,55 ,60 ,50 ,40 ,30 ,20],
                                     'Renault ZOE' : [40 ,40 ,42 ,43 ,44 ,44.5 ,44.75 ,45 ,45 ,45.5 ,45.5 ,46 ,40 ,37.5 ,35 ,30 ,27.5 ,26 ,25 ,25 ,25],
                                     'Smart FORTWO' : [41.67 ,42.67 ,43.54 ,44.07 ,44.44 ,44.14 ,42.68 ,42.55 ,42.25 ,41.95 ,41.82 ,41.69 ,41.22 ,37.5 ,36.28 ,34.06 ,31.5 ,27.83 ,23.83 ,19.83 ,15.83],
                                     'Hyundai KONA 39 kWh' : [38 ,38 ,38 ,42 ,43 ,43.5 ,44 ,44 ,44 ,44.5 ,44.5 ,36 ,36 ,36 ,24 ,16 ,16 ,16 ,12 ,10 ,10],
                                     'Hyundai KONA 64 kWh' : [72 ,72 ,72 ,73 ,75 ,75.5 ,76 ,77 ,78 ,71 ,72 ,60 ,57 ,58 ,59 ,37 ,25 ,26 ,21 ,10 ,10],
                                     'VW ID.4' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,125 ,115 ,105 ,100 ,90 ,75 ,65 ,60 ,55 ,60 ,50 ,40 ,30 ,20],
                                     'Fiat 500 E' : [74 ,74 ,74 ,76 ,80 ,80 ,77 ,77 ,68 ,68 ,68 ,65 ,52 ,52 ,52 ,49 ,42 ,47 ,12.5 ,12.5 ,12.5],
                                     'BMW I3 22kWh' : [44 ,44 ,44 ,45 ,47 ,48 ,47 ,44 ,44 ,44 ,44 ,44 ,45 ,40 ,33.5 ,25 ,17.5 ,12.5 ,8 ,5 ,5],
                                     'BMW I3 33kWh' : [45 ,45 ,46 ,44 ,44 ,44 ,44 ,45 ,45 ,45 ,45 ,45 ,45 ,47.5 ,48.5 ,48.5 ,49 ,40 ,27.5 ,10 ,5],
                                     'BMW I3 42kWh' : [47 ,47 ,48 ,46 ,46 ,46 ,46 ,47 ,47 ,47 ,47 ,47 ,47 ,47.5 ,48.5 ,48.5 ,49 ,49 ,35 ,15 ,5],
                                     'Opel CORSA' : [100 ,100 ,100 ,100 ,90 ,95 ,95 ,77 ,77 ,80 ,82 ,85 ,65 ,65 ,55 ,55 ,27.5 ,27.5 ,27.5 ,27.5 ,27.5],
                                     'MINI Cooper SE' : [32 ,32 ,33 ,35 ,42 ,43 ,44 ,45 ,46 ,47 ,48 ,48 ,48.5 ,49 ,47.5 ,48.5 ,45 ,30 ,18 ,12.5 ,12.5],
                                     'Audi E-TRON' : [140 ,140.63 ,141.25 ,141.88 ,142.5 ,143.13 ,143.75 ,144.38 ,145 ,145.63 ,146.25 ,146.88 ,147.5 ,148.13 ,148.75 ,149.38 ,150 ,100 ,75 ,65 ,60],
                                     'Peugeot 208' : [100 ,100 ,100 ,100 ,85 ,88 ,90 ,76 ,76 ,77 ,77 ,77 ,63 ,63 ,52.5 ,52.5 ,27 ,27 ,27 ,27 ,27],
                                     'Renault TWINGO' : [42.5 ,44 ,45.5 ,45 ,45.25 ,44.5 ,42 ,42 ,41.25 ,40.5 ,40 ,39.5 ,38.5 ,38.75 ,38.25 ,36.75 ,35 ,29 ,22.75 ,14 ,11.5],
                                     'Hyundai IONIQ5 77kWh' : [100 ,200 ,220 ,225 ,225 ,230 ,230 ,235 ,235 ,240 ,180 ,185 ,175 ,160 ,145 ,120 ,125 ,100 ,40 ,40 ,40],
                                     'Hyundai IONIQ5 58kWh' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,120 ,118 ,111 ,100 ,85 ,78 ,70 ,65 ,65 ,65 ,50 ,30 ,30 ,25],
                                     'Opel MOKKA' : [100 ,100 ,100 ,100 ,90 ,95 ,98 ,76 ,76.5 ,77 ,77.5 ,78 ,63 ,63 ,57 ,57 ,27 ,27 ,27 ,27 ,27],
                                     'Nissan LEAF 24 kWh' : [40 ,43 ,45 ,46 ,46.5 ,45 ,40 ,39 ,37.5 ,36 ,35 ,34 ,32 ,30 ,28 ,25 ,21 ,18 ,18 ,18 ,18],
                                     'Nissan LEAF 30 kWh' : [42.5 ,42.5 ,42.83 ,43.17 ,43.5 ,43.83 ,44.17 ,44.5 ,44.83 ,45.17 ,45.5 ,45.83 ,46.17 ,46.5 ,46.83 ,47.17 ,47.5 ,42.5 ,32.5 ,22.5 ,12.5],
                                     'Nissan LEAF 40 kWh' : [42.5 ,42.5 ,42.77 ,43.05 ,43.32 ,43.59 ,43.86 ,44.14 ,44.41 ,44.68 ,44.95 ,45.23 ,45.5 ,36 ,34 ,30 ,26 ,23 ,21 ,19 ,17],
                                     'Nissan LEAF 62 kWh' : [42.5 ,42.5 ,42.83 ,43.17 ,43.5 ,43.83 ,44.17 ,44.5 ,44.83 ,45.17 ,45.5 ,45.83 ,46.17 ,46.5 ,46.83 ,42 ,37.5 ,35 ,31 ,27 ,23],
                                     'Audi Q4' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,125 ,120 ,110 ,102 ,94 ,86 ,78 ,70 ,70 ,70 ,50 ,45 ,30 ,25],
                                     'Dacia SPRING' : [31 ,32 ,33 ,34 ,33 ,33 ,32.5 ,32 ,31 ,29 ,29 ,27 ,26 ,24 ,24 ,23 ,21 ,19 ,16 ,15 ,10],
                                     'TESLA MODEL Y' : [220 ,200 ,175 ,162.5 ,150 ,145 ,140 ,120 ,100 ,95 ,90 ,82.5 ,75 ,67.5 ,60 ,55 ,50 ,43.5 ,37 ,33.5 ,30],
                                     'VW ID.5' : [140 ,140 ,140 ,135 ,135 ,135 ,135 ,135 ,135 ,125 ,125 ,125 ,125 ,118.75 ,112.5 ,94.75 ,77 ,50 ,37.5 ,37.5 ,37.5],
                                     'SKODA ENYAQ 58kWh' : [100 ,100 ,100 ,100 ,100 ,100 ,100 ,94 ,88 ,81.5 ,75 ,68.75 ,62.5 ,56.25 ,50 ,37.5 ,25 ,25 ,25 ,15 ,15],
                                     'SKODA ENYAQ 77kWh' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,125 ,112.5 ,106.25 ,100 ,90 ,80 ,70 ,60 ,60 ,60 ,48.75 ,37.5 ,31.25 ,25],
                                     'CUPRA BORN 58kWh' : [100 ,100 ,100 ,100 ,100 ,100 ,100 ,94 ,88 ,81.5 ,75 ,68.75 ,62.5 ,56.25 ,50 ,37.5 ,25 ,25 ,25 ,15 ,15],
                                     'CUPRA BORN 77kWh' : [125 ,125 ,125 ,125 ,125 ,125 ,125 ,125 ,112.5 ,106.25 ,100 ,90 ,80 ,70 ,60 ,60 ,60 ,48.75 ,37.5 ,31.25 ,25],
                                     'RENAULT MEGANE E-TECH' : [125 ,125 ,125 ,118.75 ,112.5 ,108.75 ,105 ,92.5 ,80 ,77.5 ,75 ,70 ,65 ,63.75 ,62.5 ,56.25 ,50 ,50 ,50 ,50 ,50],
                                     'Polestar 2' : [140 ,150 ,150 ,150 ,130 ,132.5 ,135 ,135 ,135 ,112.5 ,90 ,82.5 ,75 ,72.5 ,70 ,70 ,70 ,47.5 ,25 ,25 ,25],
                                     'KIA EV6' : [212.5 ,225 ,225 ,225 ,225 ,225 ,225 ,225 ,225 ,227.5 ,230 ,175 ,175 ,162.5 ,150 ,137.5 ,125 ,82.5 ,40 ,30 ,30],
                                     'Porsche Taycan' : [250 ,250 ,250 ,250 ,250 ,200 ,200 ,150 ,150 ,150 ,150 ,150 ,150 ,150 ,125 ,125 ,50 ,40 ,50 ,50 ,50],
                                     'Others' : [42.5 ,42.5 ,42.83 ,43.17 ,43.5 ,43.83 ,44.17 ,44.5 ,44.83 ,45.17 ,45.5 ,45.83 ,46.17 ,46.5 ,46.83 ,47.17 ,47.5 ,42.5 ,32.5 ,22.5 ,12.5]}
        
        # calling functions to try and read parameter values as soon as class is created
        self.read_parameters(self.dict_parameters,control)
        self.read_series(self.dict_series,control)

        # creating dataframes that are going to be used in non-constraint functions
        self.df_mix_capacity = pd.DataFrame({'Model': self.dict_series['list_models'],
                                             'Mix': self.dict_series['list_mix_models'],
                                             'Max Capacity': self.dict_series['list_capacity']})
        
        self.df_SoC = pd.DataFrame({'SoC':self.dict_series['list_SoC'],
                               'mix initial':self.dict_series['list_mix_start_SoC'],
                               'mix final':self.dict_series['list_mix_end_SoC']})
        
        self.df_data_distribution = pd.DataFrame({'hours of peak': self.dict_series[f"{self.name_of_instance}_list_means"],
                                                  'std of each peak': self.dict_series[f"{self.name_of_instance}_list_std"],
                                                  'number of cars in each peak': self.dict_series[f"{self.name_of_instance}_list_size"]})
        
        self.param_P_to_charging_station = self.charging_demand_calculation(control)

        self.write_P_to_charging_station(control)

    def write_P_to_charging_station(self,control):
        df_power = pd.DataFrame({'param_P_to_' + self.name_of_instance : self.param_P_to_charging_station})
        df_input_series = pd.read_excel(control.path_input + self.name_file ,sheet_name = 'param_series')
        
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            df_input_series = df_input_series.drop('param_P_to_' + self.name_of_instance, axis = 1)
        
        df_input_series = pd.concat([df_input_series, df_power], axis = 1)

        with pd.ExcelWriter(control.path_input + self.name_file, mode = 'a',engine = 'openpyxl', if_sheet_exists = 'replace') as writer:
            df_input_series.to_excel(writer, sheet_name = 'param_series', index = False) 

    def read_parameters(self, parameters,control):
        count = 0
        list_sheets = parameters['list_sheets']
        del parameters['list_sheets']
        for name, default_value in parameters.items():
            df_others = pd.read_excel(control.path_input + self.name_file, index_col=0, sheet_name = list_sheets[count])
            df_others.index.name = None
            try:
                parameters[name] = df_others.loc[name, 'Value']
            except KeyError:
                pass
            count += 1

    def read_series(self, series, control):
        count = 0 
        list_sheets = series['list_sheets']
        del series['list_sheets']
        for name, default_value in series.items():
            df_series = pd.read_excel(control.path_input + self.name_file, sheet_name = list_sheets[count])
            try:
                series[name] = df_series[name].tolist()
            except KeyError:
                pass
            count += 1        

    #function for transforming timestamp to YYYY-MM_DD hh:mm
    def convert_time_stamp(self, day_integer, float_number):
        seconds_integer = int(float_number * 3600)
        delta = timedelta(seconds = seconds_integer, days = day_integer)
        converted_time_stamp = self.dict_parameters['reference_date'] + delta
        converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
        return converted_time_stamp

    #function to return elements of normal distribution
    def generate_normal_distribution(self,mean, standard_deviation, size):
        # Generate random numbers from a normal distribution
        samples = np.random.normal(mean, standard_deviation, size)
        return samples

    def get_max_power(self,car_model,SoC,final_SoC):
        closest_SoC = min(self.dict_series['list_SoC'], key = lambda x: abs(x - SoC))
        closest_index = self.dict_series['list_SoC'].index(closest_SoC)
        power = self.dict_charging_curves[car_model][closest_index]
        capacity = self.get_capacity(car_model)
        energy = power * self.dict_parameters['charging_station_time_step'] / 60

        if (SoC + energy / capacity) > final_SoC:
            power = (final_SoC - SoC) * capacity / (self.dict_parameters['charging_station_time_step'] / 60)

        return power
    
    def get_capacity(self,car_model):
        capacity = self.dict_series['list_capacity'][self.dict_series['list_models'].index(car_model)]
        return capacity

    def update_SoC(self,max_power_per_charging_point,max_charging_power,car_model,SoC,final_SoC):
        capacity = self.get_capacity(car_model)
        charging_energy = min(max_power_per_charging_point, max_charging_power) * self.dict_parameters['charging_station_time_step'] / 60
        SoC_after = min(final_SoC, SoC + charging_energy / capacity)
        actual_power = (SoC_after - SoC) * capacity / (self.dict_parameters['charging_station_time_step'] / 60)

        return SoC_after, actual_power
    
    def managing_queue(self, df_schedule):
        # creating list with time stamps for the power demand calculation if this charging station
        list_time = []
        for i in range(0, self.dict_parameters['number_hours'] * int(60 / self.dict_parameters['charging_station_time_step'])):
            delta = timedelta(minutes = i * self.dict_parameters['charging_station_time_step'])
            converted_time_stamp = self.dict_parameters['reference_date'] + delta
            converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
            list_time.append(converted_time_stamp)
        
        # building structure dataframe for power demand calculation
        df_structured = pd.DataFrame({'time_stamp': list_time})
        for i in range(1,self.dict_parameters['charging_points'] + 1):
            df_structured[f"charging_point_{i}"] = 0
            df_structured[f"SoC_{i}"] = 0
            df_structured[f"initial_SoC_{i}"] = 0
            df_structured[f"final_SoC_{i}"] = 0
            df_structured[f"max_charging_power_{i}"] = 0
            df_structured[f"actual_power_{i}"] = 0
            df_structured[f"status_{i}"] = 'free'
        df_structured[f"total_power_demand"] = 0

        # looping through dataframe and calculating power demand calculation
        for i in range(1, len(df_structured)):

            df_temp = df_schedule[df_schedule['time stamp'].between(list_time[i-1],list_time[i])].reset_index(drop = True)

            # updating of occupation of filled and empty spots
            for j in range(1, self.dict_parameters['charging_points'] + 1):
                # updating charging points that are occupied
                if df_structured[f"status_{j}"].iloc[i-1] == 'occupied':

                    df_structured[f"charging_point_{j}"].iloc[i] = df_structured[f"charging_point_{j}"].iloc[i - 1]
                    df_structured[f"final_SoC_{j}"].iloc[i] = df_structured[f"final_SoC_{j}"].iloc[i - 1]
                    df_structured[f"initial_SoC_{j}"].iloc[i] = df_structured[f"initial_SoC_{j}"].iloc[i - 1]
                    df_structured[f"max_charging_power_{j}"].iloc[i] = self.get_max_power(df_structured[f"charging_point_{j}"].iloc[i],
                                                                                          df_structured[f"SoC_{j}"].iloc[i - 1],
                                                                                          df_structured[f"final_SoC_{j}"].iloc[i])
                    df_structured[f"status_{j}"].iloc[i] = 'occupied'
                
                # if there are free charging points, update table with new cars 
                else:
                    if len(df_temp) > 0:
                        df_structured[f"charging_point_{j}"].iloc[i] = df_temp['car models'].iloc[0]
                        df_structured[f"final_SoC_{j}"].iloc[i] = df_temp['final SoC'].iloc[0]
                        df_structured[f"initial_SoC_{j}"].iloc[i] = df_temp['initial SoC'].iloc[0]
                        df_structured[f"max_charging_power_{j}"].iloc[i] = self.get_max_power(df_structured[f"charging_point_{j}"].iloc[i],
                                                                                            df_structured[f"initial_SoC_{j}"].iloc[i],
                                                                                            df_structured[f"final_SoC_{j}"].iloc[i])
                        df_structured[f"status_{j}"].iloc[i] = 'occupied'

                        df_temp.drop(0, axis = 0, inplace = True)
                        df_temp = df_temp.reset_index(drop = True)
                    
                    else:
                        df_structured[f"charging_point_{j}"].iloc[i] = 0
                        df_structured[f"SoC_{j}"].iloc[i] = 0
                        df_structured[f"initial_SoC_{j}"].iloc[i] = 0
                        df_structured[f"final_SoC_{j}"].iloc[i] = 0
                        df_structured[f"max_charging_power_{j}"].iloc[i] = 0
                        df_structured[f"actual_power_{j}"].iloc[i] = 0
                        df_structured[f"status_{j}"].iloc[i] = 'free' 


            # getting number of occupied charging points
            occupied_charging_points = 0
            for j in range(1, self.dict_parameters['charging_points'] + 1):
                if df_structured[f"status_{j}"].iloc[i] == 'occupied':
                    occupied_charging_points += 1

            # getting max avaliable power for each charging point
            if df_structured.loc[i].filter(like = 'max_charging_power_').sum() > self.dict_parameters['charging_station_max_power']:
                max_power_per_charging_point = self.dict_parameters['charging_station_max_power'] / occupied_charging_points
            else: 
                max_power_per_charging_point = self.dict_parameters['charging_station_max_power']


            # updating state of charge of occupied spots
            for j in range(1, self.dict_parameters['charging_points'] + 1):

                if df_structured[f"status_{j}"].iloc[i] != 'free':
                    if df_structured[f"status_{j}"].iloc[i - 1] == 'occupied':
                        [df_structured[f"SoC_{j}"].iloc[i],
                         df_structured[f"actual_power_{j}"].iloc[i]] = self.update_SoC(max_power_per_charging_point,
                                                                                    df_structured[f"max_charging_power_{j}"].iloc[i],
                                                                                    df_structured[f"charging_point_{j}"].iloc[i],
                                                                                    df_structured[f"SoC_{j}"].iloc[i - 1],
                                                                                    df_structured[f"final_SoC_{j}"].iloc[i])
                        
                        if df_structured[f"SoC_{j}"].iloc[i] == df_structured[f"final_SoC_{j}"].iloc[i]:
                            df_structured[f"status_{j}"].iloc[i] = 'leaving'

                    else:
                        [df_structured[f"SoC_{j}"].iloc[i],
                         df_structured[f"actual_power_{j}"].iloc[i]] = self.update_SoC(max_power_per_charging_point,
                                                                                    df_structured[f"max_charging_power_{j}"].iloc[i],
                                                                                    df_structured[f"charging_point_{j}"].iloc[i],
                                                                                    df_structured[f"initial_SoC_{j}"].iloc[i],
                                                                                    df_structured[f"final_SoC_{j}"].iloc[i])
                        
                        if df_structured[f"SoC_{j}"].iloc[i] == df_structured[f"final_SoC_{j}"].iloc[i]:
                            df_structured[f"status_{j}"].iloc[i] = 'leaving'

                df_structured[f"total_power_demand"].iloc[i] += df_structured[f"actual_power_{j}"].iloc[i]

        return df_structured

    def charging_demand_calculation(self,control):
        #creating list of car models with mix of EVs in germany
        list_models = []
        for i in self.df_mix_capacity.index:
            times = int(self.df_mix_capacity.loc[i,'Mix']*100)
            for j in range(0,times):
                list_models.append(self.df_mix_capacity.loc[i,'Model'])

        #creating list of start and end SoC with input mix
        list_start_SoC = []
        list_end_SoC = []
        for i in self.df_SoC.index:
            times_start_SoC = int(self.df_SoC.loc[i,'mix initial']*100)
            times_final_SoC = int(self.df_SoC.loc[i,'mix final']*100)
            for j in range(0,times_start_SoC):
                list_start_SoC.append(self.df_SoC.loc[i,'SoC'])
            for j in range(0,times_final_SoC):
                list_end_SoC.append(self.df_SoC.loc[i,'SoC'])

        list_days_models = []
        list_hours = []
        list_initial_SoC = []
        list_final_SoC = []

        # generating list with cars and their respective time of arrrival at the station, based on inputs.
        for i in range(0,self.dict_parameters['number_days']):
            for j in self.df_data_distribution.index:

                mean = self.df_data_distribution.loc[j,'hours of peak']
                standard_deviation = self.df_data_distribution.loc[j,'std of each peak']
                size = self.df_data_distribution.loc[j,'number of cars in each peak']

                #getting timestamp that are going to be put together with a car model
                timestamps = self.generate_normal_distribution(mean, standard_deviation, size)

                for k in timestamps:
                    converted_time_stamp = self.convert_time_stamp(i, k)
                    list_hours.append(converted_time_stamp)
                    list_days_models.append(random.choice(list_models))

                    initial_SoC = random.choice(list_start_SoC)
                    final_SoC = random.choice(list_end_SoC)
                    while initial_SoC >= final_SoC:
                        initial_SoC = random.choice(list_start_SoC)
                        final_SoC = random.choice(list_end_SoC)

                    list_initial_SoC.append(initial_SoC)
                    list_final_SoC.append(final_SoC)

        df_schedule = pd.DataFrame({'car models': list_days_models,
                                    'time stamp':list_hours,
                                    'initial SoC':list_initial_SoC,
                                    'final SoC':list_final_SoC})
        
        folder_path = self.path_charts + '/' + self.name_of_instance + '/'
        try:
            os.mkdir(folder_path)
        except FileExistsError:
            print(f"Folder '{self.name_of_instance}' already exists at {folder_path}")
        except Exception as e:
            print(f"An error occurred: {str(e)}")

        df_schedule.to_excel(folder_path + 'df_schedule_' + self.name_of_instance + '.xlsx',index = False)

        df_structured = self.managing_queue(df_schedule)
        df_structured.to_excel(folder_path + 'df_traffic_' + self.name_of_instance + '.xlsx')

        list_total_power = df_structured['total_power_demand'].tolist()
        
        list_total_power_final = []
        x = int(60 / self.dict_parameters['charging_station_time_step'])
        for i in range(0, len(list_total_power), x):
            group = list_total_power[i:i + x]
            list_total_power_final.append(sum(group) / len(group))

        list_total_power_final = list_total_power_final[:control.time_span]

        return [self.dict_parameters['charging_station_mult'] * i for i in list_total_power_final]
        
    def constraint_operation_costs(model,t):
        return model.charging_station_op_costs[t] == model.param_charging_station_spec_op_costs / (30 * 24 / model.time_step)
    
    # def constraint_revenue(model,t):
    #     return model.charging_station_revenue[t] == model.param_P_to_charging_station[t] * model.time_step * model.param_charging_station_selling_price
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.charging_station_inv_costs[t] == model.param_charging_station_inv_specific_costs
        else:
            return model.charging_station_inv_costs[t] == 0
        
    def constraint_emissions(model,t):
        return model.charging_station_emissions[t] == model.param_P_to_charging_station[t] * model.param_charging_station_spec_emissions
    


# other classes 
class control:
    def __init__(self,path_input,name_file):
        self.df = pd.read_excel(path_input + name_file, sheet_name = 'control', index_col = 0)
        self.df.index.name = None

        self.time_span = self.df.loc['time_span','value']
        self.opt_objective = self.df.loc['opt_objective','value']
        self.receding_horizon = self.df.loc['receding_horizon','value']
        self.optimization_horizon = self.df.loc['optimization_horizon','value']
        self.control_horizon = self.df.loc['control_horizon','value']
        self.path_input = self.df.loc['path_input','value']
        self.path_output = self.df.loc['path_output','value']
        self.size_optimization = self.df.loc['size_optimization','value']
        self.reference_date = self.df.loc['reference_date','value']
        self.path_charts = self.df.loc['path_charts','value']

        if self.df.loc['objective','value'] == 'emissions':
            self.opt_equation = 'emission_objective'
        elif self.df.loc['objective','value'] == 'costs':
            self.opt_equation = 'cost_objective'
        else:
            print('\n==========ERROR==========')
            print('Please insert a valid objective for the optimization\n')
            sys.exit()

        if self.df.loc['receding_horizon','value'] == 'yes':
            if self.df.loc['size_optimization','value'] == 'yes':
                print('\n==========ERROR==========')
                print('It is not possible to do a size optimization and receiding horizon simultaneously, Please choose one of the two.\n')
                sys.exit()
            elif self.optimization_horizon > self.time_span:
                print('\n==========ERROR==========')
                print('horizon cannot be bigger than time_span\n')
                sys.exit()
            elif self.control_horizon > self.optimization_horizon:
                print('\n==========ERROR==========')
                print('number of saved lines cannot be bigger than the horizon\n')
                sys.exit()
        else:
            self.horizon = self.time_span

class objective:
    def __init__(self,name_of_instance):
        self.name_of_instance = name_of_instance

        self.list_var = []
        self.list_text_var = []

        self.list_altered_var = []
        self.list_text_altered_var =[]

class net:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                   'electric_source':'yes',
                   'thermal_load':'yes',
                   'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        # self.list_var = ['net_sell_electric',
        #                  'net_buy_electric',
        #                  'net_sell_thermal',
        #                  'net_buy_thermal',
        #                  'net_emissions',
        #                  'net_inv_costs',
        #                  'P_nominal_from_net',
        #                  'Q_nominal_from_net',
        #                  'P_extra_from_net',
        #                  'Q_extra_from_net',
        #                  'net_op_costs',
        #                  'net_revenue']
        
        self.list_var = ['net_buy_electric',
                         'net_buy_thermal',
                         'net_emissions',
                         'net_inv_costs',
                         'P_nominal_from_net',
                         'Q_nominal_from_net',
                         'P_extra_from_net',
                         'Q_extra_from_net',
                         'net_op_costs']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']
        
        self.list_altered_var = []
        self.list_text_altered_var =[]

        #Setting up default values for series if none are given in input file:
        # self.param_net_costs_buy_electric = [0.4] * control.time_span # default price of electric energy bought from network [€/kWh]
        self.param_net_stock_price_electric = [0.3] * control.time_span # default price of electric energy sold from network [€/kWh]
        # self.param_net_costs_buy_thermal = [0.2] * control.time_span # default price of thermal energy bought from network [€/kWh]
        self.param_net_stock_price_thermal = [0.1] * control.time_span # default price of thermal energy sold from network [€/kWh]

        # self.write_net_costs_buy_electric(control)
        self.write_net_stock_price_electric(control)
        # self.write_net_costs_buy_thermal(control)s
        self.write_net_stock_price_thermal(control)

        self.param_net_spec_em_P = 0.56 # kg of CO2 per kWh
        self.param_net_spec_em_Q = 0.24 # kg of CO2 per kWh

        self.param_net_max_P = 1000 # nominal limit of electric power of the network connection [kW] 
        self.param_net_max_Q = 1000 # nominal limit of thermal power of the network connection [kW] 
        self.param_P_net_costs_extra = 1000 # costs of energy that exceeds nominal electric power extracted from the net. This value should be high so that otpimizer does not use it [€/kWh]
        self.param_Q_net_costs_extra = 1000 # costs of energy that exceeds nominal thermal power extracted from the net. This value should be high so that otpimizer does not use it [€/kWh]
        self.param_net_spec_em_P_extra = 1000 # emissions of energy that exceeds nominal electric power extracted from the net. This value should be high so that otpimizer does not use it [kgCO2eq/kWh]
        self.param_net_spec_em_Q_extra = 1000 # emissions of energy that exceeds nominal thermal power extracted from the net. This value should be high so that otpimizer does not use it [kgCO2eq/kWh]

    # def write_net_costs_buy_electric(self,control):
    #     df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
    #     if 'param_' + self.name_of_instance + '_costs_buy_electric' in df_input_series.columns:
    #         pass
    #     else:
    #         df_power = pd.DataFrame({'param_' + self.name_of_instance + '_costs_buy_electric': self.param_net_costs_buy_electric})
    #         df_input_series = pd.concat([df_input_series,df_power], axis =1)
    #         with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
    #             df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    def write_net_stock_price_electric(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_' + self.name_of_instance + '_stock_price_electric' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_stock_price_electric': self.param_net_stock_price_electric})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    # def write_net_costs_buy_thermal(self,control):
    #     df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
    #     if 'param_' + self.name_of_instance + '_costs_buy_thermal' in df_input_series.columns:
    #         pass
    #     else:
    #         df_power = pd.DataFrame({'param_' + self.name_of_instance + '_costs_buy_thermal': self.param_net_costs_buy_thermal})
    #         df_input_series = pd.concat([df_input_series,df_power], axis =1)
    #         with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
    #             df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)

    def write_net_stock_price_thermal(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'input.xlsx',sheet_name = 'param_series')
        if 'param_' + self.name_of_instance + '_stock_price_thermal' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_stock_price_thermal': self.param_net_stock_price_thermal})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'param_series', index = False)


    def constraint_max_P(model,t):
        return model.P_nominal_from_net[t] <= model.param_net_max_P
    
    def constraint_max_Q(model,t):
        return model.Q_nominal_from_net[t] <= model.param_net_max_Q

    def constraint_extra_P(model,t):
        return model.P_extra_from_net[t] >= 0 
    
    def constraint_extra_Q(model,t):
        return model.Q_extra_from_net[t] >= 0 
    
    def constraint_total_P_from(model,t):
        return model.P_from_net[t] == model.P_nominal_from_net[t] + model.P_extra_from_net[t]
    
    def constraint_total_Q_from(model,t):
        return model.Q_from_net[t] == model.Q_nominal_from_net[t] + model.Q_extra_from_net[t]
    
    # def constraint_sell_energy_electric(model,t):
    #     return model.net_sell_electric[t] == model.P_to_net[t] * model.time_step * model.param_net_costs_sell_electric[t]
    
    def constraint_buy_energy_electric(model,t):
        return model.net_buy_electric[t] == (model.P_nominal_from_net[t] * model.time_step * model.param_net_stock_price_electric[t] + 
                                             model.P_extra_from_net[t] * model.time_step * model.param_P_net_costs_extra)
    
    # def constraint_sell_energy_thermal(model,t):
    #     return model.net_sell_thermal[t] == model.Q_to_net[t] * model.time_step * model.param_net_costs_sell_thermal[t]
    
    def constraint_buy_energy_thermal(model,t):
        return model.net_buy_thermal[t] == (model.Q_nominal_from_net[t] * model.time_step * model.param_net_stock_price_thermal[t] +
                                            model.Q_extra_from_net[t] * model.time_step * model.param_Q_net_costs_extra)
    
    def constraint_emissions(model,t):
        return model.net_emissions[t] == (model.P_nominal_from_net[t] * model.param_net_spec_em_P + 
                                          model.Q_nominal_from_net[t] * model.param_net_spec_em_Q +
                                          model.P_extra_from_net[t] * model.param_net_spec_em_P_extra + 
                                          model.Q_extra_from_net[t] * model.param_net_spec_em_Q_extra)
    
    def constraint_investment_costs(model,t):
        return model.net_inv_costs[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.net_op_costs[t] == model.net_buy_electric[t] + model.net_buy_thermal[t]

    # def constraint_revenue(model,t):
    #     return model.net_revenue[t] == model.net_sell_electric[t] + model.net_sell_thermal[t]
