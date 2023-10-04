import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import math
import sys

class Generator:
    component_type = {'eletric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance, control):
        self.name_of_instance = name_of_instance

        #list with variables used in this class. This list should not contain variables that are used for connecting elements,
        #e.g. P_from_gen, or P_gen_bat. These variables are generated automatically when connections are constructed. 
        self.list_var = ['gen_op_cost',
                         'gen_emissions',
                         'gen_inv_cost'] # this ist contains variables that appear in the constraints of this class. Conection variables
                                         # should not be mentioned here.
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                               'within = pyo.NonNegativeReals',
                               'within = pyo.NonNegativeReals'] # this list contains the type of variable that has to be created.
                                                                # 

        #these lists are used when transforming parameters into variables, when selecting size_optimization This list should stay empty 
        self.list_altered_var = []
        self.list_text_altered_var = []

        #All parameters that are used in the methods of this class must be given am default value
        self.param_gen_eff = 0.10 # unitless
        self.param_gen_size = 50 # size of the generation device. Units can be m2, kW etc. 
        self.param_gen_spec_op_cost = 0.05 # specific operation costs, unit is defined by the rule of accounting costs
        self.param_gen_spec_em = 0.3 # specific emissions, unit is defined by the rule of accounting emissions
        self.param_gen_series = [10] * control.time_span # Units in kWh by default, can be altered if desired
        self.param_gen_spec_inv = 100 # specific emissions, unit is defined by the rule of accounting emissions
        self.param_gen_lifetime = 20 * 8760 # lifetime of the device in hours. After one lifetime, investment cost are repeated

        self.write_gen_series(control) 

    #Every method with a name that start with 'contraint_' will be turned into a constraint of the pyomo model.
    
    #constraint_generation should contain the rule for the generation ef energy of the generation unit. The method below is an example.
    def constraint_generation_rule(model,t):
        return model.P_from_gen[t] == (model.param_gen_series[t] / model.time_step) * model.param_gen_eff * model.param_gen_size
    
    #constraint_operation_costs has the rule for operation cost accounting. The method below is an example.
    def constraint_operation_costs(model,t):
        return model.gen_op_cost[t] == model.param_gen_size * model.param_gen_spec_op_cost 
    
    #constraint_emissions_costs has the rule for emissions accounting. The method below is an example.
    def constraint_emissions(model,t):
        return model.gen_emissions[t] == model.P_from_gen[t] * model.param_gen_spec_em
    
    #constraint_investment_costs has the rule for emissions accounting. In this example, investment is accounted in the first time step 
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.gen_inv_cost[t] == model.param_gen_size * model.param_gen_spec_inv
        else:
            return model.gen_inv_cost[t] == 0
    
    #write_gen_series is a function that insert the time series for param_gen_series. This will be later used as input for the optimization
    #If input sheet "series" already contains data for param_gen_series, the default value is ignored.
    def write_gen_series(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_' + self.name_of_instance + '_series' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_series': self.param_gen_series})
            df_input_series = pd.concat([df_input_series, df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

class pv(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'no'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['pv_op_cost',
                         'pv_emissions',
                         'pv_inv_cost'] 
        
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

        self.param_E_pv_solar = [0.12] * control.time_span # [kWh/m^2] series for solar radiation input, in case no data is given in "series" sheet of input file

        self.write_E_pv_solar(control)

    # write_E_pv_solar writes default solar radiaton to input file in case none is given
    def write_E_pv_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_pv_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def constraint_generation_rule(model,t):
        return model.P_from_pv[t] == (model.param_E_pv_solar[t] / model.time_step) * model.param_pv_eff * model.param_pv_area 

    def constraint_operation_costs(model,t):
        return model.pv_op_cost[t] == model.param_pv_area * model.param_pv_kWp_per_area * (model.param_pv_maintenance + model.param_pv_repair) / (365 * 24 / model.time_step)
    
    def constraint_emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.param_pv_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pv_inv_cost[t] == model.param_pv_area * model.param_pv_kWp_per_area * model.param_pv_inv_per_kWp
        
        elif t % int(model.param_pv_lifetime) == 0:
            return model.pv_inv_cost[t] == model.param_pv_area * model.param_pv_kWp_per_area * model.param_pv_inv_per_kWp
        
        else:
            return model.pv_inv_cost[t] == 0

class solar_th(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'yes'}
    
    super_class = 'generator'

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['solar_th_op_cost',
                         'solar_th_emissions',
                         'solar_th_inv_cost']
        
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

        #default series
        self.param_E_solar_th_solar = [0.12] * control.time_span # kWh/m^2 series for solar irradiation input, in case none is given

        self.write_E_solar_th_solar(control)

    # write_E_solar_th_solar writes default solar radiaton to input file in case none is given
    def write_E_solar_th_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_solar_th_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)
        
    def constraint_generation_rule(model,t):
        return model.Q_from_solar_th[t] == (model.param_E_solar_th_solar[t] * model.time_step) * model.param_solar_th_area * model.param_solar_th_eff
        
    def constraint_operation_costs(model,t):
        return model.solar_th_op_cost[t] == model.param_solar_th_area * (model.param_solar_th_maintenance + model.param_solar_th_repair) / (365 * 24 / model.time_step)
        
    def constraint_emissions(model,t):
        return model.solar_th_emissions[t] == model.Q_from_solar_th[t] * model.param_solar_th_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.solar_th_inv_cost[t] == model.param_solar_th_area * model.param_solar_th_inv_per_area
        
        elif t % int(model.param_solar_th_lifetime) == 0:
            return model.solar_th_inv_cost[t] == model.param_solar_th_area * model.param_solar_th_inv_per_area
        
        else:
            return model.solar_th_inv_cost[t] == 0

class pvt(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['pvt_op_cost',
                         'pvt_emissions',
                         'pvt_inv_cost'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_pvt_P_eff = 0.20 # approximate overall efficiency of pvt cells for electricity generation [-] https://www.solarserver.de/wissen/basiswissen/hocheffizientes-heizsystem-pvt-kollektor-und-waermepumpe/
        self.param_pvt_Q_eff = 0.40 # approximate overall efficiency of pvt cells for thermal energy generation [-] https://www.solarserver.de/wissen/basiswissen/hocheffizientes-heizsystem-pvt-kollektor-und-waermepumpe/
        self.param_pvt_area = 50 # m^2

        self.param_pvt_maintenace = 8.5 # maintenance costs per m^2 of installed device per year [€/m^2/yr] (IEP)
        self.param_pvt_repair = 4.25 # repair costs per m^2 of installed device per year [€/m^2/yr] (IEP)

        self.param_pvt_spec_em = 0 # specific emission values [kgCO2eq/kWh] 
        self.param_pvt_life_time = 20 * 8760 # lifetime of panels in years, same as PV [yr]

        self.param_pvt_inv_per_area = 850 # investments costs in relation to total area of installed device  [€/m^2] (IEP)

        #default series
        self.param_E_pvt_solar = [0.12] * control.time_span # kWh/m^2 series for solar irradiation input, in case none is given

        self.write_E_pvt_solar(control)

    def write_E_pvt_solar(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_E_' + self.name_of_instance + '_solar' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_E_' + self.name_of_instance + '_solar': self.param_E_pvt_solar})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)
        
    def constraint_generation_rule(model,t):
        return model.P_from_pvt[t] == (model.param_E_pvt_solar[t] * model.time_step) * model.param_pvt_area * model.param_pvt_P_eff
    
    def constraint_generation_rule_thermal(model,t):
        return model.Q_from_pvt[t] == (model.param_E_pvt_solar[t] * model.time_step) * model.param_pvt_area * model.param_pvt_Q_eff
        
    def constraint_operation_costs(model,t):
        return model.pvt_op_cost[t] == model.param_pvt_area * (model.param_pvt_maintenace + model.param_pvt_repair) / (365 * 24 /  model.time_step)
        
    def constraint_emissions(model,t):
        return model.pvt_emissions[t] == (model.Q_from_pvt[t] + model.P_from_pvt[t]) * model.param_pvt_spec_em

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pvt_inv_cost[t] == model.param_pvt_area * model.param_pvt_inv_per_area
        
        elif t % int(model.param_pvt_life_time) == 0:
            return model.pvt_inv_cost[t] == model.param_pvt_area * model.param_pvt_inv_per_area
        
        else:
            return model.pvt_inv_cost[t] == 0
        
class CHP(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                      'electric_source':'yes',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['CHP_op_cost',
                         'CHP_emissions',
                         'CHP_inv_cost',
                         'CHP_fuel_cons',
                         'CHP_K']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.Binary']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_P_CHP_max = 20 # max power of CHP device [kW]
        self.param_P_CHP_min = 5 # min operation power of CHP [kW]
        self.param_CHP_P_to_Q_ratio = 0.5 # In german, Stromkennzahl, relation Pel/Pth, [-] https://www.energieatlas.bayern.de/thema_energie/kwk/anlagentypen
        self.param_CHP_eff = 0.8 # Overall efficiency of CHP [-] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden.
        self.param_CHP_fuel_price = 0.09463 # cost per kWh of fuel consumed [€/kWh] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden
        self.param_CHP_spec_em = 0.200 # emissions due to burning of 1 kWh of the fuel [kgCO2eq/kWh] https://www.ris.bka.gv.at/GeltendeFassung.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20008075
        self.param_CHP_inv_cost_per_power = 1700 # investments costs per electric power https://www.heizungsfinder.de/bhkw/kosten-preise/anschaffungskosten
        self.param_CHP_maintenance_costs = 0.04 # costs of maintenance per generated kWh https://partner.mvv.de/blog/welche-bhkw-kosten-fallen-in-der-anschaffung-und-beim-betrieb-an-bhkw#:~:text=Wartung%20und%20Bedienung,75%20Cent%20pro%20kWh%20rechnen.
        self.param_CHP_bonus = 0.09 # KWK-Bonus, compensation for energy from CHP sold to net [€/kWh] https://www.heizungsfinder.de/bhkw/wirtschaftlichkeit/einspeiseverguetung#3     https://www.bhkw-infozentrum.de/wirtschaftlichkeit-bhkw-kwk/ueblicher_preis_bhkw.html
        self.param_CHP_not_used_energy_compensation = 0.01 #  compensaton for decentralised energy generation [€/kWh] https://www.heizungsfinder.de/bhkw/wirtschaftlichkeit/einspeiseverguetung#3
        #VERIFICAR SE ESSE DE CIMA REALMENTE EXISTE, OU SE EH PRA TODOS
        self.param_CHP_lifetime = 20 * 8760 # total lifespan of device in hours [hs]

    def constraint_min_generation(model,t):
        return model.P_from_CHP[t] >= model.param_P_CHP_min * model.CHP_K[t]

    def constraint_max_generation(model,t):
        return model.P_from_CHP[t] <= model.param_P_CHP_max * model.CHP_K[t]

    def constraint_generation_rule(model,t):
        return model.P_from_CHP[t] == model.Q_from_CHP[t] * model.param_CHP_P_to_Q_ratio 
    
    def constraint_fuel_consumption(model,t):
        return model.CHP_fuel_cons[t] == (model.P_from_CHP[t] + model.Q_from_CHP[t]) / model.param_CHP_eff
    
    def constraint_operation_costs(model,t):
        return model.CHP_op_cost[t] == model.CHP_fuel_cons[t] * model.param_CHP_fuel_price + model.CHP_fuel_cons[t] * model.param_CHP_eff * model.param_CHP_maintenance_costs
    
    def constraint_emissions(model,t):
        return model.CHP_emissions[t] == model.CHP_fuel_cons[t] * model.param_CHP_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.CHP_inv_cost[t] == model.param_P_CHP_max * model.param_CHP_inv_cost_per_power
        
        elif t % int(model.param_CHP_lifetime) == 0:
            return model.CHP_inv_cost[t] == model.param_P_CHP_max * model.param_CHP_inv_cost_per_power
        
        else:
            return model.CHP_inv_cost[t] == 0

class gas_boiler(Generator):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'no',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['gas_boiler_fuel_cons',
                         'gas_boiler_op_cost',
                         'gas_boiler_emissions',
                         'gas_boiler_inv_cost'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_Q_gas_boiler_max = 20 # max power that can be generated with this device [kW]
        self.param_Q_gas_boiler_min = 2 # min power limitation when this device is in operation [kW]
        self.param_gas_boiler_eff = 0.95 # efficency when converting fuel into thermal energy [-] 
        self.param_gas_boiler_fuel_cons_ratio = 0.105 #dm3 per kWh of P_from_CHP
        self.param_gas_boiler_fuel_price = 0.09463 # cost per kWh of fuel consumed [€/kWh] https://www.energieheld.de/heizung/bhkw#:~:text=Der%20Gasverbrauch%20bei%20einem%20BHKW,also%20etwa%20bei%20114.000%20Kilowattstunden
        self.param_gas_boiler_spec_em = 0.200 # emissions due to burning of 1 kWh of the fuel [kgCO2eq/kWh] https://www.ris.bka.gv.at/GeltendeFassung.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20008075
        self.param_gas_boiler_maintenance = 1875 #maintenace costs per kW of installed capacity per  year [€/kW/yr] (IEP)
        self.param_gas_boiler_repair = 1875 #repair costs per kW of installed capacity per  year [€/kW/yr] (IEP)
        self.param_gas_boiler_inv_cost_per_power = 125 # investment costs per max kW of installed device [€/kW]
        self.param_gas_boiler_lifetime = 20 * 8760 # total life span of the device [hs]

    def constraint_min_generation(model,t):
        return model.Q_from_gas_boiler[t] >= model.param_Q_gas_boiler_min

    def constraint_max_generation(model,t):
        return model.Q_from_gas_boiler[t] <= model.param_Q_gas_boiler_max 

    def constraint_generation_rule(model,t):
        return model.gas_boiler_fuel_cons[t] == model.Q_from_gas_boiler[t] * model.time_step * model.param_gas_boiler_fuel_cons_ratio
    
    def constraint_generation_rule(model,t):
        return model.Q_from_gas_boiler[t] == model.gas_boiler_fuel_cons[t] * model.param_gas_boiler_eff

    def constraint_operation_costs(model,t):
        return model.gas_boiler_op_cost[t] == (model.gas_boiler_fuel_cons[t] * model.param_gas_boiler_fuel_price +
                                               model.param_Q_gas_boiler_max * (model.param_gas_boiler_maintenance + model.param_gas_boiler_repair) / (365 * 24 / model.time_step))
    
    def constraint_emissions(model,t):
        return model.gas_boiler_emissions[t] == model.gas_boiler_fuel_cons[t] * model.param_gas_boiler_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.gas_boiler_inv_cost[t] == model.param_Q_gas_boiler_max * model.param_gas_boiler_inv_cost_per_power
        
        elif t % int(model.param_gas_boiler_lifetime) == 0:
            return model.gas_boiler_inv_cost[t] == model.param_Q_gas_boiler_max * model.param_gas_boiler_inv_cost_per_power
        
        else:
            return model.gas_boiler_inv_cost[t] == 0



class Transformer:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'no',
                      'thermal_source':'yes'}
    
    def __init__(self, name_of_instance,control):
        self.name_of_instance = name_of_instance
    
        self.list_var = ['trans_emissions',
                         'trans_inv_cost',
                         'trans_op_cost']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_trans_eff = 4
        self.param_trans_spec_op_cost = 0.05
        self.param_trans_inv_specific_costs = 10000
        self.param_trans_spec_em = 0.1

    def constraint_function_rule(model,t):
        return model.Q_from_trans[t] == model.P_to_trans[t] * model.param_trans_eff
    
    def constraint_operation_costs(model,t):
        return model.trans_op_cost[t] == model.param_trans_spec_op_cost * model.time_step

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.trans_inv_cost[t] == model.param_trans_inv_specific_costs
        else:
            return model.trans_inv_cost[t] == 0

    def constraint_emissions(model,t): 
        return model.trans_emissions[t] == model.P_to_trans[t] * model.param_trans_spec_em
    
class heat_pump(Transformer):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'no',
                      'thermal_source':'yes'}

    def __init__(self, name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['heat_pump_emissions',
                         'heat_pump_inv_cost',
                         'heat_pump_op_cost',
                         'heat_pump_K']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.Binary']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_P_heat_pump_max = 20 # max possible electric power consumption [kW]
        self.param_P_heat_pump_min = 0.2 * self.param_P_heat_pump_max # minimum power consumption for device while in operation [kW] 
        self.param_heat_pump_COP = 4 # overall average coefficent of performance, Pth/Pel [-] https://www.heizungsfinder.de/waermepumpe/wirtschaftlichkeit/cop-wert
        self.param_heat_pump_spec_em = 0 # specific emissions per consumed kWh of electricity [kgCO2eq/kWh] https://gshp.org.uk/gshps/what-are-gshps/#:~:text=element%20of%20a%20ground%20source,life%20of%20over%20100%20years.&text=Unlike%20burning%20oil%2C%20gas%2C%20LPG,is%20used%20to%20power%20them).
        self.param_heat_pump_inv_specific_costs = 733 # investment costs per kW of installed capacity (IEP)
        self.param_heat_pump_maintenance = 11 # maintenance costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_repair = 11 # repair costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_operation = 12 # operation costs per installed kW capacity and year [€/kW] (IEP)
        self.param_heat_pump_lifetime = 20 * 8760 # total lifespan of device [hs]


    def constraint_function_rule(model,t):
        return model.Q_from_heat_pump[t] == model.P_to_heat_pump[t] * model.param_heat_pump_COP
    
    def constraint_max_power(model,t):
        return model.P_to_heat_pump[t] <= model.param_P_heat_pump_max * model.heat_pump_K[t]
    
    def constraint_min_power(model,t):
        return model.P_to_heat_pump[t] >= model.param_P_heat_pump_min * model.heat_pump_K[t]
    
    def constraint_operation_costs(model,t):
        return model.heat_pump_op_cost[t] == model.param_P_heat_pump_max * (model.param_heat_pump_maintenance + model.param_heat_pump_operation + model.param_heat_pump_repair) / (365 * 24 / model.time_step)

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.heat_pump_inv_cost[t] == model.param_heat_pump_inv_specific_costs * model.param_P_heat_pump_max
        
        elif t % int(model.param_heat_pump_lifetime) == 0:
            return model.heat_pump_inv_cost[t] == model.param_heat_pump_inv_specific_costs * model.param_P_heat_pump_max

        else:
            return model.heat_pump_inv_cost[t] == 0

    def constraint_emissions(model,t): 
        return model.heat_pump_emissions[t] == model.P_to_heat_pump[t] * model.param_heat_pump_spec_em

class bat(Transformer):
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
                         'bat_op_cost',
                         'bat_emissions',
                         'bat_inv_cost',
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
        
        #default values in case of no input
        self.param_bat_E_max_initial = 100 #max energy capacity of the battery [kWh]
        self.param_bat_starting_SOC = 0.5 # starting state of charge of the battery [-]
        self.param_bat_ch_eff = 0.95 # efficiency for charging the battery [-]
        self.param_bat_dis_eff = 0.95 # efficiency for discharging the battery [-]
        self.param_bat_c_rate_ch = 1 # c-rate of the battery for charging. Specifices the max power of charging in relation to its capacity [1/hr]
        self.param_bat_c_rate_dis = 1 # c-rate of the battery for discharging. Specifices the max power of discharging in relation to its capacity [1/hr]
        self.param_bat_spec_op_cost = 0 # specifies the operation cost of the battery per kWh flow in the battery [€/kWh]
        self.param_bat_spec_em = 50 # embodied emission per kWh capacity of the battery in EU. file: 1-s2.0-S0921344922004402-main [kgCO2eq/kWh]
        self.param_bat_DoD = 0.7 # maximum depth of discharge of the battery [-]
        self.param_bat_inv_per_capacity = 650 # investment costs per installed kWh of capacity. [€/kWh]
        self.param_bat_cycles = 9000 # max number of full cyces before battery is worn out beyond operation [# cycles]
        self.param_bat_lifetime = 10 * 8760 # total lifespan of the device in hours [hs]

        self.param_bat_energy_starting_index = 1 # parameter created to connect energy of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_inv_cost_starting_index = 1 # parameter created to connect investment costs with previous optimization horizon when using receding horizon optimization

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
        

    # linearized equations to define charging power limit. Original equation was: model.P_to_bat[t] <= model.param_bat_E_max_initial * model.bat_K_ch[t] * model.param_bat_c_rate_ch
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


    # linearized equations to define discharging power limit. Original equation was: model.P_from_bat[t] <= model.param_bat_E_max_initial *  model.bat_K_dis[t] * model.param_bat_c_rate_dis  

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
        return model.bat_op_cost[t] == (model.P_to_bat[t] + model.P_from_bat[t]) * model.param_bat_spec_op_cost
    
    #emissions equation accounts for battery embodied CO2 emissions 
    def constraint_emissions(model,t):
        return model.bat_emissions[t] == (model.P_from_bat[t] + model.P_to_bat[t]) * model.param_bat_spec_em / (model.param_bat_cycles*2)
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.bat_inv_cost[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity
        
        elif t % int(model.param_bat_lifetime) == 0:
            return model.bat_inv_cost[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity
        
        else:
            return model.bat_inv_cost[t] == 0

class bat_with_aging(Transformer):
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
                         'bat_with_aging_op_cost',
                         'bat_with_aging_emissions',
                         'bat_with_aging_SOC_max',
                         'bat_with_aging_integer',
                         'bat_with_aging_cumulated_aging',
                         'bat_with_aging_inv_cost'] 
        
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',
                              'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'within = pyo.Integers', 'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_bat_with_aging_E_max_initial = 100 # max stored energy of device at the beginning of lifetime [kWh]
        self.param_bat_with_aging_starting_SOC = 0.5 # initial default state of charge of device [-]
        self.param_bat_with_aging_ch_eff = 0.95 # efficiency for charging device [-]
        self.param_bat_with_aging_dis_eff = 0.95 # efficiency for discharging device [-]
        self.param_bat_with_aging_c_rate_ch = 1 # c-rate of the battery for charging. Specifices the max power of charging in relation to its capacity [1/hr]
        self.param_bat_with_aging_c_rate_dis = 1 # c-rate of the battery for discharging. Specifices the max power of discharging in relation to its capacity [1/hr]
        self.param_bat_with_aging_spec_op_cost = 0.01 # specifies the operation cost of the battery per kWh flow in the battery [€/kWh]
        self.param_bat_with_aging_spec_em = 50 # embodied emission per kWh capacity of the battery in EU. file: 1-s2.0-S0921344922004402-main [kgCO2eq/kWh]
        self.param_bat_with_aging_DoD = 0.7 # max depth of discharge of energy without damaging battery [-]
        self.param_bat_with_aging_final_SoH = 0.7 # max state of charge of battery at the end of lifetime, state of heatlh [-]
        self.param_bat_with_aging_cycles = 9000 # full cycles before final SoH is reached and battery is replaced [# cycles]
        self.param_bat_with_aging_aging = (self.param_bat_with_aging_E_max_initial * (1 -self.param_bat_with_aging_final_SoH)) / (self.param_bat_with_aging_cycles * 2 * self.param_bat_with_aging_E_max_initial) / self.param_bat_with_aging_E_max_initial # state of charge lost for kWh going through the battery.
        self.param_bat_with_aging_inv_per_capacity = 650 # investment costs of battery per kWh of installed capacity [kWh]

        self.param_bat_with_aging_SOC_starting_index = 1 # parameter created to connect state of charge of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_with_aging_cumulated_aging_starting_index = 1 # parameter created to connect accumulated aging of the battery with previous optimization horizon when using receding horizon optimization
        self.param_bat_with_aging_inv_cost_starting_index = 1 # parameter created to connect investment costs of the battery with previous optimization horizon when using receding horizon optimization
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
        return model.bat_with_aging_op_cost[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_spec_op_cost
    
    def constraint_emissions(model,t):
        return model.bat_with_aging_emissions[t] == (model.P_from_bat_with_aging[t] + model.P_to_bat_with_aging[t]) * model.param_bat_with_aging_spec_em/(model.param_bat_with_aging_cycles*2)
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.bat_with_aging_inv_cost[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_inv_per_capacity
        
        elif t == model.starting_index and model.param_bat_with_aging_receding_horizon == 1:
            return model.bat_with_aging_inv_cost[t] == model.param_bat_with_aging_inv_cost_starting_index
        
        else:
            return model.bat_with_aging_inv_cost[t] == model.param_bat_with_aging_E_max_initial * model.param_bat_with_aging_inv_per_capacity * (model.bat_with_aging_integer[t] - model.bat_with_aging_integer[t-1])



class Consumer:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'yes',
                      'thermal_source':'no'}
    
    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance

        self.list_var = ['cons_inv_cost',
                         'cons_op_cost']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_P_to_cons = [500] * control.time_span # this is tranformed into a series in case user does not give input

        self.write_P_to_cons(control)

    def constraint_investment_costs(model,t):
        return model.cons_inv_cost[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.cons_op_cost[t] == 0
    
    def write_P_to_cons(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_P_to_' + self.name_of_instance: self.param_P_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

class demand(Consumer):
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                      'electric_source':'no',
                      'thermal_load':'yes',
                      'thermal_source':'no'}

    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance

        self.list_var = ['demand_inv_cost',
                         'demand_op_cost']
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #Setting up default values for series if none are given in input file:
        self.param_P_to_demand = [500] * control.time_span # default electric power of demand that needs to be covered [kW]
        self.param_Q_to_demand = [1000] * control.time_span # default thermal power of demand that needs to be covered  [kW]]

        self.write_P_to_demand(control)
        self.write_Q_to_demand(control)

    def write_P_to_demand(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_P_to_' + self.name_of_instance: self.param_P_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def write_Q_to_demand(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_Q_to_' + self.name_of_instance in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_Q_to_' + self.name_of_instance: self.param_Q_to_demand})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def constraint_investment_costs(model,t):
        return model.demand_inv_cost[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.demand_op_cost[t] == 0

class charging_station(Consumer):

    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                   'electric_source':'no',
                   'thermal_load':'no',
                   'thermal_source':'no'}
    
    super_class = 'demand'

    def __init__(self, name_of_instance, control):
        self.name_of_instance = name_of_instance

        #default values in case of no input
        self.list_var = ['charging_station_op_cost',
                         'charging_station_inv_cost',
                         'charging_station_emissions']
        
        self.list_text_var = ['within = pyo.NegativeReals',
                              'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        self.param_charging_station_inv_specific_costs = 200000 # investment costs related to a full operational charging station [€]
        self.param_charging_station_selling_price = 0.60 # price of selling energy [€/kWh]
        self.param_charging_station_spec_emissions = 0.05 # specific emissions generated per sold kWh [kgCO2eq/kWh]

        self.time_span = control.time_span
        self.reference_date = control.reference_date
        self.name_file = 'df_input.xlsx'

        #defining paramenters for functions that are not constraints, includiing default values:
        self.dict_parameters  = {'list_sheets':['other','other','other','other'],
                                 'charging_station_mult': 1.2,
                                 'reference_date': self.reference_date,
                                 'number_hours': self.time_span + 1,
                                 'number_days': math.ceil((self.time_span + 1)/24)}
        
        self.dict_series = {'list_sheets':['mix and capacity','mix and capacity','mix and capacity','mix SoC','mix SoC','mix SoC',
                                           'data_charging_station','data_charging_station','data_charging_station'],
                       
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
        df_power = pd.DataFrame({'param_P_to_'+self.name_of_instance : self.param_P_to_charging_station})
        df_input_series = pd.read_excel(control.path_input + self.name_file ,sheet_name = 'series')
        
        if 'param_P_to_' + self.name_of_instance in df_input_series.columns:
            df_input_series = df_input_series.drop('param_P_to_' + self.name_of_instance, axis = 1)
        
        df_input_series = pd.concat([df_input_series, df_power], axis = 1)

        with pd.ExcelWriter(control.path_input + self.name_file, mode = 'a',engine = 'openpyxl', if_sheet_exists = 'replace') as writer:
            df_input_series.to_excel(writer,sheet_name = 'series', index = False) 

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

        for i in range(0,self.dict_parameters['number_days']):
            for j in self.df_data_distribution.index:

                mean = self.df_data_distribution.loc[j,'hours of peak']
                standard_deviation = self.df_data_distribution.loc[j,'std of each peak']
                size = self.df_data_distribution.loc[j,'number of cars in each peak']

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

        df_schedule.to_excel(control.path_output + 'df_schedule_' + self.name_of_instance + '.xlsx',index = False)

        list_time = []

        for i in range(0,self.dict_parameters['number_hours']):
            delta = timedelta(hours = i + 1)
            converted_time_stamp = self.dict_parameters['reference_date'] + delta
            converted_time_stamp = converted_time_stamp.strftime('%Y-%m-%d %H:%M')
            list_time.append(converted_time_stamp)

        list_models = []
        list_start_SoC = []
        list_end_SoC = []

        for i in range(1,len(list_time)):
            df_temp = df_schedule[df_schedule['time stamp'].between(list_time[i-1],list_time[i])]
            list_models.append(df_temp['car models'].to_list())
            list_start_SoC.append(df_temp['initial SoC'].to_list())
            list_end_SoC.append(df_temp['final SoC'].to_list())

        list_time = list_time[:-1]
        df_structured = pd.DataFrame({'timestamp':list_time,
                                    'models':list_models,
                                    'start SoC':list_start_SoC,
                                    'end SoC':list_end_SoC})

        list_power = []
        list_total_power = []
        for i in df_structured.index:
            models  = df_structured['models'].iloc[i]
            start_SoCs = df_structured['start SoC'].iloc[i]
            end_SoCs = df_structured['end SoC'].iloc[i]
            capacity = 0
            list_power_partial = []
            for j in range(0,len(models)):
                model = models[j]
                start_SoC = start_SoCs[j]
                end_SoC = end_SoCs[j]
                capacity = self.df_mix_capacity.loc[self.df_mix_capacity['Model'] == model,'Max Capacity'].values[0]
                power = (end_SoC-start_SoC) * capacity
                list_power_partial.append(power)
            list_power.append(list_power_partial)
            list_total_power.append(sum(list_power_partial))

        df_structured['power'] = list_power
        df_structured['total power'] = list_total_power
        df_structured.to_excel(control.path_output + 'df_structured_' + self.name_of_instance + '.xlsx', index = False)

        lista = df_structured['total power'].tolist()

        return [self.dict_parameters['charging_station_mult'] * i for i in lista]
        
    def constraint_operation_costs(model,t):
        return model.charging_station_op_cost[t] == - model.param_P_to_charging_station[t] * model.time_step * model.param_charging_station_selling_price
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.charging_station_inv_cost[t] == model.param_charging_station_inv_specific_costs
        else:
            return model.charging_station_inv_cost[t] == 0
        
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
            print('==========ERROR==========')
            print('Please insert a valid objective for the optimization')
            sys.exit()

        if self.df.loc['receding_horizon','value'] == 'yes':
            if self.df.loc['size_optimization','value'] == 'yes':
                print('==========ERROR==========')
                print('It is not possible to do a size optimization and receiding horizon simultaneously, Please choose one of the two.')
                sys.exit()
            elif self.optimization_horizon > self.time_span:
                print('==========ERROR==========')
                print('horizon cannot be bigger than time_span')
                sys.exit()
            elif self.control_horizon > self.optimization_horizon:
                print('==========ERROR==========')
                print('number of saved lines cannot be bigger than the horizon')
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

        #defining energy type to build connections with other componets correctly
        self.super_class = 'objective'

class net:
    #defining energy type to build connections with other componets correctly
    component_type = {'electric_load':'yes',
                   'electric_source':'yes',
                   'thermal_load':'yes',
                   'thermal_source':'yes'}

    def __init__(self,name_of_instance,control):
        self.name_of_instance = name_of_instance

        self.list_var = ['net_sell_electric',
                         'net_buy_electric',
                         'net_sell_thermal',
                         'net_buy_thermal',
                         'net_emissions',
                         'net_inv_cost',
                         'P_nominal_from_net',
                         'Q_nominal_from_net',
                         'P_extra_from_net',
                         'Q_extra_from_net'] #no powers
        
        self.list_text_var = ['within = pyo.NonNegativeReals',
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
        self.param_net_cost_buy_electric = [0.4] * control.time_span # default price of electric energy bought from network [€/kWh]
        self.param_net_cost_sell_electric = [0.3] * control.time_span # default price of electric energy sold from network [€/kWh]
        self.param_net_cost_buy_thermal = [0.2] * control.time_span # default price of thermal energy bought from network [€/kWh]
        self.param_net_cost_sell_thermal = [0.1] * control.time_span # default price of thermal energy sold from network [€/kWh]

        self.write_net_cost_buy_electric(control)
        self.write_net_cost_sell_electric(control)
        self.write_net_cost_buy_thermal(control)
        self.write_net_cost_sell_thermal(control)

        self.param_net_spec_em_P = 0.56 # kg of CO2 per kWh
        self.param_net_spec_em_Q = 0.24 # kg of CO2 per kWh

        self.param_net_max_P = 1000 # nominal limit of electric power of the network connection [kW] 
        self.param_net_max_Q = 1000 # nominal limit of thermal power of the network connection [kW] 
        self.param_P_net_cost_extra = 1000 # costs of energy that exceeds nominal electric power extracted from the net. This value should be high so that otpimizer does not use it [€/kWh]
        self.param_Q_net_cost_extra = 1000 # costs of energy that exceeds nominal thermal power extracted from the net. This value should be high so that otpimizer does not use it [€/kWh]
        self.param_net_spec_em_P_extra = 1000 # emissions of energy that exceeds nominal electric power extracted from the net. This value should be high so that otpimizer does not use it [kgCO2eq/kWh]
        self.param_net_spec_em_Q_extra = 1000 # emissions of energy that exceeds nominal thermal power extracted from the net. This value should be high so that otpimizer does not use it [kgCO2eq/kWh]

    def write_net_cost_buy_electric(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_' + self.name_of_instance + '_cost_buy_electric' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_cost_buy_electric': self.param_net_cost_buy_electric})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def write_net_cost_sell_electric(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_' + self.name_of_instance + '_cost_sell_electric' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_cost_sell_electric': self.param_net_cost_sell_electric})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def write_net_cost_buy_thermal(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_' + self.name_of_instance + '_cost_buy_thermal' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_cost_buy_thermal': self.param_net_cost_buy_thermal})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)

    def write_net_cost_sell_thermal(self,control):
        df_input_series  = pd.read_excel(control.path_input + 'df_input.xlsx',sheet_name = 'series')
        if 'param_' + self.name_of_instance + '_cost_sell_thermal' in df_input_series.columns:
            pass
        else:
            df_power = pd.DataFrame({'param_' + self.name_of_instance + '_cost_sell_thermal': self.param_net_cost_sell_thermal})
            df_input_series = pd.concat([df_input_series,df_power], axis =1)
            with pd.ExcelWriter(control.path_input + 'df_input.xlsx', mode = 'a', engine = 'openpyxl', if_sheet_exists= 'replace') as writer:
                df_input_series.to_excel(writer,sheet_name = 'series', index = False)


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
    
    
    def constraint_sell_energy_electric(model,t):
        return model.net_sell_electric[t] == model.P_to_net[t] * model.time_step * model.param_net_cost_sell_electric[t]
    
    def constraint_buy_energy_electric(model,t):
        return model.net_buy_electric[t] == (model.P_nominal_from_net[t] * model.time_step * model.param_net_cost_buy_electric[t] + 
                                             model.P_extra_from_net[t] * model.time_step * model.param_P_net_cost_extra)
    
    def constraint_sell_energy_thermal(model,t):
        return model.net_sell_thermal[t] == model.Q_to_net[t] * model.time_step * model.param_net_cost_sell_thermal[t]
    
    def constraint_buy_energy_thermal(model,t):
        return model.net_buy_thermal[t] == (model.Q_nominal_from_net[t] * model.time_step * model.param_net_cost_buy_thermal[t] +
                                            model.Q_extra_from_net[t] * model.time_step * model.param_Q_net_cost_extra)
    
    def constraint_emissions(model,t):
        return model.net_emissions[t] == (model.P_nominal_from_net[t] * model.param_net_spec_em_P + 
                                          model.Q_nominal_from_net[t] * model.param_net_spec_em_Q +
                                          model.P_extra_from_net[t] * model.param_net_spec_em_P_extra + 
                                          model.Q_extra_from_net[t] * model.param_net_spec_em_Q_extra)
    
    def constraint_investment_costs(model,t):
        return model.net_inv_cost[t] == 0



