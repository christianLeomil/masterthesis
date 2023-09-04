class BAT:
    def __init__(self,name_of_instance,time_step):
        self.name_of_instance = name_of_instance

        self.list_var = ['bat_SOC','bat_K_ch','bat_K_dis','bat_op_cost','bat_emissions','bat_SOC_max',
                         'bat_integer','bat_cumulated_aging','bat_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',
                              'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'within = pyo.Integers', 'within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_bat_E_max_initial = 100
        self.param_bat_starting_SOC = 0.5
        self.param_bat_ch_eff = 0.95
        self.param_bat_dis_eff = 0.95
        self.param_bat_c_rate_ch = 1
        self.param_bat_c_rate_dis = 1
        self.param_bat_spec_op_cost = 0.01
        self.param_bat_spec_em = 50 # kgCO2eq/kWh capacity of the battery in EU
        self.param_bat_DoD = 0.7
        self.param_bat_final_SoH = 0.7
        self.param_bat_cycles = 9000 # full cycles before final SoH is reached and battery is replaced
        self.param_bat_aging = (self.param_bat_E_max_initial * (1 -self.param_bat_final_SoH)) / (self.param_bat_cycles * 2 * self.param_bat_E_max_initial) / self.param_bat_E_max_initial
        self.param_bat_inv_per_capacity = 650 # EURO per kWh capacity

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        
        self.super_class = 'transformer'

    def constraint_depth_of_discharge(model,t):
        return model.bat_SOC[t] >= (1 - model.param_bat_DoD)
    
    def constraint_max_state_of_charge(model,t):
        if t == 1:
            return model.bat_SOC[t] <= 1
        else:
            return model.bat_SOC[t] <= model.bat_SOC_max[t]
    
    def constraint_cumulated_aging(model,t):
        if t == 1:
            return model.bat_cumulated_aging[t] == (model.P_from_bat[t] + 
                                                    model.P_to_bat[t]) * model.time_step * model.param_bat_aging
        else:
            return model.bat_cumulated_aging[t] == model.bat_cumulated_aging[t-1] + (model.P_from_bat[t] + 
                                                                                     model.P_to_bat[t]) * model.time_step * model.param_bat_aging
        
    def constraint_upper_integer_rule(model,t):
        if t == 1:
            return model.bat_integer[t] == 0
        else:
            return model.bat_integer[t] <= model.bat_cumulated_aging[t] / (1-model.param_bat_final_SoH)
    
    def constraint_lower_integer_rule(model,t):
        if t == 1:
            return model.bat_integer[t] == 0 
        else:
            return model.bat_integer[t] >= model.bat_cumulated_aging[t] / (1-model.param_bat_final_SoH) - 1
     
    def constraint_aging(model,t):
        if t == 1:
            return model.bat_SOC_max[t] == 1
        else:
            return model.bat_SOC_max[t] == 1 - model.bat_cumulated_aging[t] + model.bat_integer[t] * (1 - model.param_bat_final_SoH)
    
    def constraint_function_rule(model,t):
            if t == 1:
                return model.bat_SOC[t] == model.param_bat_starting_SOC + (model.P_to_bat[t] * model.param_bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.param_bat_dis_eff) * model.time_step / model.param_bat_E_max_initial
            else:
                return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_to_bat[t] * model.param_bat_ch_eff 
                                                                 - model.P_from_bat[t]/model.param_bat_dis_eff) * model.time_step / model.param_bat_E_max_initial

    def constraint_charge_limit(model,t):
        return model.P_to_bat[t] <= model.param_bat_E_max_initial * model.bat_K_ch[t] * model.param_bat_c_rate_ch

    def constraint_discharge_limit(model,t):
        return model.P_from_bat[t] <= model.param_bat_E_max_initial *  model.bat_K_dis[t] * model.param_bat_c_rate_dis
    
    def constraint_keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1

    def constraint_operation_costs(model,t):
        return model.bat_op_cost[t] == model.param_bat_E_max_initial * model.param_bat_spec_op_cost
    
    def constraint_emissions(model,t):
        return model.bat_emissions[t] == (model.P_from_bat[t] + model.P_to_bat[t]) * model.param_bat_spec_em/(model.param_bat_cycles*2)
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.bat_inv_cost[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity
        else:
            return model.bat_inv_cost[t] == model.param_bat_E_max_initial * model.param_bat_inv_per_capacity * (model.bat_integer[t] - model.bat_integer[t-1])
