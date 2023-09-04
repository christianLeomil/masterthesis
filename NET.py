class NET:
    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        self.list_var = ['net_sell_electric','net_buy_electric','net_sell_thermal','net_buy_thermal'
                         ,'net_emissions','net_inv_cost'] #no powers
        
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals']
        
        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_net_cost_buy_electric = [0.4] * time_span
        self.param_net_cost_sell_electric = [0.3] * time_span
        self.param_net_cost_buy_thermal = [0.2] * time_span
        self.param_net_cost_sell_thermal = [0.1] * time_span
        self.param_net_spec_em_P = 0.56 # kg of CO2 per kWh
        self.param_net_spec_em_Q = 0.24 # kg of CO2 per kWh

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'external net'
        
    def constraint_sell_energy_electric(model,t):
        return model.net_sell_electric[t] == model.P_to_net[t] * model.time_step * model.param_net_cost_sell_electric[t]
    
    def constraint_buy_energy_electric(model,t):
        return model.net_buy_electric[t] == model.P_from_net[t] * model.time_step * model.param_net_cost_buy_electric[t]
    
    def constraint_sell_energy_thermal(model,t):
        return model.net_sell_thermal[t] == model.Q_to_net[t] * model.time_step * model.param_net_cost_sell_thermal[t]
    
    def constraint_buy_energy_thermal(model,t):
        return model.net_buy_thermal[t] == model.Q_from_net[t] * model.time_step * model.param_net_cost_buy_thermal[t]
    
    def constraint_emissions(model,t):
        return model.net_emissions[t] == model.P_from_net[t] * model.param_net_spec_em_P + model.Q_from_net[t] * model.param_net_spec_em_Q
    
    def constraint_investment_costs(model,t):
        return model.net_inv_cost[t] == 0