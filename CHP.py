class CHP:
    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        self.list_var = ['CHP_fuel_cons','CHP_op_cost','CHP_emissions','CHP_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals','within = pyo.NonNegativeReals'
                              ,'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_P_CHP_max = 20 #W electric
        self.param_P_CHP_min = 0
        self.param_CHP_P_to_Q_ratio = 0.5 
        self.param_CHP_fuel_cons_ratio = 0.105 #dm3 per kWh of P_from_CHP
        self.param_CHP_fuel_price = 5 # EUROS/dm3 of fuel 
        self.param_CHP_spec_em = 2.3 # kg of CO2 emitted per dm3 of fuel (gasoline)
        self.param_CHP_inv_cost_per_power = 1700 # EURO per kW power

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
    
    def constraint_min_generation(model,t):
        return model.P_from_CHP[t] >= model.param_P_CHP_min

    def constraint_max_generation(model,t):
        return model.P_from_CHP[t] <= model.param_P_CHP_max 

    def constraint_generation(model,t):
        return model.P_from_CHP[t] == model.Q_from_CHP[t] * model.param_CHP_P_to_Q_ratio

    def constraint_fuel_consumption(model,t):
        return model.CHP_fuel_cons[t] == model.P_from_CHP[t] * model.time_step * model.param_CHP_fuel_cons_ratio 

    def constraint_operation_costs(model,t):
        return model.CHP_op_cost[t] == model.CHP_fuel_cons[t] * model.param_CHP_fuel_price
    
    def constraint_emissions(model,t):
        return model.CHP_emissions[t] == model.CHP_fuel_cons[t] * model.param_CHP_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.CHP_inv_cost[t] == model.param_P_CHP_max * model.param_CHP_inv_cost_per_power
        else:
            return model.CHP_inv_cost[t] == 0