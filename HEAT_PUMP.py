class HEAT_PUMP:
    def __init__(self, name_of_instance,time_step):
        self.name_of_instance = name_of_instance

        self.list_var = ['heat_pump_emissions','heat_pump_inv_cost','heat_pump_op_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals',]

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_P_heat_pump_max = 20 #kW electric
        self.param_P_heat_pump_min = 0.3 * self.param_P_heat_pump_max #kW electric
        self.param_heat_pump_COP = 4 #value assumed to be constant 
        self.param_heat_pump_spec_em = 0.138 # kgCO2eq/kWh 
        self.param_heat_pump_inv_specific_costs = 10000 # VERIFICAR
        self.param_heat_pump_spec_op_cost = 1875 * 2 / 8760 * self.param_P_heat_pump_max # EURO per h operation and kW max power

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'transformer'

    def constraint_generation_rule(model,t):
        return model.Q_from_heat_pump[t] == model.P_to_heat_pump[t] * model.param_heat_pump_COP
    
    def constraint_max_power(model,t):
        return model.P_to_heat_pump[t] <= model.param_P_heat_pump_max
    
    def constraint_operation_costs(model,t):
        return model.heat_pump_op_cost[t] == model.param_heat_pump_spec_op_cost * model.time_step

    def constraint_investment_costs(model,t):
        if t == 1:
            return model.heat_pump_inv_cost[t] == model.param_heat_pump_inv_specific_costs
        else:
            return model.heat_pump_inv_cost[t] == 0

    def constraint_emissions(model,t): 
        return model.heat_pump_emissions[t] == model.P_to_heat_pump[t] * model.param_heat_pump_spec_em