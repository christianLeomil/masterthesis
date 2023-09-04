class PV:
    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        self.list_var = ['pv_op_cost','pv_emissions','pv_inv_cost'] #no connection powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']
        
        # super().__init__()
        # separar cada classe em um .py ou geradores todos em um. 
        # nome de classe maiusculo e maior.
        # https://pyomo.readthedocs.io/en/stable/working_models.html#changing-the-model-or-data-and-re-solving
        
        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_pv_eff = 0.15 # aproximate overall efficiency of pv cells 
        self.param_pv_area = 100 # m^2
        self.param_pv_spec_op_cost = 0.01 # cost per hour per m^2 area of pv installed
        self.param_pv_kWp_per_area  = 0.12 # kWP per m2 of PV
        self.param_pv_inv_per_kWp = 1000 # EURO per kWp
        self.param_pv_life_time = 30 #lifetime of panels in years
        self.param_pv_spec_em = 0.50 #kgCO2eq/kWh generated

        #default series
        self.param_E_pv_solar = [0.12] * time_span # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'no'}
        self.super_class = 'generator'

    def constraint_generation_rule(model,t):
        return model.P_from_pv[t] == (model.param_E_pv_solar[t] / model.time_step) * model.param_pv_eff * model.param_pv_area 
    
    def constraint_operation_costs(model,t):
        return model.pv_op_cost[t] == model.param_pv_area * model.param_pv_spec_op_cost
    
    def constraint_emissions(model,t):
        return model.pv_emissions[t] == model.P_from_pv[t] * model.param_pv_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.pv_inv_cost[t] == model.param_pv_area * model.param_pv_kWp_per_area * model.param_pv_inv_per_kWp
        else:
            return model.pv_inv_cost[t] == 0