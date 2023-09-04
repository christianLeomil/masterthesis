class SOLAR_TH:
    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        self.list_var = ['solar_th_op_cost','solar_th_emissions','solar_th_inv_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals',
                              'within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_solar_th_eff = 0.20 # aproximate overall efficiency of pv cells 
        self.param_solar_th_area = 50 # m^2
        self.param_solar_th_spec_op_cost = 0.02 # cost per hour per m^2 area of pv installed
        self.param_solar_th_spec_em = 0.50 #kgCO2eq/kWh, same value assumed as for PV
        self.param_solar_th_life_time = 15 #lifetime of panels in years
        self.param_solar_th_inv_per_area = 700 #EURO per m2 aperture

        #default series
        self.param_E_solar_th_solar = [0.12] * time_span # kWh/m^2 series for solar irradiation input, in case none is given

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'no',
                            'thermal':'yes'}
        
        self.super_class = 'generator'
        
    def constraint_generation_rule(model,t):
        return model.Q_from_solar_th[t] == (model.param_E_solar_th_solar[t] * model.time_step) * model.param_solar_th_area * model.param_solar_th_eff
        
    def constraint_operation_cost(model,t):
        return model.solar_th_op_cost[t] == model.param_solar_th_area * model.param_solar_th_spec_op_cost
        
    def constraint_emission(model,t):
        return model.solar_th_emissions[t] == model.Q_from_solar_th[t] * model.param_solar_th_spec_em
    
    def constraint_investment_costs(model,t):
        if t == 1:
            return model.solar_th_inv_cost[t] == model.param_solar_th_area * model.param_solar_th_inv_per_area
        else:
            return model.solar_th_inv_cost[t] == 0