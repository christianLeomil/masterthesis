class DEMAND:
    def __init__(self,name_of_instance,time_span):
        self.name_of_instance = name_of_instance

        self.list_var = ['demand_inv_cost','demand_op_cost'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals','within = pyo.NonNegativeReals']

        self.list_altered_var = []
        self.list_text_altered_var =[]

        #default values in case of no input
        self.param_P_to_demand = [500] * time_span # this is tranformed into a series in case user does not give input
        self.param_Q_to_demand = [1000] * time_span # this is tranformed into a series in case user does not give input

        #defining energy type to build connections with other componets correctly
        self.energy_type = {'electric':'yes',
                            'thermal':'yes'}
        
        self.super_class = 'demand'

    def constraint_investment_costs(model,t):
        return model.demand_inv_cost[t] == 0
    
    def constraint_operation_costs(model,t):
        return model.demand_op_cost[t] == 0