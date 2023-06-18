#region superclasses

class Generator:
    def __init__(self,type_, id,eff,E_in,op_cost,inv_cost,emission):
        self.class_type = 'generator'
        self.type_ = type_
        self.id = id
        self.eff = eff
        self.E_in= E_in
        self.inv_cost = inv_cost
        self.op_cost = op_cost
        self.emission = emission

class pv(Generator):
    # def __init__(self,type_,id_number,eff,E_in,op_cost,inv_cost,emission):
    #     super().__init__('pv',id_number,eff,E_in,op_cost,inv_cost,emission)
    def __init__(self):
        self.list_var = [] #no powers
        self.list_text_var = []
        self.list_param = ['pv_eff','pv_area']
        self.list_text_param = ['','']

    def solar_rule(model,t):
        return model.P_from_pv[t] == model.P_solar[t] * model.pv_eff * model.pv_area
    
class bat:
    def __init__(self):
        self.list_var = ['bat_SOC','bat_K_ch','bat_K_dis'] #no powers
        self.list_text_var = ['within = pyo.NonNegativeReals, bounds=(0, 1)',
                              'domain = pyo.Binary','domain = pyo.Binary',]
        self.list_param = ['bat_starting_SOC','bat_ch_eff','bat_dis_eff','bat_E_max',
                           'bat_c_rate_ch','bat_c_rate_dis']
        self.list_text_param = ['','','','','','']

    def function_rule(model,t):
            if t ==1:
                return model.bat_SOC[t] == model.bat_starting_SOC
            else:
                return model.bat_SOC[t] == model.bat_SOC[t-1] + (model.P_to_bat[t-1] * model.bat_ch_eff 
                                                                 - model.P_from_bat[t-1]/model.bat_dis_eff) * model.time_step / model.bat_E_max

    def charge_limit(model,t):
        return model.P_to_bat[t] <= model.bat_E_max * model.bat_K_ch[t] * model.bat_c_rate_ch
    
    def discharge_limit(model,t):
        return model.P_from_bat[t] <= model.bat_E_max * model.bat_K_dis[t] * model.bat_c_rate_dis
    
    def keys_rule(model,t):
        return model.bat_K_ch[t] + model.bat_K_dis[t] <= 1
    
# class general: # tenho que criar uma regra de conexao pra gerar isso dinamicamente -----------------------
#     def sell_energy(model,t):
#         return model.E_sell[t] == model.P_to_net1[t] * model.time_step

#     def buy_energy(model,t):
#         return model.E_buy[t] == model.P_from_net1[t] * model.time_step

# endregion
# ---------------------------------------------------------------------------------------------------------------------